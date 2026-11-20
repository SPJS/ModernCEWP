import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
  Environment,
  EnvironmentType,
  DisplayMode
} from '@microsoft/sp-core-library';

import styles from './ModernCewpWebPart.module.scss';
import * as strings from 'ModernCewpWebPartStrings';
import jQuery from 'jquery';

export interface IModernCewpWebPartProps {
  spPageContextInfo: boolean;
  cewpTitle: string;
  content: string;
  contentLink: string;
}

interface Window {
  _spPageContextInfo: {};
  jQuery: {};
  $: {};
}

declare let window: Window;

export default class ModernCewpWebPart extends BaseClientSideWebPart<IModernCewpWebPartProps> {

  private static readonly _jsMimeTypes: string[] = [
    'text/javascript', 'application/javascript', 'application/x-javascript',
    'text/ecmascript', 'application/ecmascript', 'text/jscript'
  ];

  private static _escapeHtml(value: string): string {
    return value.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  }

  // True for script elements that the web part executes. Other script blocks
  // (templates, JSON data etc.) are left in the HTML for the user's code to read.
  private static _isExecutableScript(script: HTMLScriptElement): boolean {
    const type: string = (script.getAttribute('type') || '').trim().toLowerCase();
    return type === '' || type === 'module' || ModernCewpWebPart._jsMimeTypes.indexOf(type) !== -1;
  }

  // Finds markup that SharePoint's Content Security Policy blocks: inline event
  // handler attributes (onclick etc.) and javascript: URLs. Returns short
  // descriptions such as "<button onclick>" (without duplicates).
  private static _findCspBlockedMarkup(root: ParentNode): string[] {
    const found: string[] = [];
    Array.from(root.querySelectorAll('*')).forEach((el: Element) => {
      Array.from(el.attributes).forEach((attr: Attr) => {
        const name: string = attr.name.toLowerCase();
        const isHandler: boolean = name.indexOf('on') === 0 && name.length > 2;
        const isJsUrl: boolean = ['href', 'src', 'action', 'formaction'].indexOf(name) !== -1
          && /^\s*javascript:/i.test(attr.value);
        if (isHandler || isJsUrl) {
          const desc: string = `<${el.tagName.toLowerCase()} ${isJsUrl ? name + '="javascript:..."' : name}>`;
          if (found.indexOf(desc) === -1) {
            found.push(desc);
          }
        }
      });
    });
    return found;
  }

  private static _warnCspBlockedMarkup(found: string[], source: string): void {
    if (found.length > 0) {
      console.warn(`Modern CEWP - ${source} contains markup that is blocked by SharePoint's Content Security Policy and will not work: ${found.join(', ')}. Use addEventListener in a script instead.`);
    }
  }

  private get _version(): string {
    return this.context.manifest.version;
  }

  public _renderEdit(): void {
    const hasPath: string = this.properties.contentLink ? strings.Yes : strings.No;
    const hasHtml: string = this.properties.content !== undefined && this.properties.content !== '' ? strings.Yes : strings.No;
    const hasLegacyContext: string = this.properties.spPageContextInfo ? strings.Yes : strings.No;
    this.domElement.innerHTML = `
    <div class="${styles.modernCewp}">
      <div class="${styles.container}">
        <div class="${styles.header}">${ModernCewpWebPart._escapeHtml(this.properties.cewpTitle || "")}</div>
        <div class="${styles.row}">
          <div class="${styles.spjsLink}"><a href='https://spjsblog.com/modern-cewp/' target='_blank'>${strings.Link}</a></div>
          <div class="${styles.title}">${strings.webPartName}</div>
          <div class="${styles.subTitle}">${strings.webPartSettings}</div>
          <p class="${styles.label}">${strings.WebPartHasContentLinkLabel}${hasPath}</p>
          <p class="${styles.label}">${strings.WebPartHasHTMLLabel}${hasHtml}</p>
          <p class="${styles.label}">${strings.WebPartHasPageContextLabel}${hasLegacyContext}</p>
          <p class="${styles.label}">${strings.VersionLabel}${this._version}</p>
          <div class="cewpCspWarnings"></div>
        </div>
      </div>
    </div>`;
    this._showCspWarningsInEditMode().catch((err: unknown) => {
      console.log("Modern CEWP - Could not check content for CSP issues: ", err);
    });
  }

  // Lets editors see markup that will not work because of CSP, without opening the console.
  private async _showCspWarningsInEditMode(): Promise<void> {
    const parser: DOMParser = new DOMParser();
    const warnings: string[] = [];

    const html: string = this.properties.content || "";
    if (html !== "") {
      const found: string[] = ModernCewpWebPart._findCspBlockedMarkup(parser.parseFromString(html, 'text/html'));
      if (found.length > 0) {
        warnings.push(`${strings.WebPartHasHTMLLabel}${found.join(', ')}`);
      }
    }

    const path: string = (this.properties.contentLink || "").trim();
    if (path !== "") {
      try {
        const response: Response = await fetch(path);
        if (response.ok) {
          const found: string[] = ModernCewpWebPart._findCspBlockedMarkup(parser.parseFromString(await response.text(), 'text/html'));
          if (found.length > 0) {
            warnings.push(`${strings.WebPartHasContentLinkLabel}${found.join(', ')}`);
          }
        }
      } catch {
        // Load errors are shown in read mode.
      }
    }

    const target: HTMLElement | null = this.domElement.querySelector('.cewpCspWarnings');
    if (target && warnings.length > 0) {
      target.innerHTML = `
        <p class="${styles.label} ${styles.warning}">${strings.CspWarningLabel}</p>
        ${warnings.map((w: string) => `<p class="${styles.label} ${styles.warning}">${ModernCewpWebPart._escapeHtml(w)}</p>`).join('')}`;
    }
  }

  public async processFetchedHTML(htmlString: string, pId: string, baseUrl?: string): Promise<void> {
    const parser: DOMParser = new DOMParser();

    // Parse the string into a temporary HTML Document
    const doc: Document = parser.parseFromString(htmlString, 'text/html');

    ModernCewpWebPart._warnCspBlockedMarkup(
      ModernCewpWebPart._findCspBlockedMarkup(doc),
      baseUrl ? `The content link ${baseUrl}` : 'The HTML code'
    );

    // Extract the scripts that will be executed. Non-JavaScript script blocks stay in the HTML.
    const allScripts: HTMLScriptElement[] = Array.from(doc.querySelectorAll('script'));
    const scriptArray: HTMLScriptElement[] = allScripts.filter(ModernCewpWebPart._isExecutableScript);

    // Only doc.body is injected, so move non-JavaScript script blocks from <head> into the body.
    allScripts
      .filter((script: HTMLScriptElement) => !ModernCewpWebPart._isExecutableScript(script) && doc.head.contains(script))
      .forEach((script: HTMLScriptElement) => doc.body.insertBefore(script, doc.body.firstChild));

    const styles: NodeListOf<HTMLStyleElement> = doc.querySelectorAll('style');
    const styleArray: HTMLStyleElement[] = Array.from(styles);

    const links: NodeListOf<HTMLLinkElement> = doc.querySelectorAll('link');
    const stylesheetLinks: HTMLLinkElement[] = Array.from(links).filter((link: HTMLLinkElement) => {
      const rel: string = link.getAttribute('rel') || '';
      return rel.toLowerCase().split(/\s+/).indexOf('stylesheet') !== -1;
    });

    // Remove extracted elements before injecting doc.body so they are not duplicated.
    scriptArray.forEach((script: HTMLScriptElement) => script.remove());
    styleArray.forEach((styleTag: HTMLStyleElement) => styleTag.remove());
    stylesheetLinks.forEach((linkTag: HTMLLinkElement) => linkTag.remove());

    // Find your target container
    const container: HTMLElement | null = this.domElement.querySelector('#' + pId);

    if (container) {
      // Inject the HTML without executable scripts, style blocks, or stylesheet links.
      // Script blocks left in the HTML (templates, JSON) are not executed by innerHTML.
      container.innerHTML = doc.body.innerHTML;

      // Inject inline styles into the container.
      styleArray.forEach((styleTag: HTMLStyleElement) => {
        container.appendChild(styleTag);
      });

      // Inject external stylesheets into document.head.
      stylesheetLinks.forEach((linkTag: HTMLLinkElement) => {
        const href: string | null = linkTag.getAttribute('href');
        if (!href) {
          return;
        }

        const resolvedHref: string = baseUrl ? new URL(href, baseUrl).href : href;
        const existingLink: HTMLLinkElement | undefined = Array.from(document.head.querySelectorAll<HTMLLinkElement>('link[rel~="stylesheet"]'))
          .find((existing: HTMLLinkElement) => existing.getAttribute('href') === resolvedHref);

        if (!existingLink) {
          const newLink: HTMLLinkElement = document.createElement('link');
          Array.from(linkTag.attributes).forEach((attr: Attr) => {
            newLink.setAttribute(attr.name, attr.name.toLowerCase() === 'href' ? resolvedHref : attr.value);
          });
          document.head.appendChild(newLink);
        }
      });

      // Pass the extracted script elements to the ordered handler
      await this.handleScriptsOrdered(scriptArray, baseUrl);
    } else {
      console.error("Modern CEWP - Target container not found in the DOM.");
    }
  }

  public async handleScriptsOrdered(scripts: NodeListOf<HTMLScriptElement> | HTMLScriptElement[], baseUrl?: string): Promise<void> {
    for (const script of Array.from(scripts)) {
      if (!ModernCewpWebPart._isExecutableScript(script)) {
        continue;
      }
      const isModule: boolean = (script.getAttribute('type') || '').trim().toLowerCase() === 'module';
      if (isModule && !script.hasAttribute('src')) {
        console.warn("Modern CEWP - Inline <script type=\"module\"> is not supported. Move the code to an external .js file.");
        continue;
      }
      if (script.hasAttribute('src')) {
        // Handle External Scripts
        // A script that fails to load (e.g. blocked by CSP because the domain is not a
        // trusted script source) is logged, and the remaining scripts still run.
        await new Promise<void>((resolve) => {
          const newScript: HTMLScriptElement = document.createElement('script');

          // Copy all attributes, resolving relative script URLs against the HTML file.
          Array.from(script.attributes).forEach((attr: Attr) => {
            const value: string = attr.name.toLowerCase() === 'src' && baseUrl
              ? new URL(attr.value, baseUrl).href
              : attr.value;
            newScript.setAttribute(attr.name, value);
          });

          newScript.onload = () => resolve();
          newScript.onerror = () => {
            console.error(`Modern CEWP - Failed to load script: ${newScript.src}. If it is hosted outside this tenant, check that the domain is added to Trusted script sources in the SharePoint admin center.`);
            resolve();
          };

          document.head.appendChild(newScript);
        });
      } else if (script.textContent?.trim()) {
        // Handle Inline Scripts
        try {
          // Indirect eval runs the code in the global scope, like a regular <script> tag,
          // so top-level functions and variables are available to other scripts.
          // eslint-disable-next-line no-eval
          (0, eval)(script.textContent);
        } catch (err) {
          console.error("Modern CEWP - Inline script error:", err);
        }
      }
    }
  }

  public _renderView(): void {
    // Make jQuery globally available
    if (window.jQuery === undefined) {
      window.jQuery = jQuery;
    }
    if (window.$ === undefined) {
      window.$ = window.jQuery;
    }
    // Make _spPageContextInfo available in the global window scope
    if (this.properties.spPageContextInfo && !window._spPageContextInfo) {
      window._spPageContextInfo = this.context.pageContext.legacyPageContext;
    }
    const uid: string = String(Math.random()).substring(2);
    const contentPlaceholderId: string = 'modernCEWP_ContentPlaceholder_' + uid;
    const contentLinkPlaceholderId: string = 'modernCEWP_ContentLinkPlaceholder_' + uid;
    // Properties are undefined on a newly added web part
    const html: string = this.properties.content || "";
    const path: string = (this.properties.contentLink || "").trim();
    let innerHTML: string = "";
    if (html !== "") {
      innerHTML += '<div id="' + contentPlaceholderId + '"></div>';
    }
    if (path !== "") {
      innerHTML += '<div id="' + contentLinkPlaceholderId + '"></div>';
    }
    this.domElement.innerHTML = innerHTML;
    if (html !== "") {
      this.processFetchedHTML(html, contentPlaceholderId).catch((err: unknown) => {
        console.log("Modern CEWP - Error in processFetchedHTML: ", err);
      });
    }
    if (path !== "") {
      const showLinkMessage = (messageHtml: string): void => {
        const placeholder: HTMLElement | null = this.domElement.querySelector('#' + contentLinkPlaceholderId);
        if (placeholder) {
          placeholder.innerHTML = messageHtml;
        }
      };
      fetch(path).then(async (data) => {
        if (data.ok) {
          const content = await data.text();
          this.processFetchedHTML(content, contentLinkPlaceholderId, data.url || path).catch((err: unknown) => {
            console.log("Modern CEWP - Error in processFetchedHTML for contentlink: ", err);
          });
        } else {
          showLinkMessage("Content link error: " + String(data.status));
        }
      }).catch((err: Error) => {
        const str: string = `
        <div class="${styles.modernCewp}">
            <div class="${styles.row}">
              <div class="${styles.title}">${strings.FailedToLoadLabel}</div>
              <div style="margin-bottom:5px;">${ModernCewpWebPart._escapeHtml(path)}</div>
              <div class="${styles.title}">${strings.ErrorMessageLabel}</div>
              ${ModernCewpWebPart._escapeHtml(err && err.message ? err.message : String(err))}
            </div>
        </div>`;
        showLinkMessage(str);
      });
    }
    if (path === "" && html === "") {
      const str: string = `
        <div class="${styles.modernCewp}">
          <div class="${styles.container}">
            <div class="${styles.row}">
              <div class="${styles.title}">${strings.DispModeEmpty}</div>
            </div>
          </div>
        </div>`;
      this.domElement.innerHTML = str;
    }
  }

  public render(): void {
    // Classic pages always render the content. Modern pages, Teams and the local
    // workbench show the settings summary in edit mode and the content in read mode.
    if (Environment.type !== EnvironmentType.ClassicSharePoint && this.displayMode === DisplayMode.Edit) {
      this._renderEdit();
    } else {
      this._renderView();
    }
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: `SPJSWorks.com/ModernCEWP - ${strings.VersionLabel}${this._version}`
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('cewpTitle', {
                  label: strings.CEWPTitle,
                  rows: 1,
                  multiline: false,
                  resizable: false
                }),
                PropertyPaneTextField('contentLink', {
                  label: strings.ContentlinkFieldLabel,
                  multiline: true,
                  rows: 2,
                  resizable: true
                }),
                PropertyPaneTextField('content', {
                  label: strings.ContentFieldLabel,
                  multiline: true,
                  rows: 20,
                  resizable: true
                }),
                PropertyPaneToggle('spPageContextInfo', {
                  label: strings.AddspPageContextInfo,
                  checked: this.properties.spPageContextInfo,
                  onText: 'Enabled',
                  offText: 'Disabled'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
