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
import * as jQuery from 'jquery';

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
        </div>
      </div>
    </div>`;
  }

  public async processFetchedHTML(htmlString: string, pId: string, baseUrl?: string): Promise<void> {
    const parser: DOMParser = new DOMParser();

    // Parse the string into a temporary HTML Document
    const doc: Document = parser.parseFromString(htmlString, 'text/html');

    // Extract scripts, inline styles, and external stylesheet links.
    const scripts: NodeListOf<HTMLScriptElement> = doc.querySelectorAll('script');
    const scriptArray: HTMLScriptElement[] = Array.from(scripts);

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
      // Inject the HTML without scripts, style blocks, or stylesheet links.
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
        const existingLink: HTMLLinkElement | null = Array.from(document.head.querySelectorAll('link[rel~="stylesheet"]'))
          .find((existing: HTMLLinkElement) => existing.getAttribute('href') === resolvedHref) as HTMLLinkElement | null;

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
      // Only execute JavaScript. Leave templates, JSON data blocks etc. alone.
      const type: string = (script.getAttribute('type') || '').trim().toLowerCase();
      const isClassicJs: boolean = type === '' || ModernCewpWebPart._jsMimeTypes.indexOf(type) !== -1;
      const isModule: boolean = type === 'module';
      if (!isClassicJs && !(isModule && script.hasAttribute('src'))) {
        if (isModule) {
          console.warn("Modern CEWP - Inline <script type=\"module\"> is not supported. Move the code to an external .js file.");
        }
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
      fetch(path).then(async (data) => {
        if (data.ok) {
          const content = await data.text();
          this.processFetchedHTML(content, contentLinkPlaceholderId, data.url || path).catch((err: unknown) => {
            console.log("Modern CEWP - Error in processFetchedHTML for contentlink: ", err);
          });
        } else {
          this.domElement.querySelector('#' + contentLinkPlaceholderId).innerHTML = "Content link error: " + String(data.status);
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
        this.domElement.querySelector('#' + contentLinkPlaceholderId).innerHTML = str;
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
