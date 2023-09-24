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
  content: string;
  contentLink: string;
}

interface Window {
  _spPageContextInfo: {};
  jQuery: {};
}

declare let window: Window;

export default class ModernCewpWebPart extends BaseClientSideWebPart<IModernCewpWebPartProps> {

  public _renderEdit(): void {
    let path: string = this.properties.contentLink;
    const hasPath: string = path !== undefined && path !== '' ? strings.Yes : strings.No;
    if (path === '') {
      path = strings.PathNotSet;
    }
    const hasHtml: string = this.properties.content !== undefined && this.properties.content !== '' ? strings.Yes : strings.No;
    const hasLegacyContext: string = this.properties.spPageContextInfo ? strings.Yes : strings.No;
    this.domElement.innerHTML = `
      <div class="${styles.modernCewp}">
        <div class="${styles.container}">
          <div class="${styles.row}">
              <div class="${styles.spjsLink}"><a href='https://spjsblog.com/modern-cewp/' target='_blank'>${strings.Link}</a></div>
              <div class="${styles.title}">${strings.webPartName}</div>
              <div class="${styles.subTitle}">${strings.webPartSettings}</div>
              <p class="${styles.label}">${strings.WebPartHasContentLinkLabel}${hasPath}</p>
              <p class="${styles.label}">${strings.WebPartHasHTMLLabel}${hasHtml}</p>
              <p class="${styles.label}">${strings.WebPartHasPageContextLabel}${hasLegacyContext}</p>
          </div>
        </div>
      </div>`;
  }

  public _renderView(): void {
    // Make jQuery globally available
    if (window.jQuery === undefined) {
      window.jQuery = jQuery;
    }
    // Make _spPageContextInfo available
    if (this.properties.spPageContextInfo && !window._spPageContextInfo) {
      window._spPageContextInfo = this.context.pageContext.legacyPageContext;
    }
    const uid: string = String(Math.random()).substring(2);
    const contentPlaceholderId: string = 'modernCEWP_ContentPlaceholder_' + uid;
    const contentLinkPlaceholderId: string = 'modernCEWP_ContentLinkPlaceholder_' + uid;
    const html: string = this.properties.content;
    const path: string = this.properties.contentLink;
    let innerHTML: string = "";
    if (html !== "") {
      innerHTML += '<div id="' + contentPlaceholderId + '"></div>';
    }
    if (path !== "") {
      innerHTML += '<div id="' + contentLinkPlaceholderId + '"></div>';
    }
    this.domElement.innerHTML = innerHTML;
    if (html !== undefined && html !== "") {
      jQuery('#' + contentPlaceholderId).html(html);
    }
    if (path !== undefined && path !== "") {
      fetch(this.properties.contentLink).then(async (data) => {
        const responseCode = data.status;
        if (responseCode === 200) {
          const content = await data.text();
          jQuery('#' + contentLinkPlaceholderId).html(content);
        } else {
          document.getElementById(contentLinkPlaceholderId).innerHTML = "Content link error: " + String(responseCode);
        }
      }).catch((err) => {
        const str: string = `
        <div class="${styles.modernCewp}">
            <div class="${styles.row}">
              <div class="${styles.title}">${strings.FailedToLoadLabel}</div>
              <div style="margin-bottom:5px;">${this.properties.contentLink}</div>
              <div class="${styles.title}">${strings.ErrorMessageLabel}</div>
              ${err.responseText}
            </div>
        </div>`;
        document.getElementById(contentLinkPlaceholderId).innerHTML = str;
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
    // Detect display mode on classic and modern pages pages
    if (Environment.type === EnvironmentType.ClassicSharePoint) {
      this._renderView();
    } else if (Environment.type === EnvironmentType.SharePoint) {
      if (this.displayMode === DisplayMode.Edit) {
        // Modern SharePoint in Edit Mode
        this._renderEdit();
      } else if (this.displayMode === DisplayMode.Read) {
        // Modern SharePoint in Read Mode
        this._renderView();
      }
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
            description: "SPJSWorks.com/ModernCEWP"
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
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
