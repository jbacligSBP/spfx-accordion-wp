import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { SPComponentLoader } from '@microsoft/sp-loader';

import styles from './Accordion.module.scss';
import * as strings from 'accordionStrings';
import { IAccordionWebPartProps } from './IAccordionWebPartProps';
import MockHttpClient from './MockHttpClient';
import {
  SPHttpClient,
  SPHttpClientResponse   
} from '@microsoft/sp-http';
import {
  Environment,
  EnvironmentType
} from '@microsoft/sp-core-library';

SPComponentLoader.loadCss('https://cdnjs.cloudflare.com/ajax/libs/twitter-bootstrap/3.3.7/css/bootstrap.min.css');
SPComponentLoader.loadScript('https://cdnjs.cloudflare.com/ajax/libs/jquery/3.2.1/jquery.min.js').then(() => {
  SPComponentLoader.loadScript('https://cdnjs.cloudflare.com/ajax/libs/twitter-bootstrap/3.3.7/js/bootstrap.min.js');
});

export interface ISPLists {
  value: ISPList[];
}

export interface ISPList {
  Title: string;
  Content: string;
}

export default class AccordionWebPart extends BaseClientSideWebPart<IAccordionWebPartProps> {

  private id: number;

  private _renderListAsync(): void {
    // Local environment
    if (Environment.type === EnvironmentType.Local) {
      console.log('local environment detected');
      this._getMockListData().then((response) => {
        this._renderList(response.value);
      });
    }
    else if (Environment.type == EnvironmentType.SharePoint || 
              Environment.type == EnvironmentType.ClassicSharePoint) {
      this._getListData()
        .then((response) => {
          this._renderList(response.value);
        })
        .catch((error) => {
          this._renderError(error);
        });
    }
  }

  private _renderList(items: ISPList[]): void {
    let html: string = '', count: number = 0;
    items.forEach((item: ISPList) => {
      html += `
        <div class="panel panel-default">
          <div class="panel-heading" role="tab" id="heading` + count + `">
            <h4 class="panel-title">
              <a 
                role="button" 
                data-toggle="collapse" 
                data-parent="#accordion` + this.id + `" 
                href="#collapse` + count + `" 
                aria-expanded="true" 
                aria-controls="collapse` + count + `"
              >
                ${item.Title}
              </a>
            </h4>
          </div>
          <div 
            id="collapse` + count + `" 
            class="panel-collapse collapse"
            role="tabpanel"
            aria-labelledby="heading` + count + `"
          >
            <div class="panel-body">
              ${item.Content}
            </div>
          </div>
        </div>`;
      count++;
    });

    const listContainer: Element = this.domElement.querySelector('#accordion' + this.id);
    listContainer.innerHTML = html;
  }

  private _renderError(error: any): void {
    let html: string = `<div>List Name is currently empty or the list does not exist.  
      Please update List Name and/or Web URL (if used) in web part settings.</div>`;
    const listContainer: Element = this.domElement.querySelector('#spErrorContainer');
    listContainer.innerHTML = html;
  }

  private _getListData(): Promise<ISPLists> {
    let listName: string = this.properties.listName, webUrl: string = this.properties.webUrl, apiCall: string;
    if(listName) {
      if(!webUrl) { webUrl = this.context.pageContext.web.absoluteUrl; }
      return this.context.spHttpClient.get(webUrl + `/_api/web/lists/GetByTitle('` + listName + `')/Items`, SPHttpClient.configurations.v1)  
      .then((response: SPHttpClientResponse) => {   
        return response.json();  
      });
    } else {
      return Promise.reject(new Error('no list name!'));
    }  
      
  }  

  private _getMockListData(): Promise<ISPLists> {
    return MockHttpClient.get(this.context.pageContext.web.absoluteUrl)
      .then(() => {
        const listData: ISPLists = { 
          value: [
            { Title: 'Heading 1', Content: 'Content 1' },
            { Title: 'Heading 2', Content: 'Content 2' },
            { Title: 'Heading 3', Content: 'Content 3' }
          ]
        };
        return listData;
      }) as Promise<ISPLists>;
  }

  public render(): void {
    this.id = Math.floor(Math.random()*90000) + 10000;
    this.domElement.innerHTML = `
      <div id="spErrorContainer" />
      <div class="panel-group" id="accordion` + this.id + `" role="tablist" aria-multiselectable="true">
      </div>`;
      this._renderListAsync();
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('listName', {
                  label: "List Name",
                  description: "name of list that contains data for this accordion menu"
                }),
                PropertyPaneTextField('webUrl', {
                  label: "Web URL",
                  description: "URL of site where list lives.  Leave blank if list is in the current site."
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
