import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown,
  IPropertyPaneDropdownOption
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import * as strings from 'SortableAgendaWebPartStrings';
import SortableAgenda from './components/SortableAgenda';
import { ISortableAgendaProps } from './components/ISortableAgendaProps';
import { SPHttpClient, SPHttpClientResponse, SPHttpClientConfiguration } from '@microsoft/sp-http'; 

export interface ISortableAgendaWebPartProps {
  description: string;
  listName:string;
}

export default class SortableAgendaWebPart extends BaseClientSideWebPart <ISortableAgendaWebPartProps> {
  private lists: IPropertyPaneDropdownOption[];
  private listsDropdownDisabled: boolean = true;
  protected onPropertyPaneConfigurationStart(): void {
    this.listsDropdownDisabled = !this.lists;

    if (this.lists) {
      this.render();
      return;
    }

    //this.context.statusRenderer.displayLoadingIndicator(this.domElement, 'lists');
    this.loadLists()
      .then((listOptions: IPropertyPaneDropdownOption[]): void => {
        this.lists = listOptions;
        this.listsDropdownDisabled = false;
        this.context.propertyPane.refresh();
        //this.context.statusRenderer.clearLoadingIndicator(this.domElement);
        this.render();
      });
  }

  public render(): void {
    const element: React.ReactElement<ISortableAgendaProps> = React.createElement(
      SortableAgenda,
      {
        listName: "CalendarList"
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  private loadLists(): Promise<IPropertyPaneDropdownOption[]> {
    return new Promise<IPropertyPaneDropdownOption[]>((resolve: (options: IPropertyPaneDropdownOption[]) => void, reject: (error: any) => void) => {
      this.context.spHttpClient.get(`${this.context.pageContext.web.absoluteUrl}/_api/lists?$filter=BaseTemplate eq 106`, SPHttpClient.configurations.v1, {
        headers: {
          'odata-version': '3.0',
          'accept': 'application/json;odata=verbose',
          'content-type': 'application/json;odata=verbose'
        }
      }).then((res: SPHttpClientResponse) => {
        console.log("www");
       
        var mappedArray = [];
        resolve(mappedArray);
        
      }).catch(error => {
      
      });
      /*setTimeout((): void => {
        resolve([{
          key: 'sharedDocuments',
          text: 'Shared Documents'
        },
        {
          key: 'myDocuments',
          text: 'My Documents'
        }]);
      }, 2000);*/
    });
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
                PropertyPaneDropdown('listName', {
                  label: strings.ListNameFieldLabel,
                  options: this.lists,
                  disabled: this.listsDropdownDisabled
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
