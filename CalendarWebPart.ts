import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown,
  IPropertyPaneDropdownOption,
  IPropertyPaneGroup
} from '@microsoft/sp-webpart-base';

import { SPComponentLoader } from '@microsoft/sp-loader';

import * as strings from 'CalendarWebPartStrings';
import Calendar from './components/Calendar';
import { ICalendarProps } from './components/ICalendarProps';

import { PropertyFieldCollectionData, CustomCollectionFieldType } from '@pnp/spfx-property-controls/lib/PropertyFieldCollectionData';

import {  
  SPHttpClient,
  SPHttpClientResponse,
  IDigestCache, DigestCache 
} from '@microsoft/sp-http';

export interface ICalendarWebPartProps {
  siteUrl: string;  
  listName: string;
  listView: string;
  listColumn: string;
  viewCollection: any[];
}

export default class CalendarWebPart extends BaseClientSideWebPart<ICalendarWebPartProps> {

  private lists: IPropertyPaneDropdownOption[] = [];
  private listViews:IPropertyPaneDropdownOption[] = [];
  private listColumns: IPropertyPaneDropdownOption[] = [];

  public render(): void {
    
    const element: React.ReactElement<ICalendarProps > = React.createElement(
      Calendar,
      {
        spHttpClient: this.context.spHttpClient,
        siteUrl: this.properties.siteUrl,
        listName: this.properties.listName,        
        viewCollection: this.properties.viewCollection,
        showPanel: false,
        dom:this.domElement,
        statusRender:this.context.statusRenderer
      }
    );

    ReactDom.render(element, this.domElement);
  }

  private async loadLists(): Promise<IPropertyPaneDropdownOption[]>
  {
    const _lists: IPropertyPaneDropdownOption[] = [];    
    try {
    const results = await this.context.spHttpClient.get(this.properties.siteUrl + `/_api/web/lists?$select=Id,Title&$filter=(BaseTemplate eq 106)`, SPHttpClient.configurations.v1)  
        .then((response: SPHttpClientResponse) => {
          return response.json();  
        });
        for (const list of results.value) {
          _lists.push({ key: list.Title, text: list.Title });
        }
      } catch (error) {    
      }
      return _lists; 
  }

  private async loadViews(): Promise<IPropertyPaneDropdownOption[]> {
    
    const _listViews: IPropertyPaneDropdownOption[] = [];   
    try {
      const results = await this.context.spHttpClient.get(this.properties.siteUrl + `/_api/web/lists/GetbyTitle('`+this.properties.listName+`')/views`, SPHttpClient.configurations.v1)  
      .then((response: SPHttpClientResponse) => {
        return response.json();  
      });
      for (const listView of results.value) {
        if(listView.Title!="")
        {
          _listViews.push({ key: listView.Title, text: listView.Title });
        }
      }
    } catch (error) {
    }
    return _listViews; 
  }

  private async loadListColumns(): Promise<IPropertyPaneDropdownOption[]> {
    
    const _listColumns: IPropertyPaneDropdownOption[] = [];   
    try {
      const results = await this.context.spHttpClient.get(this.properties.siteUrl + `/_api/web/lists/GetbyTitle('`+this.properties.listName+`')/Fields`, SPHttpClient.configurations.v1)  
      .then((response: SPHttpClientResponse) => {
        return response.json();  
      });
      for (const listColumn of results.value) {
        if(listColumn.Title!="")
        {
          _listColumns.push({ key: listColumn.Title, text: listColumn.Title });
        }
      }
    } catch (error) {
    }
    return _listColumns; 
  }

  public  async onInit(): Promise<any> {    
    console.log("OnInit");
    SPComponentLoader.loadCss('https://cdnjs.cloudflare.com/ajax/libs/fullcalendar/3.4.0/fullcalendar.min.css');
    this.properties.siteUrl = this.properties.siteUrl; //? this.properties.siteUrl : this.context.pageContext.site.absoluteUrl;
    
    const _lists = await this.loadLists();
    this.lists = _lists;
    this.properties.listName = this.lists.length > 0 ? this.lists[0].key.toString() : undefined;
    
    const _listViews=await this.loadViews();
    this.listViews =_listViews;   
    
    const _listColumns= await this.loadListColumns();
    this.listColumns= _listColumns;     
    return Promise.resolve();
  }

  protected async onPropertyPaneConfigurationStart() {
    console.log("onPropertyPaneConfigurationStart");
    try {
      if (this.properties.siteUrl) {
        const _lists = await this.loadLists();
        this.lists = _lists;
        if(this.properties.listName)
        {
          const _listViews = await this.loadViews();
          this.listViews = _listViews;          
          const _listColumns = await this.loadListColumns();
          this.listColumns = _listColumns;
        }
        this.context.propertyPane.refresh();

      } else {
        this.lists = [];
        this.listViews = [];
        this.listColumns =[];
        this.properties.listName = '';
        this.properties.listView = '';
        this.properties.listColumn = '';
        this.context.propertyPane.refresh();
      }
    } catch (error) {
    }
  }

  protected async onPropertyPaneFieldChanged(propertyPath: string, oldValue: string, newValue: string) {    
    try {    
      this.context.propertyPane.refresh();

      if (propertyPath === 'siteUrl' && newValue) {      
        super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
        const _oldValue = this.properties.listName;
        this.onPropertyPaneFieldChanged('list', _oldValue, this.properties.listName);
        this.context.propertyPane.refresh();
        const _lists = await this.loadLists();        
        this.lists = _lists;
        this.properties.listName = this.lists.length > 0 ? this.lists[0].key.toString() : undefined;
        this.context.propertyPane.refresh();
        this.render();
      }
      else if(propertyPath === 'list' && newValue) {  
        this.properties.viewCollection=[];     
        super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
        const _oldValue = this.properties.listView;
        this.onPropertyPaneFieldChanged('viewCollection', _oldValue, this.properties.listView);
        this.context.propertyPane.refresh();
        const _listViews = await this.loadViews();
        const _listColumns = await this.loadListColumns();
        this.listViews = _listViews;
        this.listColumns = _listColumns;       
        this.context.propertyPane.refresh();
        this.render();
      }
      
      else {        
        super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
      }
    } catch (error) {
    }
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
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
                PropertyPaneTextField('siteUrl', {
                  label: strings.SiteUrlFieldLabel,                  
                  value: this.properties.siteUrl,
                  deferredValidationTime: 1200,
                }),
                PropertyPaneDropdown('list', {
                  label: strings.ListFieldLabel,
                  options: this.lists,
                }),
                PropertyFieldCollectionData("viewCollection", {
                  key: "viewCollection",
                  label: "",
                  panelHeader: "",
                  manageBtnLabel: "Overlay properties",
                  value: this.properties.viewCollection,
                  fields: [
                    {
                      id: "View",
                      title: "View",
                      type: CustomCollectionFieldType.dropdown,
                      options: this.listViews,
                      required: true                                               
                    },
                    {
                      id: "Color",
                      title: "Color",
                      type: CustomCollectionFieldType.string,
                      required: true
                    },
                    {
                      id:"Column",
                      title:"Column to display",
                      type: CustomCollectionFieldType.dropdown,
                      options: this.listColumns,
                      required: true
                    },
                    {
                      id:"CellView",
                      title:"Highlight",
                      type: CustomCollectionFieldType.dropdown,
                      options:[
                        {key:"Yes", text:"Yes"},
                        {key:"No", text:"No"}
                      ]
                    }
                  ],
                  disabled: false
                })
              ]
            }
          ]
        }
      ]
    };
  }

  protected get disableReactivePropertyChanges(): boolean {
    return true;
  } 
}
