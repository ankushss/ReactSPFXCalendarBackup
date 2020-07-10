import { SPHttpClient } from '@microsoft/sp-http';

export interface ICalendarProps {
  listName: string;
  spHttpClient: SPHttpClient;
  siteUrl: string;
  showPanel: boolean;  
  viewCollection: any[];
  dom:any;
  statusRender:any;
}
