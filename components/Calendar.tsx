import * as React from 'react';
import styles from './Calendar.module.scss';
import { ICalendarProps } from './ICalendarProps';
import { escape } from '@microsoft/sp-lodash-subset';

import { Panel, PanelType,IPanelProps } from 'office-ui-fabric-react/lib/Panel';
import { IFramePanel } from "@pnp/spfx-controls-react/lib/IFramePanel";
import { initializeIcons } from 'office-ui-fabric-react/lib/Icons';
import { getIconClassName } from '@uifabric/styling';
import {Spinner, SpinnerSize } from 'office-ui-fabric-react';
import Iframe from 'react-iframe';
import { CommandBar, ICommandBarItemProps } from 'office-ui-fabric-react/lib/CommandBar';

import {  
  SPHttpClient,
  SPHttpClientResponse
} from '@microsoft/sp-http';

import { sp } from "@pnp/sp";
import { Web } from "@pnp/sp/webs";
import { ICamlQuery} from "@pnp/sp/lists";
import "@pnp/sp/lists";

import {IPanelState} from './IPanelState';

var $: any = require('jquery');
import 'fullcalendar';
import * as FC from 'fullcalendar';
var moment: any = require('moment');
require('../css/customCss.css');



export default class Calendar extends React.Component<ICalendarProps, IPanelState>{

    constructor(props: ICalendarProps, state: IPanelState){
    super(props);
    this.state = {
      iFrameUrl:"", 
      showPanel: false  
    };
  }
  
  public async componentDidMount() {
    
    console.log("componentDidMount");
    await this.displayTasks();
    
  }

  public async componentDidUpdate() { 
    
    console.log("componentDidUpdate");
    await this.displayTasks();
    
  }

  private setShowPanel(showPanel: boolean) {
    this.setState({
      showPanel: showPanel
    });
  }

  private onPanelClosed() {
    this.setState({
      showPanel: false
    });
  }

  private async getListViewQuery(viewName: string):Promise<string> { 
    let query:string;
    try {
      const result = await this.props.spHttpClient.get(this.props.siteUrl + `/_api/web/lists/GetbyTitle('`+this.props.listName+`')/Views/getbyTitle('`+viewName+`')`, SPHttpClient.configurations.v1)  
      .then((response: SPHttpClientResponse) => {
        return response.json();  
      });
      query=result.ViewQuery;            
    } catch (error) {
      
    }
    return query;   
  }

  private async loadListItems(query: string)
  {
    let _listItems;
    const caml: ICamlQuery = {
      ViewXml: '<View><Query>'+query+'</Query></View>'
      }; 
    const web=Web(this.props.siteUrl);
    const result= web.lists.getByTitle(this.props.listName).getItemsByCAMLQuery(caml)
    .then((resp)=>
    {
      return resp;
    });    
    _listItems=result;
    return _listItems;
  }

  private async getListId()
  {
    let _listUrl;
    const web=Web(this.props.siteUrl);
    const listUrl=web.lists.getByTitle(this.props.listName).get()
    .then((resp)=>
    {
      return resp;
    });
    _listUrl=listUrl;
    return _listUrl;
  }

  private async displayTasks() {
    
    console.log("display Tasks");
    var events=[];
    var cellEvents=[];
    var backgroundColor;
    var colorEvents;
    var listId= await this.getListId();    
    var Url=this.props.siteUrl +"/_layouts/15/listform.aspx?PageType=6&ListId="+listId.Id +"&IsDlg=1";
    if(this.props.viewCollection.length>0)
    {
      for(var i=0;i<this.props.viewCollection.length;i++)
      {
        var query= await this.getListViewQuery(this.props.viewCollection[i].View);
        query=`<XML>`+query+`</XML>`;
        var xmlDoc= $.parseXML(query); 
        if($(xmlDoc).find("Where").length>0)
        {
          if($(xmlDoc).find("DateRangesOverlap").length>0)
          {
            $(xmlDoc).find("DateRangesOverlap").remove();
            var oSerializer = new XMLSerializer();
            query = oSerializer.serializeToString(xmlDoc);            
            if($(xmlDoc).find("And").length>0)
            {
              query=`<Where>`+$(xmlDoc).find("And")[0].innerHTML+`</Where>`;            
            }
            else{
                         
            }
          }
        }
        console.log(query);
        var items = await this.loadListItems(query);
        if(this.props.viewCollection[i].CellView=="Yes")
        {
          items.map((item)=>{
            cellEvents.push(
              {
                title:item[this.props.viewCollection[i].Column],
                id:item.Id,
                start: moment.utc(item.EventDate),
                end: moment.utc(item.EndDate),
                color: this.props.viewCollection[i].Color,
                extendedProps:{
                  dateDiff: moment(item.EndDate).diff(moment(item.EventDate),'days')
                }
              });
          });
              colorEvents =this.getColouredDates(cellEvents);
              backgroundColor= cellEvents[0].color;
        }
        else{
          items.map((item) => {
            events.push(
              {
                //title:item[this.properties.listColumn],
                title:item[this.props.viewCollection[i].Column],
                id:item.Id,
                start: moment.utc(item.EventDate),
                end: moment.utc(item.EndDate),
                color: this.props.viewCollection[i].Color,
                extendedProps:{
                  order: this.props.viewCollection[i].Order
                }
              });
            });          
        }        
      }
    }    
    
    $('#calendar').fullCalendar('destroy');
    $('#calendar').fullCalendar({
      weekends: true,
      eventBackgroundColor: "red",
      header: {
        left: 'prev,next,title',
        center: '',
        right: ''
      },
      views: {
        month: { 
          columnFormat: 'dddd'
        }
      },
      displayEventTime: false,       
      editable: false,      
      timezone: "UTC",
      droppable: false,
      eventLimit: 4,
      fixedWeekCount:false,
      eventClick: (calEvent: FC.EventObjectInput, jsEvent: MouseEvent, view: FC.View) => { 
          this.setState({
            iFrameUrl: Url +"&ID="+calEvent.id,            
          });
          this.setShowPanel(true);       
        return false;
      },
      events: events, 
      dayRender: (date,cell)=>
      { 
        if(colorEvents!=undefined)
        { 
          if(colorEvents.indexOf(cell.attr("data-date"))!=-1)
          {
            $("td[data-date='"+cell.attr("data-date")+"']").css("background-color",backgroundColor);
          }
        }
      }   
    });
    this.createLegend(listId.Id);
  }

  private getColouredDates(colorEvents)
  {
    var allDates = [];
    console.log(colorEvents);
    for(var i=0;i<colorEvents.length;i++)
    {
      var start=colorEvents[i].start;
      var end=colorEvents[i].end;
      while(start<end)
      {
        if(allDates.indexOf(start.format('YYYY-MM-DD'))==-1)
        {
          allDates.push(start.format('YYYY-MM-DD'));
        }       
        start = start.add(1, 'days');
      }
    }
    return allDates;
  }

  private createLegend(id)
  {    
    $(".fc-left").prepend(`<a style="cursor:pointer;text-align:center;margin-bottom: 6px;" class="commandNew">
    <span style="align-items:center;justify-content:center;text-align:center;">
      <i class="${getIconClassName('Add')} ${styles.addNew}" /><span>New Event</span>
    </span>
    </a><br/>
    `);
    
    this.props.dom.querySelector('.commandNew').addEventListener("click", () => {
      this.setState({
        iFrameUrl:this.props.siteUrl +"/_layouts/15/listform.aspx?PageType=8&ListId="+id +"&IsDlg=1",
        //this.props.siteUrl +"/Lists/TestModern/NewForm.aspx"
                  
      });
      this.setShowPanel(true); 
    });
    var legend='';//
    for(var i=0;i<this.props.viewCollection.length;i++)
    {
      legend=legend+'<span style="margin-left:0px;height:15px;width:15px;background-color:'+this.props.viewCollection[i].Color+'"></span><div style="font-size:12px;padding-left:6px;">'+this.props.viewCollection[i].View+'</div><br/>';
      
    }
    $(".fc-right").html(legend);
  }//

  private async newButtonClick()
  {
    var listId= await this.getListId();
    this.setState({
      iFrameUrl: this.props.siteUrl +"/_layouts/15/listform.aspx?PageType=8&ListId="+listId.Id +"&IsDlg=1",            
    });
    this.setShowPanel(true);
    
  }
  

  private _ButtonClick(): void {  
    
    this.props.dom.querySelector('.fc-next-button').addEventListener('click', () => {  
   
               alert("click next");
   
     });
     this.props.dom.querySelector('.fc-prev-button').addEventListener('click', () => {  
   
      alert("click prev");

}); 
 }  

  public render(): React.ReactElement<ICalendarProps> {

    return (      
      <div className={styles.calendar}>        
       
        <div className="ms-Grid">
          <div id="commandBar">            
          </div>
          <div className="ms-Grid-row">
            <div id="calendar"></div>
          </div>
        </div> 
          
        <IFramePanel url={this.state.iFrameUrl}
            closeButtonAriaLabel="Close"
            isOpen={this.state.showPanel}
            scrolling="auto"
            onDismiss={this.onPanelClosed.bind(this)}
    />        
      </div>
    );
  }
}

/*
<a onClick={this.newButtonClick.bind(this)}>New</a>

iframeOnLoad={this._onIframeLoaded.bind(this)}
<Panel isBlocking={false} isOpen={this.state.showPanel} onDismiss={this.onPanelClosed.bind(this)} type={PanelType.medium}
           closeButtonAriaLabel="Close">
              <Iframe url={this.state.iFrameUrl} height="100%" width="100%" display="block" position="relative"/>
        </Panel>

<IFramePanel url= type={PanelType}
              closeButtonAriaLabel="Close"      
             headerText="Edit"
             isOpen={this.state.showPanel}
             onDismiss={this.onPanelClosed.bind(this)}
              />

<Panel isBlocking={false} isOpen={this.state.showPanel} onDismiss={this.onPanelClosed.bind(this)} type={PanelType.custom}
          customWidth="500px" closeButtonAriaLabel="Close">
          <Label style={{fontWeight: "bolder", textAlign: "center", marginBottom: "30px"}}>Booking Details</Label>
          <Label style={{fontWeight: "bold"}}>Vehicle Details</Label>
          <Label style={{fontWeight: "bold"}}>Start Date and Time</Label>
          <Label>{this.state.StartDate}</Label>
          <Label style={{fontWeight: "bold"}}>End Date and Time</Label>
          <Label>{this.state.EndDate}</Label>          
        </Panel>
      */
