import * as React from 'react';
import styles from './SortableAgenda.module.scss';
import { ISortableAgendaProps } from './ISortableAgendaProps';
import { escape } from '@microsoft/sp-lodash-subset';
import DragSortableList from 'react-drag-sortable';
import { Environment, EnvironmentType } from '@microsoft/sp-core-library';   
import { SPHttpClient, SPHttpClientResponse, SPHttpClientConfiguration } from '@microsoft/sp-http'; 
import { sp } from "@pnp/sp";
import "@pnp/sp/webs"; 
import "@pnp/sp/lists";
import "@pnp/sp/items";

const agendaStyling = {
    agendaItemParent:{
      display: "grid",
      gridTemplateColumns: "repeat(5, 1fr)",
      padding:"10px",
      border:"1px solid #e3e3e3",
      background: "#fafafa",
      margin:"5px"
    },
    agendaItemChild:{
      width:"100%"
    }  
};

var list = [
  {content: 
    (<div style={agendaStyling.agendaItemParent}>
        <div style={agendaStyling.agendaItemChild}>1</div>
        <div style={agendaStyling.agendaItemChild}>9:30-10:30</div>
        <div style={agendaStyling.agendaItemChild}>30min</div>
        <div style={agendaStyling.agendaItemChild}>Meeting ss</div>
        <div style={agendaStyling.agendaItemChild}>
          
          <button>More info</button>
        </div>
    </div>)
  },
  {content: 
    (<div style={agendaStyling.agendaItemParent}>
        <div style={agendaStyling.agendaItemChild}>2</div>
        <div style={agendaStyling.agendaItemChild}>8:30-10:30</div>
        <div style={agendaStyling.agendaItemChild}>30min</div>
        <div style={agendaStyling.agendaItemChild}>Meeting ss</div>
        <div style={agendaStyling.agendaItemChild}>
          <button>More info</button>
        </div>
    </div>)
  },
  {content: 
    (<div style={agendaStyling.agendaItemParent}>
        <div style={agendaStyling.agendaItemChild}>3</div>
        <div style={agendaStyling.agendaItemChild}>8:30-10:30</div>
        <div style={agendaStyling.agendaItemChild}>30min</div>
        <div style={agendaStyling.agendaItemChild}>Meeting ss</div>
        <div style={agendaStyling.agendaItemChild}>
          <button>More info</button>
        </div>
    </div>)
  }
];

var placeholder = (
  <div></div>
);

var onSort = (sortedList, dropEvent) => {

};

export default class SortableAgenda extends React.Component <ISortableAgendaProps, {}> {  

  /*private getCalendarEvents(): Promise<Array<string>> {
    return new Promise<Array<string>>((resolve: (options: Array<string>) => void, reject: (error: any) => void) => {
      this.context.spHttpClient.get(`${this.context.pageContext.web.absoluteUrl}/_api/list/GetListByTitle(${this.props.listName})/items`, SPHttpClient.configurations.v1, {
        headers: {
          'odata-version': '3.0',
          'accept': 'application/json;odata=verbose',
          'content-type': 'application/json;odata=verbose'
        }
      }).then((res: SPHttpClientResponse) => {
        console.log(res.json());
       
        var mappedArray = [];
        resolve(mappedArray);
        
      }).catch(error => {
      
      });
    });
  }
  private getEventItems(): Promise<Array<string>> {
    return new Promise<Array<string>>((resolve: (options: Array<string>) => void, reject: (error: any) => void) => {
      this.context.spHttpClient.get(`${this.context.pageContext.web.absoluteUrl}/_api/list/GetListByTitle('AgendaMaster')/items?$fitler=MeetingRef eq ${this.props.currentMeeting}`, SPHttpClient.configurations.v1, {
        headers: {
          'odata-version': '3.0',
          'accept': 'application/json;odata=verbose',
          'content-type': 'application/json;odata=verbose'
        }
      }).then((res: SPHttpClientResponse) => {
        console.log(res.json());
       
        var mappedArray = [];
        resolve(mappedArray);
        
      }).catch(error => {
      
      });
    });
  }*/
  public async recalculateItems(){
    const calendarList = "CalendarList";
    const list = sp.web.lists.getByTitle(calendarList);
    const r = await list.select("Id")();
    alert(r.Id);


  }

  private async getCalendarEvents(){

    const calendarList = "CalendarList";

    const r1 = await sp.web.lists.getByTitle(calendarList).items.getAll();
    console.log(r1.length);

    // set page size
    const r2  = await sp.web.lists.getByTitle(calendarList).items.getAll(4000);
    console.log(r2.length);

    const r3  = await sp.web.lists.getByTitle(calendarList).items.select("Title").top(4000).getAll();
    console.log(r3.length);

    const r4 = await sp.web.lists.getByTitle(calendarList).items.select("Title").filter("Title eq 'Test'").getAll();
    console.log(r4.length);

  }
  private async getEventItems(){

    const agendaMaster = "AgendaMaster";

    const r1 = await sp.web.lists.getByTitle(agendaMaster).items.getAll();
    console.log(r1.length);

    // set page size
    const r2  = await sp.web.lists.getByTitle(agendaMaster).items.getAll(4000);
    console.log(r2.length);

    const r3  = await sp.web.lists.getByTitle(agendaMaster).items.select("Title").top(4000).getAll();
    console.log(r3.length);
  }
  private async saveToMaster(itemID){
    const calendarList = "CalendarList";
    let list = sp.web.lists.getByTitle(calendarList);
    const i = await list.items.getById(itemID).update({
      Title: "My New Title",
      Description: "Here is a new description"
    });

  }
  public render(): React.ReactElement<ISortableAgendaProps> {
    this.recalculateItems();
    this.getCalendarEvents();
    return (
      <div className={ styles.sortableAgenda }>
        <div>{this.props.listName}</div>
        
        <div><select id='calenarEvents'></select></div>
        <div className={ styles.container }>
          <div style={{display:"grid",gridTemplateColumns: "1fr 1fr 1fr 1fr 1fr", padding:"10px"}}>
            <span style={agendaStyling.agendaItemChild}>Order</span>
            <span style={agendaStyling.agendaItemChild}>Hours</span>
            <span style={agendaStyling.agendaItemChild}>Duration</span>
            <span style={agendaStyling.agendaItemChild}>Title</span>
            <span style={agendaStyling.agendaItemChild}>Controls</span>
          </div>
            <DragSortableList items={list} placeholder={placeholder} onSort={onSort} type="vertical"/>
        </div>
      </div>
    );
  }
}
