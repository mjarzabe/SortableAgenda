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
import { Dropdown, MessageBar, PrimaryButton, MessageBarType,IDropdownOption } from 'office-ui-fabric-react';
import {
  DetailsList,
  Selection,
  IColumn,
  buildColumns,
  IColumnReorderOptions,
  IDragDropEvents,
  IDragDropContext,
} from 'office-ui-fabric-react/lib/DetailsList';


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
      
    }  
};

var list = [
  {content: 
    (<div style={agendaStyling.agendaItemParent}>
        <div style={agendaStyling.agendaItemChild}>1</div>
        <div style={agendaStyling.agendaItemChild}>9:30-10:30</div>
        <div style={agendaStyling.agendaItemChild}>30min</div>
        <div style={agendaStyling.agendaItemChild}>Introcution</div>
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
        <div style={agendaStyling.agendaItemChild}>Dinner</div>
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
        <div style={agendaStyling.agendaItemChild}>Break</div>
        <div style={agendaStyling.agendaItemChild}>
          <button>More info</button>
        </div>
    </div>)
  }
];

var placeholder = (
  <div>placeholder</div>
);

export interface IDropdownControlledExampleState {
  selectedItem?: { key: string | number | undefined };
}

export interface ISelectState {      
  agendaState?:any[] | any;  
  meetingList?:any[] | any;
  agendaItemList?:any[] | any;
  selectedList: string;
  saveStatus: string;
  selectedItem?: { key: string | number | undefined };
} 

export default class SortableAgenda extends React.Component <ISortableAgendaProps,ISelectState> {
  constructor(props) {    
    super(props);    
    this.state = {  
      agendaItemList:[],
      meetingList: [],
      selectedList:"",
      saveStatus:"",
      selectedItem: undefined,
    };   
  }

  public componentWillMount() {
    const calendarList = "CalendarList";

    sp.web.lists.getByTitle(calendarList).items.getAll().then((items:any) => {      
        this.setState({meetingList: items});
    })
  };

  private onSort = (sortedList, dropEvent) => {
    this.setState({ agendaItemList: sortedList});
  };

  private handleChange = (event: React.FormEvent<HTMLDivElement>, item: IDropdownOption): void => { 
    const meetingAgendaList: string = "MeetingAgendaList";
    const data: any = sp.web.lists
                      .getByTitle(meetingAgendaList)
                      .items
                      .select("Title","AgendaOrder","Duration", "MeetingRef/Title", "MeetingRef/Id")
                      .expand("MeetingRef").filter(`MeetingRef/Title eq '${item.text}'`)
                      .getAll()
                      .then((items:any) => {
                        let data = items.map((item) =>
                          ({ content: 
                            <div style={agendaStyling.agendaItemParent}>
                                <div style={agendaStyling.agendaItemChild}>{item.AgendaOrder}</div>
                                <div style={agendaStyling.agendaItemChild}>8:30-10:30</div>
                                <div style={agendaStyling.agendaItemChild}>{item.Duration}min</div>
                                <div style={agendaStyling.agendaItemChild}>{item.Title}</div>
                                <div style={agendaStyling.agendaItemChild}>
                                  <button>More info</button>
                                </div>
                            </div>
                            , class: ""})
                          )

                        this.setState({selectedItem: item, agendaItemList:data});
                      })

  } 

  private saveStateOnSharepoint = () =>{
      alert("Save works");
      this.setState({saveStatus:"Your meeting agenda has been saved succesfuly!"});
  }

  public render(): React.ReactElement<ISortableAgendaProps> {

    const {listName} = this.props;
    const {meetingList, selectedList, agendaItemList, saveStatus} = this.state;

    return(
      <div className={ styles.sortableAgenda }>
        {/* <div>{listName}</div> */}

        <Dropdown
          label="Select Meeting Item"
          onChange={this.handleChange}
          placeholder="Select an option"
          options={meetingList.map(meeting => ({ key: meeting.Id, text: meeting.Title }))}
        />
        {saveStatus ? (
          <MessageBar messageBarType={MessageBarType.success} isMultiline={false}>
            {saveStatus}
          </MessageBar>
        ) : (
        "")}
        <div className={ styles.container }>
          <div style={{display:"grid",gridTemplateColumns: "1fr 1fr 1fr 1fr 1fr", padding:"10px"}}>
            <span style={agendaStyling.agendaItemChild}>Order</span>
            <span style={agendaStyling.agendaItemChild}>Hours</span>
            <span style={agendaStyling.agendaItemChild}>Duration</span>
            <span style={agendaStyling.agendaItemChild}>Title</span>
            <span style={agendaStyling.agendaItemChild}>Controls</span>
          </div>
           <DragSortableList items={agendaItemList} placeholder={placeholder} onSort={this.onSort} type="vertical"/>
        </div>
        <PrimaryButton text="Save" onClick={this.saveStateOnSharepoint} allowDisabledFocus/>
      </div>
    )
  } 
}
