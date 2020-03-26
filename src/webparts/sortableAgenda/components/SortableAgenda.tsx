import * as React from 'react';
import styles from './SortableAgenda.module.scss';
import { ISortableAgendaProps } from './ISortableAgendaProps';
import { escape } from '@microsoft/sp-lodash-subset';
import DragSortableList from 'react-drag-sortable';
import { Environment, EnvironmentType } from '@microsoft/sp-core-library';   

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
  public RetrieveSPData(){    
  var reactHandler = this;    
  var spRequest = new XMLHttpRequest();    

  spRequest.open('GET', 'https://nhy.sharepoint.com/sites/project-004066/_api/web/lists',true);    
  spRequest.setRequestHeader("Accept","application/json");        
  spRequest.onreadystatechange = () =>{    
        
      if (spRequest.readyState === 4 && spRequest.status === 200){    
          var result = JSON.parse(spRequest.responseText);    
          
          console.log(result);
          // reactHandler.setState({    
          //     items: result.value  
          // });    
      }    
      else if (spRequest.readyState === 4 && spRequest.status !== 200){    
          console.log('Error Occured !');    
      }    
  };    
  spRequest.send();    
  }    

  public componentDidMount() {
    this.RetrieveSPData();    
  }

  public render(): React.ReactElement<ISortableAgendaProps> {

    let tesst: {}; 

    if (Environment.type === EnvironmentType.Local) {  
      tesst = "local";
      // Local mode
    } else {
      tesst = "online";
      // Online mode
    }

    return (
      <div className={ styles.sortableAgenda }>
        <div className={ styles.container }>
          <div style={{display:"grid",gridTemplateColumns: "1fr 1fr 1fr 1fr 1fr", padding:"10px"}}>
            <span style={agendaStyling.agendaItemChild}>Order</span>
            <span style={agendaStyling.agendaItemChild}>Hours</span>
            <span style={agendaStyling.agendaItemChild}>Duration</span>
            <span style={agendaStyling.agendaItemChild}>Title</span>
            <span style={agendaStyling.agendaItemChild}>Controls</span>
          </div>
          {tesst}
            <DragSortableList  items={list} placeholder={placeholder} onSort={onSort} type="vertical"/>
        </div>
      </div>
    );
  }
}
