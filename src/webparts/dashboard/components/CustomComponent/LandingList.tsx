import * as React from 'react';
import { escape, find } from "@microsoft/sp-lodash-subset";

import {
  PeoplePicker,
  PrincipalType,
} from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { PrimaryButton } from "office-ui-fabric-react/lib/Button";
import { TextField } from 'office-ui-fabric-react/lib';

import { Checkbox } from "office-ui-fabric-react/lib/Checkbox";
import { sp } from "@pnp/sp/presets/all";
import * as _ from "lodash";
import { Link } from 'react-router-dom';

import * as moment from 'moment';

import './Custom.css';

import {RequestServices} from '../../Services/ServiceRequest';

export default class LandingList extends React.Component<any, any> {

    public myServices: RequestServices;
    constructor(props){
      super(props);
      this.state={
        itemColl:[],
        attachmentfiles:[]
      }
      this.myServices=new RequestServices();
    }
  
    componentDidMount(){
      debugger
      this.myServices.GetRequestDetails('MyTestList').then((itemColl)=>{
        console.log(itemColl);
        debugger;
        let itemFromListColl=[];
        itemColl.map((element,index)=>{
          itemFromListColl.push({
            Id:element.Id,
            Title:element.Title,
            Description:element.Description,
            Department:element.Department,
            PublishDate:element.PublishDate,
            Technology: element.Technology,
            Sex:element.Sex,
            AssignedTo:element.AssignedTo == "" || element.AssignedTo == null || element.AssignedTo == undefined?"":element.AssignedTo.Title,
            LookUpColumn:element.LookUpColumn == "" || element.LookUpColumn == null || element.LookUpColumn == undefined?"":element.LookUpColumn.Title
          
          })
        })
        this.setState({itemColl:itemFromListColl},()=>{
          //Get All Attachment From List
          this.myServices.GetAllListItemAttachment('MyTestList').then((attachmentColl)=>{
            debugger;
            //let attachmentfiles=[]
            attachmentColl.forEach((listItem: any) => {
              listItem.AttachmentFiles.forEach((afile: any) => {
                debugger;
                //let downloadUrl = this.context.pageContext.web.absoluteUrl + "/_layouts/download.aspx?sourceurl=" + afile.ServerRelativeUrl; 
                let downloadUrl = this.props.webURL + "/_layouts/download.aspx?sourceurl=" + afile.ServerRelativeUrl; 
                
               this.state.attachmentfiles.push({
                  Id:listItem.Id,
                  Title:listItem.Title,
                  FileName:afile.FileName,
                  DownLoadURL:downloadUrl
                })
                //attachmentfiles += `<li>(${listItem.Id}) ${listItem.Title} - ${afile.FileName}</li>`;  
              });
            });
            console.log(this.state.attachmentfiles);
            this.setState({Item:"Updated"})
            debugger;


          })
        });
      })
    }

    viewMyApproval(){
      this.myServices.GetMyApprovalsRestAPI(this.props.context).then((itemColl)=>{
        //this.myServices.GetMyApprovals(this.props.context.pageContext.user.email).then((itemColl)=>{
        
        console.log(itemColl);
        debugger;
        this.setState({itemColl:itemColl},()=>{
          //this.setState({item:"Added"})
        });
      })
    }

    filterRecordDropDown(item){
      this.myServices.GetFilteredRecordDropdown(item).then((itemColl)=>{
        console.log(itemColl);
        debugger;
        this.setState({itemColl:itemColl},()=>{
          //this.setState({item:"Added"})
          console.log('Item Fetched');
          debugger;
        });
      })
    }

    filterRecordCheckBox(item){
      this.myServices.GetFilteredRecordCheckBox(item).then((itemColl)=>{
        console.log(itemColl);
        debugger;
        this.setState({itemColl:itemColl},()=>{
          //this.setState({item:"Added"})
          console.log('Item Fetched');
          debugger;
        });
      })
    }

    filterRecordDate(item){
      this.myServices.GetFilteredRecordDate(item).then((itemColl)=>{
        console.log(itemColl);
        debugger;
        this.setState({itemColl:itemColl},()=>{
          //this.setState({item:"Added"})
          console.log('Item Fetched');
          debugger;
        });
      })
    }

    filterMultipleRecord(item){
      this.myServices.GetMultipleFilteredRecord("Test First Item","HR").then((itemColl)=>{
        console.log(itemColl);
        debugger;
        if(itemColl){
        this.setState({itemColl:itemColl},()=>{
          //this.setState({item:"Added"})
          console.log('Item Fetched');
          debugger;
        });
      }
      })
    }

    filterRecordLookupColumn(item){
      this.myServices.GetFilteredRecordLookUpColumn(item).then((itemColl)=>{
        console.log(itemColl);
        debugger;
        if(itemColl){
        this.setState({itemColl:itemColl},()=>{
          //this.setState({item:"Added"})
          console.log('Item Fetched');
          debugger;
        });
      }
      })
    }

    renderDynamicDOM(){
      sp.web.currentUser.groups.get().then((r: any) => {  
        let grpNames: string ="";  
        r.forEach((grp: any) =>{  
          grpNames += "<li>"+grp["Title"]+"</li>"  
        });      
        grpNames = "<ul>"+grpNames+"</ul>";  
        this.renderData(grpNames);  
      });
    }

    private renderData(strResponse: string): void { 
      debugger 
      //const htmlElement : HTMLElement = document.getElementById("#pnpinfo"); 
      let myContainer: HTMLDivElement | null = document.querySelector("#pnpinfo");
      if (myContainer instanceof HTMLDivElement) {
        myContainer.innerHTML = strResponse//'<h1>Test</h1>';
      }
      //htmlElement.innerHTML = strResponse;  
    }
    
    currentUserDataDynamicDOM(){
      sp.web.currentUser.get().then((r: any) => {  
        this.renderData(r['Title']);  
      }); 
    }

    private async createBatchRequest(){
      const list = sp.web.lists.getByTitle('MyTestList');
      const ent = await list.getListItemEntityTypeFullName();

      const batch = sp.web.createBatch();
      const items = list.items.inBatch(batch);

      items.add({ Title: 'Item 1' }, ent).catch(console.warn);
      items.add({ Title: 'Item 2' }, ent).catch(console.warn);
      items.add({ Title: 'Item 3' }, ent).catch(console.warn);
      items.add({ Title: 'Item 4' }, ent).catch(console.warn);
      items.add({ Title: 'Item 5' }, ent).catch(console.warn);

      await batch.execute();
      console.log('Done');
    }

  public render(): React.ReactElement<any> {
    
    return (
        <div className="w100">
            <div className="list-container">
            <h3>List Dashboard</h3>
            <h4>Dynamic DOM Manipulation</h4>
            <PrimaryButton onClick={()=>this.renderDynamicDOM()}>Bind Dynamic Logged User Group</PrimaryButton>
            <PrimaryButton onClick={()=>this.currentUserDataDynamicDOM()}>Bind Current User Details</PrimaryButton>
            
            <div id="pnpinfo"></div>

            <h4>List Attachment Operation</h4>
            <table style={{width:"100%"}} className="tableStyle">
            <tbody>
              <tr>
              <th>Item ID</th>
              <th>Title</th>
              <th>File Name</th>
              <th></th>
            </tr>
            {this.state.attachmentfiles && this.state.attachmentfiles.map((item)=>{
              debugger
              return(
              <tr>
                <td>{item.Id}</td>
                <td>{item.Title}</td>
                <td>{item.FileName}</td>
                <td><a href={item.DownLoadURL}>Download</a></td>
                </tr>
            )})
            }
            </tbody>
          </table>

            <h4>List Item Operation</h4>
            <PrimaryButton onClick={()=>this.createBatchRequest()}>Create Batch Request</PrimaryButton>
            <PrimaryButton onClick={()=>this.viewMyApproval()}>View My Approval</PrimaryButton>
            <PrimaryButton onClick={()=>this.filterRecordCheckBox('Java')}>Checkbox Filter Record</PrimaryButton>
            <PrimaryButton onClick={()=>this.filterRecordDropDown('Admin')}>Dropdown Filter Record</PrimaryButton>
            <PrimaryButton onClick={()=>this.filterRecordDate('suman')}>Date Filter Record</PrimaryButton>
            <PrimaryButton onClick={()=>this.filterMultipleRecord('Item1')}>Multiple Filter</PrimaryButton>
            <PrimaryButton onClick={()=>this.filterRecordLookupColumn('Item1')}>Lookup Filter From Other List</PrimaryButton>
            <Link className="editBtn" to={`/list/addedit`}>Add New List Item</Link>
            
            <table style={{width:"100%"}} className="tableStyle">
            <tbody>
              <tr>
              <th>Title</th>
              <th>Department</th>
              <th>Sex</th>
              <th>Technology</th>
              <th>Assigned To</th>
              <th>Publish Date</th>
              <th>Look Up Column</th>
              <th></th>
            </tr>
            {this.state.itemColl.map((item)=>{
              return(
              <tr>
                <td>{item.Title}</td>
                <td>{item.Department}</td>
                <td>{item.Sex}</td>
                <td>{item.Technology}</td>
                <td>{item.AssignedTo}</td>
                <td>{moment(item.PublishDate).format("DD/MM/YYYY")}</td>
                <td>{item.LookUpColumn}</td>
                <td><Link className="editBtn" to={`/list/addedit?Id=${item.Id}`}>Edit</Link></td>
                {/* <td><a href="#!" className="editBtn">Edit</a></td> */}
              </tr>
            )})
          }
          </tbody>
          </table>             
        </div>

        
            <Link to={`/list`}>Back</Link>  
        </div>
      );
    }
  }
  