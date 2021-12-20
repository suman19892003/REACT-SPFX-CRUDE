import * as React from 'react';
import { escape, find } from "@microsoft/sp-lodash-subset";

import {
  PeoplePicker,
  PrincipalType,
} from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { PrimaryButton } from "office-ui-fabric-react/lib/Button";
import * as _ from "lodash";
import { Link } from 'react-router-dom';

import * as moment from 'moment';

import './Custom.css';

import {RequestServices} from '../../Services/ServiceRequest';

export default class LandingLibrary extends React.Component<any, any> {

    public myServices: RequestServices;
    constructor(props){
      super(props);
      this.state={

        itemColl:[]
      }
      this.myServices=new RequestServices();
    }
  
    componentDidMount(){
      debugger
      this.myServices.getAllFileFromLibrary().then((itemColl)=>{
        console.log(itemColl);
        debugger;
        
        this.setState({itemColl:itemColl},()=>{
          //this.setState({item:"Added"})
        });
      })
    }

    viewMyApproval(){
      this.myServices.GetMyApprovalsFromLibraryRestAPI(this.props.context).then((itemColl)=>{        
        console.log(itemColl);
        debugger;
        
        let itemFromListColl=[];
        itemColl.map((element,z)=>{
          itemFromListColl.push({
            FileName:element.File.Name,
            Id:element.Id,
            Title:element.Title,
            
            AssignedTo:element.AssignedTo == "" || element.AssignedTo == null || element.AssignedTo == undefined?"":element.AssignedTo.Title      
          })
        })
        this.setState({itemColl:itemFromListColl},()=>{
          //this.setState({item:"Added"})
        });

      })
    }

    getFileFromFolder(){
      this.myServices.getFileFromFolder().then((itemColl)=>{        
        console.log(itemColl);
        debugger;
        
        // let itemFromListColl=[];
        // itemColl.map((element,z)=>{
        //   itemFromListColl.push({
        //     FileName:element.File.Name,
        //     Id:element.Id,
        //     Title:element.Title,
        //     AssignedTo:element.AssignedTo == "" || element.AssignedTo == null || element.AssignedTo == undefined?"":element.AssignedTo.Title      
        //   })
        // })
        this.setState({itemColl:itemColl},()=>{
          //this.setState({item:"Added"})
        });

      })
    }

    deleteFileFromLibrary(itemURL){
      debugger;
      this.myServices.deleteFileFromLibrary(itemURL);
    }

  public render(): React.ReactElement<any> {
    
    return (
        <div className="w100">
            <div className="list-container">
            <h3>Library Dashboard</h3>

            <PrimaryButton onClick={()=>this.viewMyApproval()}>View My Approval</PrimaryButton>
            <PrimaryButton onClick={()=>this.getFileFromFolder()}>Get File From Sub Folder</PrimaryButton>
            <Link className="editBtn" to={`/library/addedit`}>Upload Files</Link>
            
            <table style={{width:"100%"}} className="tableStyle">
            <tbody>
              <tr>
              <th>File Name</th>
              <th>Title</th>
              <th>Uploaded On</th>
              <th>Assigned To</th>
              
              <th></th>
            </tr>
            {this.state.itemColl.map((item)=>{
              return(
              <tr>
                <td>{item.FileName}</td>
                <td>{item.Title}</td>
                <td>{moment(item.UploadedDate).format("DD/MM/YYYY")}</td>
                <td>{item.AssignedTo}</td>               
                {/* <td><Link className="editBtn" to={`/library/addedit?Id=${item.Id}`}>Delete</Link></td> */}
                <td><PrimaryButton onClick={()=>this.deleteFileFromLibrary(`${item.FileURL}`)}>Delete</PrimaryButton></td>
              </tr>
            )})
          }
          </tbody>
          </table>             
        </div>
        <Link to={`/`}>Back</Link>  
        </div>
      );
    }
  }
  