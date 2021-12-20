import * as React from 'react';
import { escape, find } from "@microsoft/sp-lodash-subset";

import {
  PeoplePicker,
  PrincipalType,
} from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { sp } from "@pnp/sp/presets/all";
import { Link } from 'react-router-dom';
import {RequestServices} from '../../Services/ServiceRequest';
import { TextField,Label,PrimaryButton } from 'office-ui-fabric-react/lib';

export default class LibraryOperation extends React.Component<any, any> {

    public myServices: RequestServices;
    constructor(props){
      super(props);
      this.state={
        FormTitle:'',
        FormAssignedTo:'',
        FormAssignedToId:'',
        EnclouserUploadedFiles:[]
      }
      this._onChange=this._onChange.bind(this)
      this.myServices=new RequestServices();
    }

    public _getPeoplePickerItems = async (items: any[]) => {
      debugger;
      this.setState({ FormAssignedTo: items[0].id });
      this.setState({ FormAssignedToId: items[0].id });
    };

    public onFileUpload (file) {
      debugger;
      let File = file;
      let isValid = true;
      let fileArray = [];
      let fileNames = [];
      let files = [];
      files = this.state.EnclouserUploadedFiles;
      
      for (let i=0;i<File.length;i++)
      {
        let currentFile = File[i];
        var isFiileExists = false;
        
        for (let index = 0; index< files.length;index++)
        {
          const element = files[index];
          if(element.File.name && element.File.name == currentFile.name)
          {
            isFiileExists =  true;
          }
          else if (element.File.FileLeafRef && element.File.FileLeafRef == currentFile.name)
          {
            isFiileExists =  true;
          }
        }
        if (isFiileExists == false)
        {
          fileNames.push(currentFile.name);
          files.push({
            File:currentFile
          })
        }
         
      }
      this.setState({
        EnclouserUploadedFiles:files
      })
    }
  
    public bindSavedEnclouserUploadedFiles () {
      debugger;
      let uploadedFiles = this.state.EnclouserUploadedFiles;
      let Data = uploadedFiles,
      MakeItem = (y,i)=>{
        let fileName = y.File;
        if( fileName && fileName.name){
          let files = [];
          files = fileName.name;
           return (
            <li>     
              <span>
               <a href="javascript:{}">{files}
             </a>
             </span>
             {<a title={'Deselect'} onClick={(e)=>this.removeSelectedFile(fileName)}><i
             className="fa fa-times-circle" arua-aria-hidden="true"
           ></i>Remove</a>}
          </li>
           )       
        }
        else if (fileName && fileName.FileLeafRef){
          let files = fileName.FileLeafRef;
          if(files.indexOf('~') != -1){
            let data = files.split('~');
            files = data[1];
          }
          return(
            <li>
              <span>
                 {/* <a href="javascript:{}" onClick={() => {this.downloadFile(fileName.FileRef)}}> */}
                 <a href="javascript:{}">
                   {files}
                  </a>
                 </span>
                {/* <a title={"Delete"} onClick={(e)=>this.deleteUploadedFile(fileName.FileLeafRef)}> */}
                <a title={"Delete"}>
                 <i className="fa fa-times-circle" aria-hidden="true"></i>
                </a>
            </li>
          )
        }
      }
       uploadedFiles.map((y,i)=>{
        let fileName = y.File;
        if( fileName && fileName.name){
          let files = [];
          files = fileName.name;
          return (
            <li>
              <span>
            <a href="javascript:{}">{files}
            </a>
            </span>
            {<a title={'Deselect'} onClick={(e)=>this.removeSelectedFile(fileName)}><i
            className="fa fa-times-circle-o" arua-aria-hidden="true"
          ></i></a>}
          </li>     
          )
        }
        else if (fileName && fileName.FileLeafRef){
          let files = fileName.FileLeafRef;
          if(files.indexOf('~') != -1){
            let data = files.split('~');
            files = data[1];
          }
          return(
            <li>
              <span>
                {/* <a href="javascript:{}" onClick={() => {this.downloadFile(fileName.FileRef)}}> */}
                <a href="javascript:{}">
                  {files}
                  </a>
                </span>
                {/* <a title={"Delete"} onClick={(e)=>this.deleteUploadedFile(fileName.FileLeafRef)}> */}
                <a title="Delete">
                  <i className="fa fa-times-circle-o" aria-hidden="true"></i>
                </a>
            </li>
            )
           }
          }) 
          if( Data && Data.length){
            return(
              <div>
                <label>Attached Files</label>
                <ul>
                  {Data.map(MakeItem)}
                </ul>
              </div>
            )
          }
      }
  
      public removeSelectedFile (row) {
        debugger;
        let rowData =row;
        if(this.state.EnclouserUploadedFiles && this.state.EnclouserUploadedFiles.length>0){
          let newArray = [];
          for(let index = 0; index<this.state.EnclouserUploadedFiles.length;index++){
            const element = this.state.EnclouserUploadedFiles[index];
            if(element && element.File.name){
              if(row.name != element.File.name){
                newArray.push(element);
              }
            }
            else{
              newArray.push(element);
            }
          }
          this.setState({
            EnclouserUploadedFiles:newArray
          })
        }
      }
  
      public onSubmit () {
        debugger;
        var uploadedFileName = '';
        for(let index = 0;index<this.state.EnclouserUploadedFiles.length;index++){
          const element = this.state.EnclouserUploadedFiles[index];
          var fileName = element.File.name;
          var fileExist:string = fileName.substring(fileName.lastIndexOf('.'),fileName.length);
          uploadedFileName = fileName;
          fileName = fileName.replace(/[\)!@#$%^&*_+;<(){}>?/|\,:-]+/g,"-");
          const file = sp.web.getFolderByServerRelativeUrl('MyTestLibrary').files.add(fileName,element.File,true).then((
          result:any
        )=>{
          debugger
          result.file.getItem().then((file)=>{
            debugger;
            file.update({  
              Title: this.state.FormTitle,
              AssignedToId:this.state.FormAssignedTo
          }).then((myupdate) => {  
            console.log(myupdate);  
            console.log("Metadata Updated");    
          });
        })
        })
        }
      }

      private _onChange(event){
        this.setState({[event.target.name]: event.target.value});
        debugger;
      }

  public render(): React.ReactElement<any> {
    return (<div>
        <div className="w100">
        <h3>Add Files</h3>

        <TextField
          label="Title"
          id="txtTitle"
          required={false}
          multiline={false}
          value={this.state.FormTitle}
          name='FormTitle'
          onChange={this._onChange}
        />

        <Label>Assigned To</Label>
          <PeoplePicker
            context={this.props.context}
            personSelectionLimit={1}
            required={false}
            onChange={this._getPeoplePickerItems}
            defaultSelectedUsers={[this.state.FormAssignedTo?this.state.FormAssignedTo:""]}
            showHiddenInUI={false}
            principalTypes={[PrincipalType.User]}
            resolveDelay={1000}
            ensureUser={true}
          />

        <div className="col col-md-3">
          <label className="form-control-label" >Enclosure:</label>
          </div>
          <div className="col-8 col-md-6">
          <input type="file" multiple={true} id='uploadFile' onChange={(e)=>this.onFileUpload(e.target.files)}></input>
          </div>
          </div>
          {this.state.EnclouserUploadedFiles && this.state.EnclouserUploadedFiles.length>0?  
            this.bindSavedEnclouserUploadedFiles():""
          }

        <div className="row"> 
          <div className='col-sm-12'>
          <div className="col-sm-2"> 
              <button className="btn btn-primary"  onClick={(e)=>this.onSubmit()} style={{color:'white',padding:'5px',backgroundColor:'orange',display:'inline-block',borderRadius:'20px',border:'1px solid orange',width:'150px',marginLeft:'20px'}} >Save</button>            
          </div>
          </div>
        </div>

        <Link to={`/`}>Back</Link>  
        </div>
      );
    }
  }
  