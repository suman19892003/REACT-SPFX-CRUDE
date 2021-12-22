import * as React from 'react';
import { escape, find } from "@microsoft/sp-lodash-subset";
import { sp } from "@pnp/sp/presets/all";
import * as _ from "lodash";
import { Link } from 'react-router-dom';
//Custom Field Office UI Fabric
import { DatePicker, IDatePickerStrings } from 'office-ui-fabric-react/lib/DatePicker';
import { TextField,Dropdown,IDropdownOption,Label,PrimaryButton,ChoiceGroup,IChoiceGroupOption } from 'office-ui-fabric-react/lib';
import { Checkbox} from 'office-ui-fabric-react/lib/Checkbox';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import IDashboardState from '../IDashboardState'
//Custom Variable Import
import {DepartmentOptions,SexOptions,DatePickerStrings,FormatDate,checkOptions} from './IListVariable';



import {RequestServices} from '../../Services/ServiceRequest';

//Columns used in List Operation
// Title	Single line of text	
// Description	Multiple lines of text	
// Department	Choice	- Single Select Dropdown
// IsActive	Yes/No	
// PublishDate	Date and Time	
// Technology	Choice --Multiple Select Checkbox
//AssignedTo -- People Picker
//LookUp -- Lookup Column from another list

export default class ListOperation extends React.Component<any, any> {

    public myServices: RequestServices;
    constructor(props) {
      super(props);
      this.state = {

        itemColl: [],
        UserDetails:[],
        EmpName:'',
        EmpManagerName:'',
        LookupColumnColl:[],

        CurrentItemID:0,
        FormTitle:'',
        FormDescription:'',
        FormDepartment:'',
        FormPublishDate:new Date(),
        FormAssignedTo:'',
        FormAssignedToId:'',
        FormAssignedMultiple:[],
        FormAssignedMultipleID:[],
        FormSex:'',
        FormLookUp:'',
        SelectedTechnology:[]
      };
      this.myServices=new RequestServices();
      this._onChange=this._onChange.bind(this);
      this.onChoiceChange=this.onChoiceChange.bind(this);
      this._onChangeCheck=this._onChangeCheck.bind(this);
      this.addFile=this.addFile.bind(this);
    }
  
    componentDidMount(){
      this.getCurrentLoginUser().then(()=>{
        this.getLookUpAndDropdownFromList('lookup').then(()=>{
          if (window.location.href.indexOf("Id") > -1) {
            let itemId = window.location.href.split("#/")[1].split("?Id=")[1];
            this.setState({ CurrentItemID: itemId });
            this.myServices.GetRequestDetailsId(itemId).then((res)=>{
              debugger;
              console.log(res);
              this.setState({ FormTitle: res.Title });
              this.setState({ FormDescription: res.Description });
              this.setState({ FormDepartment: res.Department });
              this.setState({ FormPublishDate: new Date(res.PublishDate) })
              this.setState({ FormAssignedTo: res.AssignedTo== null ? "" : res.AssignedTo.EMail });
              this.setState({ FormAssignedToId: res.AssignedTo== null ? "" : res.AssignedTo.ID });
              this.setState({ SelectedTechnology: res.Technology== null ? [] : res.Technology });
              this.setState({ FormSex: res.Sex });
              this.setState({ FormDepartment: res.Department });
              //this.setState({ FormAssignedMultiple: res.AssignedMultipleTeam== null ? [] : res.AssignedMultipleTeam.EMail });

              let res1:any[]=res.AssignedMultipleTeam;
              let getSelectedUsers:string[] = []//res.AssignedMultipleTeam== null ? [] : res.AssignedMultipleTeam.EMail  
              let getSelectedUsersID:string[] = []
              for (let item in res1) {  
                getSelectedUsers.push(res1[item].EMail); 
                getSelectedUsersID.push(res1[item].ID);  
              }  
              this.setState({ FormAssignedMultiple: getSelectedUsers }); 
              this.setState({ FormAssignedMultipleID: getSelectedUsersID });
              debugger;
              console.log(this.state)

            })
          }
        })
      }) 
    }

    
  private addFile(event) {
    var hasinvalidcharacters = false;
    let resultFile = event.target.files;
    console.log(resultFile);
    let fileInfos = [];
    for (var i = 0; i < resultFile.length; i++) {
      var fileName = resultFile[i].name;
      console.log(fileName);
      var file = resultFile[i];
      var reader = new FileReader();
      reader.onload = (function (file) {
        if (!hasinvalidcharacters) {
          var format = /[`!@#$%^&*()+\=\[\]{};':"\\|,<>\/?~]/;
          hasinvalidcharacters = format.test(file.name);
        }

        return function (e) {
          fileInfos.push({
            name: file.name,
            content: e.target.result,
          });
        };
      })(file);
      reader.readAsArrayBuffer(file);

      this.setState({
        hasinvalidcharacters: hasinvalidcharacters,
      });
    }
    this.setState({ fileInfos });
    console.log(fileInfos);
  }

    public onChoiceChange(ev, option: any): void {  
      this.setState({[ev.target.name]: option.key});  
      console.log(this.state)
      debugger;
    }

    public getCurrentLoginUser(){
      debugger;
      return new Promise<any[]>((resolve,reject)=>{
          sp.web.currentUser.get().then((r: any) => {
            debugger; 
            this.setState({EmpName: r.LoginName},()=>{
              this.myServices.getUserDetails(r.LoginName).then((itemColl)=>{
                console.log(itemColl);
                 debugger;
                this.setState({itemColl:itemColl},()=>{
                  console.log(this.state);
                  debugger;
                  resolve(this.state.EmpName),
                  reject("Error")
                });
              })
            }); 
          });        
      })
    }

    private async getLookUpAndDropdownFromList(bindType){  
      await this.myServices.BindDropDown('TestLookUp','Title,ID').then(res=>{
        debugger;
        this.setState({LookupColumnColl:res})
        debugger;
      })
    }

    private _onChange(event){
      this.setState({[event.target.name]: event.target.value});
      debugger;
    }

    _onChangeCheckOld(ev: React.FormEvent<HTMLElement>, isChecked: boolean) {
      debugger;
      //console.log("The option has been changed to ${isChecked}.");
      console.log(`The option ${ev.currentTarget.title} has been changed to ${isChecked}`);
      this.setState({[ev.currentTarget.title]: isChecked});
    }

    public _onChangeCheck = (ev: React.FormEvent<HTMLElement>,isChecked: boolean) => {
      if (isChecked) {
        if (this.state.SelectedTechnology === null) {
          let OCASChangeTypeArray: string[] = [];
          OCASChangeTypeArray.push(ev.currentTarget.title);
          this.setState({ SelectedTechnology: OCASChangeTypeArray });
        } else {
          let OCASChangeTypeArray = this.state.SelectedTechnology;
          OCASChangeTypeArray.push(ev.currentTarget.title);
          this.setState({ SelectedTechnology: OCASChangeTypeArray });
        }
      } else {
        var OCASChangeTypeArray = this.state.SelectedTechnology;
        this.setState({
          SelectedTechnology: OCASChangeTypeArray.filter(
            (e) => e !== ev.currentTarget.title
          ),
        });
      }
    };

    public _getPeoplePickerItems = async (items: any[]) => {
      debugger;
      this.setState({ FormAssignedTo: items[0].id });
      this.setState({ FormAssignedToId: items[0].id });
    };

    public _getPeoplePickerItemsM = async (items: any[]) => {
      debugger;
      //this.setState({ FormAssignedTo: items[0].id });

      let getSelectedUsers =[]// this.state.FormAssignedMultipleID;  
      for (let item in items) {  
        getSelectedUsers.push(items[item].id);  
      }  
      this.setState({ FormAssignedMultiple: getSelectedUsers }); 
      this.setState({ FormAssignedMultipleID: getSelectedUsers }); 
    };
 

    saveRecord=()=>{
      console.log(this.state);
      debugger;
      var item={
        Title:this.state.FormTitle,
        Description:this.state.FormDescription,
        Department:this.state.FormDepartment,
        PublishDate:this.state.FormPublishDate,
        Technology: {
          results: this.state.SelectedTechnology,
        },
        IsActive:false,
        Sex:this.state.FormSex,
        AssignedToId:this.state.FormAssignedTo,
        AssignedMultipleTeamId:{ results: this.state.FormAssignedMultiple }
      };
      this.myServices.AddRequestDetails("MyTestList",item).then((res:any)=>{
       debugger
        alert(res.Id +' was created successfully.')
      })
    }

    addListItemAttachment=()=>{
      var item={
        Title:"Test Add Attachment"
      }
      this.myServices.AddListItemAttachment("MyTestList",item).then(async(res:any)=>{
        debugger;
        if (this.state.fileInfos != null) {
          let updateAttachments = await sp.web.lists
            .getByTitle("MyTestList")
            .items.getById(res.Id)
            .attachmentFiles.addMultiple(this.state.fileInfos);
          console.log(updateAttachments, "UpdateAttachments");
        }

         alert(res.Id +' attachment uploaded successfully.')
       })
    }

    updateRecord=()=>{
      console.log(this.state);
      debugger;
      var item={
        Title:this.state.FormTitle,
        Description:this.state.FormDescription,
        Department:this.state.FormDepartment,
        PublishDate:this.state.FormPublishDate,
        Technology: {
          results: this.state.SelectedTechnology,
        },
        IsActive:false,
        Sex:this.state.FormSex,
        AssignedToId:this.state.FormAssignedToId,
        AssignedMultipleTeamId:{ results: this.state.FormAssignedMultipleID }
      };
      this.myServices.UpdateRequestDetails("MyTestList",item,parseInt(this.state.CurrentItemID)).then((res:any)=>{
       debugger
        alert(res.Id +' was updated successfully.')
      })
    }

    

    public render(): React.ReactElement<IDashboardState> {
      let departmentList = DepartmentOptions.length > 0
        && DepartmentOptions.map((item, i) => {
        return (
          <option key={i} value={item.text}>{item.text}</option>
        )}, this);

        let lookupList = this.state.LookupColumnColl.length > 0
        && this.state.LookupColumnColl.map((item, i) => {
        return (
          <option key={i} value={item.text}>{item.text}</option>
        )}, this);

        

      return (
        <div  className="w100">
          <h3>List Add Attachment</h3>

          <label >Attachments:</label>
            <input
              type="file"
              multiple={true}
              id="file"
              onChange={this.addFile.bind(this)}
            />
            <PrimaryButton onClick={this.addListItemAttachment}>Submit</PrimaryButton>

          <h3>Add/Edit List operation</h3>
          <TextField
              label="Title"
              id="txtTitle"
              required={false}
              multiline={false}
              value={this.state.FormTitle}
              name='FormTitle'
              onChange={this._onChange}
              />
          
            <TextField
              label="Description"
              id="txtDescription"
              required={false}
              multiline={true}
              value={this.state.FormDescription}
              name='FormDescription'
              onChange={this._onChange}
              />

            <Label>Select Department</Label>
            <select value={this.state.FormDepartment} onChange={(e) => this.setState({FormDepartment: e.target.value})}>
              {departmentList}
            </select>

            <Label>Publish Date</Label>
            {/* <DatePicker maxDate={new Date()} allowTextInput={false} strings={DatePickerStrings} */}
            <DatePicker allowTextInput={false} strings={DatePickerStrings}
              value={this.state.FormPublishDate}
              onSelectDate={(e) => { this.setState({FormPublishDate: e });}}
              ariaLabel="Select a date" formatDate={FormatDate} />

            <Label>Select Technology(multiselect)</Label>
            {
              checkOptions.map((checkBoxItem: any) => {
                return (
                    <Checkbox label={checkBoxItem.Title} checked={this.state.SelectedTechnology.includes(
                      checkBoxItem.Title
                    )}  title={checkBoxItem.Title} onChange={this._onChangeCheck} />
                  );
                })
            }
      
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

              <PeoplePicker   
                context={this.props.context}  
                titleText="Project Members(Multiple Assigned People)"  
                personSelectionLimit={5}  
                showtooltip={true}
                disabled={false}
                onChange={this._getPeoplePickerItemsM}
                defaultSelectedUsers={this.state.FormAssignedMultiple}
                showHiddenInUI={false}  
                ensureUser={true}  
                principalTypes={[PrincipalType.User]}  
                resolveDelay={1000} />
 
              <ChoiceGroup id="sex" //defaultSelectedKey="Male" 
                options={SexOptions} 
                onChange={this.onChoiceChange} name='FormSex'
                label="Sex"
                selectedKey={this.state.FormSex}
                required={true} />

              <Label>Look Up Column</Label>
              <select value={this.state.FormLookUp} name="FormLookUp" 
              //onChange={(e) => this.setState({FormLookUp: e.target.value})}
              onChange={this._onChange}
              >
                {lookupList}
              </select>
              <PrimaryButton onClick={this.saveRecord}>Submit</PrimaryButton>
              <PrimaryButton onClick={this.updateRecord}>Update</PrimaryButton>
          <Link to={`/`}>Back</Link>
        </div>
      );
    }
  }
  
