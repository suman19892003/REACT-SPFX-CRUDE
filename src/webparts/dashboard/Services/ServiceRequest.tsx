import { sp } from '@pnp/sp/presets/all';
import { IDropdownOption } from 'office-ui-fabric-react/lib/components/Dropdown';
import { SPHttpClient, SPHttpClientResponse, SPHttpClientConfiguration } from '@microsoft/sp-http';
import { escape, find } from "@microsoft/sp-lodash-subset";
export class RequestServices {
    /*

public GetLegacyServiceDetails(listName: string,pnpSelect: string,pnpExpand: string,pnpFilter: string):Promise<any[]>{
    return new Promise<any[]>((resolve,reject)=>{
        pnp.sp.web.lists.getByTitle(listName).items
        .filter(pnpFilter)
        .select(pnpSelect)
        .expand(pnpExpand)
        .get()
        .then((results: any[])=>{
            // if(listName == "ME%20Leader"){resolve(results[0].ME_x0020_Leader.EMail);}else{
            resolve(results);
        // }
    },(error:any[])=>{
        $('#loading').hide();
        console.log(error);
        reject('Error Occured');
    })
    })
}
public updateDocumentLibrary(MainIndex) {
    try{
    pnp.sp.web.lists.getByTitle('LegacyServiceAttachments').items
    .getById(MainIndex)
    .delete()
          // .then(function(newRequest) {
            .then((newRequest) => {
              alert('Document deleted successfully..');
              console.log(newRequest);
          })
          // (error: any): void => {
            .catch((error)=>{
            //   $('#loading').hide();
            $('#loading').hide();
              console.log("Error in updateDocumentLibrary : ", error.message);  
            // this.setState({ status: "Not Ready" });
          });
        } catch (error) {  
            $('#loading').hide();
          console.log("Error in updateDocumentLibrary : ", error.message);  
        //   this.setState({ status: "Not Ready" });
        }
}
public UpdateDocument(listName:string,DocType:string,DocumentArray:any,LegacyServiceId:string):Promise<string>{
    // try{
        return new Promise<string>((resolve,reject)=>{
   var array = DocumentArray;
   for (var i = 0; i < array.length; i++) {
    //  if (array[i] != undefined) {
       var temp = array[i].toString().split('|');
       pnp.sp.web.lists.getByTitle(listName).items.getById(parseInt(temp[1])).update({
        FormID: LegacyServiceId,
        BeforeAfter:DocType
       }).then((results:any)=>{
         console.log("Add New document in IOM_CRM_Attachments :  " + LegacyServiceId); 
         resolve(results);
    },(error:any)=>{
        console.log(error);
        reject('Error Occured');
    });
        }
    });
}
public GetLegacyServiceDocuments(listName:string,LegacyServiceId:string):Promise<any[]>{
    try{
    return new Promise<any[]>((resolve,reject)=>{
    pnp.sp.web.lists.getByTitle(listName).items
    .filter("FormID eq "+LegacyServiceId+"")
    .select("EncodedAbsUrl","File/Name","Id","BeforeAfter")
    .expand("File").get().then((items: any[]) => {
      var str = [];
      var DocsArray = [];
    //   for (var i = 0; i < items.length; i++) {
    //     if (items[i] != undefined) {
    //       DocsArray.push(items[i].Id);
    //     str.push(<li >&nbsp; <a href={items[i].EncodedAbsUrl} id={items[i].Id} target="_blank">{items[i].File.Name}</a></li>);
    //       //  str.push(<li key={tempx[0]} onClick={this.onChangeDeleteDocument.bind(this)} data-id={tempx[1]}> Uploaded File : {tempx[0]} - <a >Delete </a></li>);
    //     }
    //   }
      //  this.state.getData.push(<ul>{str}</ul>);
    //   this.setState({getDocumentData:str});
    //   this.setState({PreviousDocs:DocsArray});
    resolve(items);
        },(error:any[])=>{
            console.log(error);
            reject('Error Occured');
        })
    })
  } catch (error) {  
    $('#loading').hide();
    console.log("Error in GetIOMDocuments : ", error.message);  
    // this.setState({ status: "Not Ready" });
  } 
}
*/

public getUserProperties(webUrl, context, userLogin) {
    return new Promise<any[]>((resolve,reject)=>{
    debugger;
  //let apiUrlold = this.props.webURL + "/_api/SP.UserProfiles.PeopleManager/GetUserProfilePropertyFor(accountName=@v)?@v='" + encodeURIComponent("i:0#.f|membership|") + userEmail + "'";  
  let apiUrl = webUrl + "/_api/SP.UserProfiles.PeopleManager/GetPropertiesFor(accountName=@v)?@v='" + encodeURIComponent('i:0#.f|membership|amitd@smartek21.com') +"'";
  //let apiUrl = this.props.webURL + "/_api/SP.UserProfiles.PeopleManager/GetPropertiesFor(accountName=@v)?@v='" + encodeURIComponent('"'+userLogin+'"') +"'";
  let httpClient: SPHttpClient = context.spHttpClient;  
     httpClient.get(apiUrl, SPHttpClient.configurations.v1).then(response => {
      debugger;
       response.json().then(responseJson => {
          debugger;
          resolve(responseJson.UserProfileProperties);         
            },(error:any[])=>{
                console.log(error);
                reject('Error Occured');
            });
        })
    })
}

public getCurrentUserDetails(webUrl, context) {
    return new Promise<any[]>((resolve,reject)=>{
    debugger;
  //let apiUrlold = this.props.webURL + "/_api/SP.UserProfiles.PeopleManager/GetUserProfilePropertyFor(accountName=@v)?@v='" + encodeURIComponent("i:0#.f|membership|") + userEmail + "'";  
  let apiUrl = webUrl + "/_api/SP.UserProfiles.PeopleManager/GetMyProperties";
  //let apiUrl = this.props.webURL + "/_api/SP.UserProfiles.PeopleManager/GetPropertiesFor(accountName=@v)?@v='" + encodeURIComponent('"'+userLogin+'"') +"'";
  let httpClient: SPHttpClient = context.spHttpClient;  
     httpClient.get(apiUrl, SPHttpClient.configurations.v1).then(response => {
      debugger;
       response.json().then(responseJson => {
          debugger;
          resolve(responseJson.UserProfileProperties);         
            },(error:any[])=>{
                console.log(error);
                reject('Error Occured');
            });
        })
    })
}

public AddRequestDetailsRESTApi(webUrl, context,body):Promise<string>{
    debugger;
    return new Promise<string>((resolve,reject)=>{
     let apiUrl = webUrl + "/_api/web/lists/getbytitle('MyList')/items";
     let httpClient: SPHttpClient =  context.spHttpClient;  
     httpClient.post(apiUrl, SPHttpClient.configurations.v1,
        
        {  
            headers: {  
              'Accept': 'application/json;odata=nometadata',  
              'Content-type': 'application/json;odata=nometadata',  
              'odata-version': ''  
            },  
            body: JSON.stringify(body)  
          }
        
        )
        .then((response: SPHttpClientResponse): Promise<any> => {  
            return response.json();  
          }) 
          .then((item: any): void => {  
            alert(`Item '${item.Title}' (ID: ${item.Id}) successfully created`);  
          }, (error: any): void => {  
            alert('Error while creating the item: ' + error);  
          });
        })
  }


  //Actual Code start

  //List Dashboard

  public GetRequestDetails(listname){
    debugger;
    return new Promise<any[]>((resolve,reject)=>{
      sp.web.lists.getByTitle(listname).items
      .select("Title","Department","PublishDate","Technology", "Id",'Description','Sex','AssignedTo/Name','AssignedTo/ID','AssignedTo/EMail','AssignedTo/Title', "LookUpColumn/ID",'LookUpColumn/Title')
      .expand("AssignedTo","LookUpColumn")
      .get().then(response => {
        debugger;
        resolve(response);         
        },(error:any[])=>{
            console.log(error);
            reject('Error Occured');
        })
    })
  }

  public GetMyApprovals(userEmail) {
  debugger;
    return new Promise<any[]>((resolve,reject)=>{
    sp.web.lists.getByTitle("MyTestList").items
    .filter(`AssignedTo/EMail eq '${encodeURIComponent(userEmail)}'`)
    .select("Title","Department","PublishDate","Technology", "Id",'Description','Sex','AssignedTo/Name','AssignedTo/ID','AssignedTo/EMail','AssignedTo/Title', "LookUpColumn/ID",'LookUpColumn/Title')
    .expand("AssignedTo","LookUpColumn")
    .get().then(p => {  
        var itemColl = p;
        debugger;
        resolve(p) 
        },err=>{
            reject(err)
        });
    });
  }

  public GetFilteredRecordLookUpColumn(filterVal){
    debugger;
    return new Promise<any[]>((resolve,reject)=>{
    sp.web.lists.getByTitle("MyTestList").items
    .filter(`LookUpColumn/Title eq '${encodeURIComponent(filterVal)}'`)
    .select("Title","Department","IsActive","PublishDate","Technology", "Id",'Description','Sex','AssignedTo/Name','AssignedTo/ID','AssignedTo/EMail','AssignedTo/Title', "LookUpColumn/ID",'LookUpColumn/Title')
    .expand("AssignedTo","LookUpColumn")
    .get().then(p => {  
        var itemColl = p;
        debugger;
        resolve(p) 
        },err=>{
            reject(err)
        });
    });
  }

  public GetFilteredRecordDropdown(titleVal){
    return new Promise<any[]>((resolve,reject)=>{
    debugger
    return sp.web.lists.getByTitle('MyTestList').items
    .filter(`Department eq '${encodeURIComponent(titleVal)}'`)
    .select("Title","Department","IsActive","PublishDate","Technology", "Id",'Description','Sex','AssignedTo/Name','AssignedTo/ID','AssignedTo/EMail','AssignedTo/Title', "LookUpColumn/ID",'LookUpColumn/Title')
    .expand("AssignedTo","LookUpColumn")
    .get().then(p => {  
        var itemColl = p;
        debugger;
        resolve(p) 
        },err=>{
            reject(err)
        });
    });
  }

  public GetFilteredRecordCheckBox(titleVal){
    return new Promise<any[]>((resolve,reject)=>{
    debugger
    return sp.web.lists.getByTitle('MyTestList').items
    .filter(`Technology eq '${encodeURIComponent(titleVal)}'`)
    .select("Title","Department","IsActive","PublishDate","Technology", "Id",'Description','Sex','AssignedTo/Name','AssignedTo/ID','AssignedTo/EMail','AssignedTo/Title', "LookUpColumn/ID",'LookUpColumn/Title')
    .expand("AssignedTo","LookUpColumn")
    .get().then(p => {  
        var itemColl = p;
        debugger;
        resolve(p) 
        },err=>{
            reject(err)
        });
    });
  }

  public GetFilteredRecordDate(titleVal){
    let timeNow = new Date().toISOString();
    return new Promise<any[]>((resolve,reject)=>{
    debugger
    return sp.web.lists.getByTitle('MyTestList').items
    .filter(`PublishDate le datetime'${timeNow}'`)
    .select("Title","Department","IsActive","PublishDate","Technology", "Id",'Description','Sex','AssignedTo/Name','AssignedTo/ID','AssignedTo/EMail','AssignedTo/Title', "LookUpColumn/ID",'LookUpColumn/Title')
    .expand("AssignedTo","LookUpColumn")
    .get().then(p => {  
        var itemColl = p;
        debugger;
        resolve(p) 
        },err=>{
            reject(err)
        });
    });
  }

  public GetMultipleFilteredRecord(titleVal,deptVal){
    let timeNow = new Date().toISOString();
    return new Promise<any[]>((resolve,reject)=>{
    debugger
    return sp.web.lists.getByTitle('MyTestList').items
    .filter(`PublishDate le datetime'${timeNow}'`)
    .filter(`Title eq '${encodeURIComponent(titleVal)}'`)
    .filter(`Department eq '${encodeURIComponent(deptVal)}'`)
    .select("Title","Department","IsActive","PublishDate","Technology", "Id",'Description','Sex','AssignedTo/Name','AssignedTo/ID','AssignedTo/EMail','AssignedTo/Title', "LookUpColumn/ID",'LookUpColumn/Title')
    .expand("AssignedTo","LookUpColumn")
    .get().then(p => {  
        var itemColl = p;
        debugger;
        resolve(p) 
        },err=>{
            reject(err)
        });
    });
  }

  public GetMyApprovalsRestAPI(context) {
    return new Promise<any[]>((resolve,reject)=>{
      let apiUrl = "https://libertymutual.sharepoint.com/sites/SmartekClaimsform_dev/_api/web/Lists/getbytitle('MyTestList')/items?$filter=AssignedTo/EMail eq '" + context.pageContext.user.email + "'&$select=Title,PublishDate,Technology,Sex,Department,Description,LookUpColumn/Title,AssignedTo/ID,AssignedTo/EMail,AssignedTo/Title&$expand=AssignedTo,LookUpColumn";
      let httpClient: SPHttpClient = context.spHttpClient;  
          httpClient.get(apiUrl, SPHttpClient.configurations.v1).then(response => {
          debugger;
            response.json().then(responseJson => {
            debugger;
            resolve(responseJson.value);         
              },(error:any[])=>{
                  console.log(error);
                  reject('Error Occured');
              });
          })
      })
  }
    

  public BindDropDown(listName: string,pnpSelect: string):Promise<IDropdownOption[]>{
    return new Promise<IDropdownOption[]>(async(resolve,reject)=>{
        sp.web.lists.getByTitle(listName).items
        //.filter(pnpFilter)
        .select(pnpSelect)
        //.expand(pnpExpand).top(1000).orderBy("ID", true)
        .orderBy("ID", true)
        .get()
        .then((results:any)=>{
            var DropdownOption : IDropdownOption[] = [];         
            results.map((result)=>{
            DropdownOption.push({key:result.Id,text:result.Title});
            });
            resolve(DropdownOption);
        },(error:any)=>{
            //$('#loading').hide();
            console.log(error);
            reject('Error Occured');
        })
        .catch(function(data) {
            //$('#loading').hide();
            console.log(data.data.responseBody["odata.error"].message.value);
        });

    });
    
}

//List Add Form

getUserDetails = (userId) => {
    let userDetailedInformation:any=[];
    return new Promise<any>((resolve,reject)=>{
    sp.profiles.getPropertiesFor(userId).then((profile) => {
      debugger;
      sp.profiles
        .getPropertiesFor(
          find(profile.UserProfileProperties, ["Key", "Manager"]).Value
        )
        .then((Managerprofile) => {
          debugger;
          //this.setState({
            userDetailedInformation= [
              {
                Employee_Name: find(profile.UserProfileProperties, [
                  "Key",
                  "PreferredName",
                ]).Value,
                Employee_Email: find(profile.UserProfileProperties, [
                  "Key",
                  "UserName",
                ]).Value,
                OfficeName: find(profile.UserProfileProperties, [
                  "Key",
                  "Office",
                ]).Value,
                DepartmentName: find(profile.UserProfileProperties, [
                  "Key",
                  "Department",
                ]).Value,                
                Fax: find(profile.UserProfileProperties, ["Key", "Fax"]).Value,
                Phone: find(profile.UserProfileProperties, ["Key", "CellPhone"])
                  .Value,               
                Manager_Name: find(Managerprofile.UserProfileProperties, [
                  "Key",
                  "PreferredName",
                ]).Value,
                Manager_Email: find(Managerprofile.UserProfileProperties, [
                  "Key",
                  "UserName",
                ]).Value
              }
            ];
            debugger;
            resolve(userDetailedInformation);
          },(err)=>{
              reject(err)
          });
        //});
        });
    });
  };

  public AddRequestDetails(listname,body):Promise<string>{
    debugger;
    return new Promise<string>((resolve,reject)=>{
        sp.web.lists
          .getByTitle("MyTestList")
          .items.add(body)
          .then(async (response: any) => {
              debugger;
              resolve(response.data)
            },()=>{
              reject("Error")
          })
    })
  }

  public AddListItemAttachment(listname,body):Promise<string>{
    debugger;
    return new Promise<string>((resolve,reject)=>{
        sp.web.lists
          .getByTitle("MyTestList")
          .items.add(body)
          .then(async (response: any) => {
              debugger;
              resolve(response.data)
            },()=>{
              reject("Error")
          })
    })
  }

  public GetRequestDetailsId(itemid) {
      debugger;
      return new Promise<any>((resolve,reject)=>{
        sp.web.lists.getByTitle("MyTestList").items.getById(itemid)
        .select("Title","Department","PublishDate","Technology", "Id",'Description','Sex','AssignedTo/Name','AssignedTo/ID','AssignedTo/EMail','AssignedMultipleTeam/ID','AssignedMultipleTeam/EMail','AssignedMultipleTeam/Title','AssignedTo/Title', "LookUpColumn/ID",'LookUpColumn/Title')
        .expand("AssignedTo","AssignedMultipleTeam","LookUpColumn")
        .get().then(response => {
          debugger;
          resolve(response);         
          },(error:any[])=>{
              console.log(error);
              reject('Error Occured');
          })
      })
  }

  public UpdateRequestDetails(listName: string,Body,ItemId:number):Promise<string>{
    return new Promise<string>((resolve,reject)=>{
        sp.web.lists.getByTitle("MyTestList").items
        .getById(ItemId)
        .update(Body)
        .then((results:any)=>{
            resolve(results);
        },(error:any)=>{
            console.log(error);
            reject('Error Occured');
        });
    });

    }

    //Library Code Starts

    getAllFileFromLibrary(){
        return new Promise<any>((resolve,reject)=>{
        debugger;
        let array=[]
        sp.web.lists.getByTitle('MyTestLibrary').items.select('Id','FileRef','File','Title','AssignedTo/Title','AssignedTo/ID','AssignedTo/EMail')
        .expand('File','AssignedTo').get().then(file=>{
          debugger;
          console.log(file);
          file.map((item)=>{ 
              debugger;      
            console.log(item.FileRef);
            if(item.File != "" && item.File != null && item.File != undefined){
            array.push({
                //FileName:item.File.Name,
                FileName:item.File == "" || item.File == null || item.File == undefined?"":item.File.Name,
                FileURL:item.FileRef,
                ItemID:item.ID,
                Title:item.Title,
                AssignedTo:item.AssignedTo == "" || item.AssignedTo == null || item.AssignedTo == undefined?"":item.AssignedTo.Title,

            }); 
            }      
          });
          resolve(array)
        },(err)=>{reject(err)})
      })
    }

    public GetMyApprovalsFromLibraryRestAPI(context) {
        return new Promise<any[]>((resolve,reject)=>{
          let apiUrl = "https://libertymutual.sharepoint.com/sites/SmartekClaimsform_dev/_api/web/Lists/getbytitle('MyTestLibrary')/items?$filter=AssignedTo/EMail eq '" + context.pageContext.user.email + "'&$select=Id,FileRef,File,Title,AssignedTo/Title,AssignedTo/ID,AssignedTo/EMail&$expand=AssignedTo,File";
          let httpClient: SPHttpClient = context.spHttpClient;  
              httpClient.get(apiUrl, SPHttpClient.configurations.v1).then(response => {
              debugger;
                response.json().then(responseJson => {
                debugger;
                resolve(responseJson.value);         
                  },(error:any[])=>{
                      console.log(error);
                      reject('Error Occured');
                  });
              })
          })
      }

      getFileFromFolder(){
        return new Promise<any>((resolve,reject)=>{
        debugger;
        let array=[]

        // sp.web.getFolderByServerRelativeUrl("/sites/SmartekClaimsform_dev/MyTestLibrary/Test")
        // //.getItem('Id','FileRef','File','Title','AssignedTo/Title','AssignedTo/ID','AssignedTo/EMail')
        // .files
        // .expand('Files/ListItemAllFields','AssignedTo') // For Metadata extraction
        // .select('Id','FileRef','File','Title','AssignedTo/Title','AssignedTo/ID','AssignedTo/EMail','Name')              // Fields to retrieve
        // .get().then(function(item) {
        //     console.log(item);
        // });

        sp.web.getFolderByServerRelativeUrl("/sites/SmartekClaimsform_dev/MyTestLibrary/Test").files.get().then(files => {
            for (var i = 0; i < files.length; i++) {
                var _ServerRelativeUrl = files[i].ServerRelativeUrl;
                var fileN = files[i].Name;
                sp.web.getFileByServerRelativeUrl(_ServerRelativeUrl).getItem().then(item=> {
                    debugger;
                    console.log(item);
                    let itemColl:any=item
                    // itemColl.map((item)=>{ 
                    debugger;      
                      console.log(itemColl.FileRef);
                      array.push({
                          //FileName:item.File.Name,
                          FileName:fileN,
                          FileURL:_ServerRelativeUrl,
                          ItemID:itemColl.ID,
                          Title:itemColl.Title,
                          AssignedTo:itemColl.AssignedTo == "" || itemColl.AssignedTo == null || itemColl.AssignedTo == undefined?"":itemColl.AssignedTo.Title
                      });       
                    // });
                    resolve(array)
                })
            }
        },(err)=>{
            reject(err)
        });
      })
    }

    deleteFileFromLibrary(fileURL){
        sp.web.getFileByServerRelativeUrl(fileURL).
        recycle().then(function(data){
            console.log(data);
            alert('deleted')
        });
    }
}