import { Version } from '@microsoft/sp-core-library';
import { sp, Items, ItemVersion, Web } from "@pnp/sp";

import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import {
  Environment,
  EnvironmentType
 } from '@microsoft/sp-core-library';

 import {
  SPHttpClient,
  SPHttpClientResponse   
 } from '@microsoft/sp-http';

import { escape } from '@microsoft/sp-lodash-subset';

import styles from './GsvrAnnouncementsWebPart.module.scss';
import * as strings from 'GsvrAnnouncementsWebPartStrings';

export interface ISPLists {
  value: ISPList[];
 }

 export interface ISPList {
  Title: string; // this is the department name in the List
  Id: string;
  AnncURL:string;
  DeptURL:string;
  CalURL:string;
  a85u:string; // this is the LINK URL
 }

   //global vars
   var userDept = "";

export interface IGsvrAnnouncementsWebPartProps {
  description: string;
}

export default class GsvrAnnouncementsWebPart extends BaseClientSideWebPart<IGsvrAnnouncementsWebPartProps> {
  
  // get all the user properties
  getuser = new Promise((resolve,reject) => {
    // SharePoint PnP Rest Call to get the User Profile Properties
    return sp.profiles.myProperties.get().then(function(result) {
      var props = result.UserProfileProperties;
      var propValue = "";
      var userDepartment = "";
  
      props.forEach(function(prop) {
        //this call returns key/value pairs so we need to look for the Dept Key
        if(prop.Key == "Department"){
          // set our global var for the users Dept.
          userDept += prop.Value;
        }
      });
      return result;
    }).then((result) =>{
      this._getListData().then((response) =>{
        this._renderList(response.value);
      });
    });
  
  });

  public render(): void {
    this.domElement.innerHTML = `
      <div class="${ styles.gsvrAnnouncements }">
        <div class="${ styles.container }">
          <div class="${ styles.row }">
            <div class="${ styles.column }">
              <span class="${ styles.title }">Welcome to SharePoint!</span>
              <p class="${ styles.subTitle }">Customize SharePoint experiences using Web Parts.</p>
              <p class="${ styles.description }">${escape(this.properties.description)}</p>
              <a href="https://aka.ms/spfx" class="${ styles.button }">
                <span class="${ styles.label }">Learn more</span>
              </a>

              <h1>Department Annoucements</h1>
            <h3><div id="deptAnnouce"/></h3>

            </div>
          </div>
        </div>
      </div>`;
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  // main REST Call to the list...passing in the deaprtment into the call to 
  //return a single list item
  public _getListData(): Promise<ISPLists> {  
    return this.context.spHttpClient.get(`https://girlscoutsrv.sharepoint.com/_api/web/lists/GetByTitle('TeamDashboardSettings')/Items?$filter=Title eq '`+ userDept +`'`, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        return response.json();
      });
   }

   private _renderList(items: ISPList[]): void {
    let html: string = '';
    let libHTML: string ='';
  
    var siteURL = "";
    //list name
    var announcementsListName =  "";
    // items in the list
    var annoucementItems = "";
    var date = new Date();
    var strToday = "";
    var mm = date.getMonth()+1;
    console.log(mm);
    
    var dd = date.getDate();
    console.log(dd);

    var yyyy = date.getFullYear();
    console.log(yyyy);

    if(dd < 10){
      dd = 0 + dd;
      console.log(dd);
    }

    if(mm < 10){
      mm = 0 + mm;
      console.log(mm);
    }

    strToday = mm + "/" + dd + "/" + yyyy;
    console.log(strToday);
    
    items.forEach((item: ISPList) => {
      siteURL = item.DeptURL;
      announcementsListName = item.AnncURL;
    });
    //1st we need to override the current web to go to the department sites web
    const w = new Web("https://girlscoutsrv.sharepoint.com" + siteURL);
    
    // then use PnP to query the list
    w.lists.getByTitle(announcementsListName).items.filter("Expires ge '" + strToday + "'").top(1)
    .get()
    .then((data) => {
      console.log(data);

      for (var x = 0; x < data.length; x++){
        //console.log(data[x].URL);
        //this gets the HTTP URL of the hyper link
        console.log(data[x].Title);
        //this gets body of the annoucement
        console.log(data[x].Body);
        //date it expires
        console.log(data[x].Expires);

        annoucementItems += data[x].Title + '\r\n';
       // libHTML += `<p>${hrItems.toString()}</p>`;
      }
      document.getElementById("deptAnnouce").innerText = annoucementItems;
  }).catch(e => { console.error(e); });

    const listContainer: Element = this.domElement.querySelector('#ListItems');
    listContainer.innerHTML = html;
  }

  // this is required to use the SharePoint PnP shorthand REST CALLS
  public onInit():Promise<void> {
    return super.onInit().then (_=> {
      sp.setup({
        spfxContext:this.context
      });
    });
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
