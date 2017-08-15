
import * as pnp from 'sp-pnp-js';

import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './GetUserProfileProperties.module.scss';


import * as strings from 'getUserProfilePropertiesStrings';
import { IGetUserProfilePropertiesWebPartProps } from './IGetUserProfilePropertiesWebPartProps';


export default class GetUserProfilePropertiesWebPart extends BaseClientSideWebPart<IGetUserProfilePropertiesWebPartProps> {




public GetUserProperties(query): any {


  console.log(query) ;

/*
  pnp.sp.web.siteUsers.getByEmail("m.a.goss@bristol.ac.uk").get().then(function(result) {
    var userInfo = "";
    console.log(result) ;
    var prop ;
    for (prop in result) {
        userInfo += prop + " : " + result[prop] + "<br/>";
    }
    //document.getElementById("sample").innerHTML = userInfo;
});
*/



//Get all user profile properties of given user
var result = {} ; 
var loginName ; 
var filter = "Title eq '"+query + "'" ;
console.log(filter); 
//pnp.sp.profiles.getPropertiesFor("i:0#.f|membership|ciago@bristol.ac.uk").then(function(result) {
//
  pnp.sp.web.siteUsers.filter(filter).get().then(function(result) {


      console.log(result[0].LoginName) ;
      loginName = result[0].LoginName ;
    
  

    pnp.sp.profiles.getPropertiesFor(loginName).then(function(result) {
   
    var userInfo ;
    var prop = "";
    var userProperties = result;
    var email ;
    var phone ; 
    var test ;
    
   
   
    
     for (prop in result) {
        //userInfo += prop + " : " + result[prop] + "<br/>";
        if (prop == "LoginName")
          {
            loginName = result[prop] ;
            console.log(loginName) ;
            
          }

          if (prop == "UserProfileProperties") {
            var userProfileProp = result[prop] ;
            console.log(userProfileProp) ;

          for(var i=0; i< userProfileProp.length; i++) {
              //console.log(users[i]);
              var userProp = userProfileProp[i] ;
              console.log(userProp)
              if (userProp.Key == "WorkEmail")
                {
                  console.log(userProp)
                  email = userProp.Value ;
                } 
              if (userProp.Key == "WorkPhone")
                {
                  console.log(userProp)
                  phone = userProp.Value ;
                }
            }
          }
          //console.log(result[prop][0]) ;

          
          
  //userPropertyValues += property.Key + " - " + property.Value + "<br/>";
          }

         


    
    
    document.getElementById("spUserProfileProperties").innerHTML = email + "<br>" + phone ;
}).catch(function(err) {
    console.log("Error: " + err);
});

});

/*
  pnp.sp.profiles.myProperties.get().then(function(result) {
  var userProperties = result.UserProfileProperties;
  console.log(userProperties) ;

  var firstName ;
  var lastName ;
  var name ;
  var email ;
  var number ;
  var department ;
  var title ;
  var picture ;

  //console.log(userProperties) ;
  userProperties.forEach(function(property) {
  //userPropertyValues += property.Key + " - " + property.Value + "<br/>";



  if (property.Key == "FirstName")
  {
    firstName = property.Value ;
  }
  if (property.Key == "LastName")
  {
    lastName = property.Value ;
  }
  if (property.Key == "WorkPhone")
  {
    number = property.Value ;
  }
  if (property.Key == "UserName")
  {
    email = property.Value ;
  }
  if (property.Key == "Department")
  {
    department = property.Value ;
  }
  if (property.Key == "Title")
  {
    title = property.Value ;
  }
  if (property.Key == "PictureURL")
  {
    picture = property.Value;
  }


  });
  name = firstName + " " + lastName + "<br/>" ;
  number = "Phone Number: +" +  number + "<br/>" ;
  email = "email address: " + email + "<br/>" ;
  department = "Department: " + department + "</br>" ;
  title = "Title " + title + "</br>" ;
  picture = '<img style="float:right" class="displayPic" src="'+picture+'" alt="display Picture" height="64" width="64">' ;

  document.getElementById("spUserProfileProperties").innerHTML = name + picture + number + email + department + title  ;
  }).catch(function(error) {
   console.log("Error: " + error);
});
*/
 }




  public render(): void {

     this.domElement.innerHTML = `
     <div class="${styles.helloWorld}">
  <div class="${styles.container}">
    <div class="ms-Grid-row ms-bgColor-themeDark ms-fontColor-white ${styles.row}">
      <div class="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
        <span class="ms-font-xl ms-fontColor-white" style="font-size:28px">Search for a user in SharePoint..</span>
        <br><br>
        <input type="text" name="searchInput" id="searchInput" placeholder="Enter full name here">
        <br><br>
        <input type="submit" name="search" id="search">
        <!--<p class="ms-font-l ms-fontColor-white" style="text-align: left">Demo : Retrieve User Profile Properties</p>-->
      </div>
    </div>
    <div class="ms-Grid-row ms-bgColor-themeDark ms-fontColor-white ${styles.row}">
    <div style="background-color:Black;color:white;text-align: center;font-weight: bold;font-size:18px;">User Profile Details</div>
    <br>
<div id="spUserProfileProperties" />
    </div>
  </div>
</div>`;

var search = document.getElementById('search');

  search.addEventListener('click', function () {
  var query = (<HTMLInputElement>document.getElementById('searchInput')).value;
  //console.log(query) ;
   var user = new GetUserProfilePropertiesWebPart
    user.GetUserProperties(query) ;
    
  });
  
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
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

