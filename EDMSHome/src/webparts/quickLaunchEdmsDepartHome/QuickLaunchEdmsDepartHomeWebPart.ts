import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart,IPropertyPaneConfiguration,PropertyPaneTextField} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';
import * as strings from 'QuickLaunchEdmsDepartHomeWebPartStrings';
import "jquery";
import { readItems, checkUserinGroup } from "../../commonJS";
declare var $;

export interface IQuickLaunchEdmsDepartHomeWebPartProps {
  description: string;
}

export default class QuickLaunchEdmsDepartHomeWebPart extends BaseClientSideWebPart<IQuickLaunchEdmsDepartHomeWebPartProps> {
 // USER GROUP VALIDATION
 
 userflag: boolean = false;
 public render(): void {
   var _this = this;
   checkUserinGroup("Quick Launch", this.context.pageContext.user.email, function (result) {
     if (result == 1) {
       _this.userflag = true;
       _this.QuickLaunchDisplay(_this.userflag);
     }
   })
 }

 // STRUCTURE

 QuickLaunchDisplay(userflag){
   var webURL = this.context.pageContext.web.absoluteUrl;
   this.domElement.innerHTML = `
   <section class="vertical-menu">
     <div class="panel-group" id="accordionMenu" role="tablist" aria-multiselectable="true">
       <div class='panel panel-default'>
         <div style="background-color: #E42313;" class='panel-heading' role='tab' id="addNew">
           <h4  class='panel-title'>
             <a id='quickLaunchTitleId' style="color:#fff;" target='_blank' href='../Pages/ListView.aspx?CName=Quick Launch'><i class='icon-new'></i>Customize</a>
           </h4>
         </div>
       </div>
     </div>
   </section>
   `;
   this.displayQuickLinks(userflag);
  //  var divHeight = $('#right-side').height(); 
  //  $('.vertical-menu').css('min-height', divHeight+'px');

 }

 // BIND DATA TO HTML

 async displayQuickLinks(userflag){
   let Renderhtml="";
   let linkListItems = await readItems("Quick Launch", ["Title", "LinkURL"],5, "Modified", "Display", 1);
   let linksListItemsLength = linkListItems.length;
   if(linksListItemsLength == 0){
     Renderhtml += "<div class='panel panel-default'>" +
                       "<div class='panel-heading' role='tab' id='NoItemToDisp'>" +
                         "<h4 class='panel-title'>No Item To Display </h4>" +
                       "</div>"
                     "</div>"
   }else{
     for (var i = 0 ; i < linksListItemsLength; i ++ ){
       Renderhtml += "<div class='panel panel-default'>" +
                       "<div class='panel-heading' role='tab' id="+ linkListItems[i].Title +">" +
                         "<h4 class='panel-title'>" +
                           "<a  target='_blank' href='" + linkListItems[i].LinkURL.Url + "'><i class='icon-file'></i>" +  linkListItems[i].Title +"</a>" +
                         "</h4>" +
                       "</div>"
                     "</div>"
     }
   }
   $('#accordionMenu').append(Renderhtml);
   if(userflag == true){
      $('#quickLaunchTitleId').show();
    }else if(userflag == false){
      $('#quickLaunchTitleId').hide();
    }
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