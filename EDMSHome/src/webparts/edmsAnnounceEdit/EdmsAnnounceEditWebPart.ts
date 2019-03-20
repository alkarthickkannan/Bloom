import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';
import 'jquery';
import styles from './EdmsAnnounceEditWebPart.module.scss';
import * as strings from 'EdmsAnnounceEditWebPartStrings';
import {readItems,updateItem,GetQueryStringParams} from '../../commonJS';
require('../../ExternalRef/js/jquery.richtext.js');
import '../../ExternalRef/css/richtext.min.css';
declare var $;

export interface IEdmsAnnounceEditWebPartProps {
  description: string;
}

var ItemID;
export default class EdmsAnnounceEditWebPart extends BaseClientSideWebPart<IEdmsAnnounceEditWebPartProps> {

  public render(): void {
    var siteURL = this.context.pageContext.web.absoluteUrl;
  
    var strLocalStorage = GetQueryStringParams("CName");

    strLocalStorage = strLocalStorage.split("%20").join(' ');

    this.domElement.innerHTML = "<div class='breadcrumb'>" +
    "<ol>" +
    "<li><a href='" + siteURL + "/Pages/Home.aspx' title='Home'>Home</a></li>" +
    "<li><a class='pointer' id='breadTilte' title='" + strLocalStorage + "'>" + strLocalStorage + " List View</a></li>" +
    "<li><span>Add " + strLocalStorage + "</span></li>" +
    "</ol>" +
    "</div>" +
    "<div class='title-section'>" +
    "<div class='button-field save-button'>" +
    "<a  title='Save' class='addbutton pointer' id='AddItem'><i class='commonicon-save addbutton'></i>Save</a>" +
    "<a href='../Pages/Home.aspx' class='delete-icon close-icon pointer deletebutton' class='closebutton' title='Close' id='DelItem'><i class='commonicon-close deletebutton'></i>Close</a>" +
    "</div>" +
    "<h2 id='ComponentName'>Announcements</h2>" +
    "</div>" +
    "<div  class='form-section required'>" + 

    "<div class='input text'>" +
    "<label class='control-label'>Title</label>" +
    "<input class='form-control' type='text' value='' maxlength='30' id='txtTitle' disabled /></div>"+

    "<div class='textarea input'>"+
    "<label class='control-label'>Description</label>"+
    "<textarea id='txtrequiredDescription' class='form-control content'></textarea>"+
    "</div>"+
    "</div>";
    this.getAnnouncements();
    let Addevent = $('#AddItem');
    Addevent.on("click", (e: Event) => this.UpdateAnnouncement());
    $('.content').richText();
  }
  async getAnnouncements(){
    var listName = "Announcements";
    let columnArray = ["Announcements","ID","Title"];
    var Username = this.context.pageContext.user.displayName;

    var getItems = await readItems(listName, columnArray, 1, "Modified","Title","Announcements");
    if(getItems.length > 0)
    {
      $('#txtrequiredDescription').val(getItems[0].Announcements);
      $('.richText-editor').html(getItems[0].Announcements)
      $('#txtTitle').val(getItems[0].Title);
      ItemID = getItems[0].ID;
    }
  }

  async UpdateAnnouncement(){
      var listName = "Announcements";
      let itemObj = {
        Announcements: $('.richText-editor').html(),
      };

      updateItem(listName, ItemID, itemObj).then(result => {
        window.location.href = this.context.pageContext.web.absoluteUrl + "/pages/Home.aspx";
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
