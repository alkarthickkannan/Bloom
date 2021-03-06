import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';
import 'jquery';
import styles from './EdmsDeptAnnouncementWebPart.module.scss';
import * as strings from 'EdmsDeptAnnouncementWebPartStrings';
import {readItems,checkUserinGroup} from '../../commonJS';

declare var $;

export interface IEdmsDeptAnnouncementWebPartProps {
  description: string;
}

export default class EdmsDeptAnnouncementWebPart extends BaseClientSideWebPart<IEdmsDeptAnnouncementWebPartProps> {

  userflag: boolean = false;
  public render(): void {
    this.domElement.innerHTML = '<section class="about-section announ-section">'+
      "<h3 id='HeadingAnnounce'><a id='AnnounceEdit' href='../Pages/EditListItem.aspx?CName=Announcements'>Edit</a></h3>"+
      "<p id='ParaAnnounce'></p>"+
    "</section>";

    var _this = this;
    //Checking user details in group
    checkUserinGroup("Admin", this.context.pageContext.user.email, function (result) {
      if (result == 1) {
        _this.userflag = true;
      }
      _this.getAnnouncements(_this.userflag);
    })

    // $("#Showmore").click(function(){

    //   if($("#ParaAnnounce").hasClass("ParaAnnounce")) {
    //       $(this).text("Show Less");
          
    //   } else {
    //       $(this).text("Show More");
    //   }
      
    //   $("#ParaAnnounce").toggleClass("ParaAnnounce");
    //   var divHeight = $('#right-side').height();
    //   $('.vertical-menu').css('min-height', divHeight + 'px');

    //   });
  }

  async getAnnouncements(userflag){
    var listName = "Announcements";
    let columnArray = ["Announcements","ID","Title"];
    var Username = this.context.pageContext.user.displayName;

    var getItems = await readItems(listName, columnArray, 1, "Modified","ID",1);
    if(getItems.length > 0)
    {
      $('#ParaAnnounce').html(getItems[0].Announcements);
      $('#HeadingAnnounce').prepend(getItems[0].Title);
      if(userflag == true)
      {
        $('#AnnounceEdit').show();
      }
      else{
        $('#AnnounceEdit').hide();
      }
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
