import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './EdmsHomeAnnouncementWebPart.module.scss';
import * as strings from 'EdmsHomeAnnouncementWebPartStrings';

import {readItems,checkUserinGroup} from '../../commonJS';

declare var $;


export interface IEdmsHomeAnnouncementWebPartProps {
  description: string;
}

export default class EdmsHomeAnnouncementWebPart extends BaseClientSideWebPart<IEdmsHomeAnnouncementWebPartProps> {

  userflag: boolean = false;
  public render(): void {
    this.domElement.innerHTML = `
    <section class="cont-section about-section">
    <h3 id='HeadingAnnounce' class="tt-head"><a id='AnnounceEdit' href='../Pages/EditListItem.aspx?CName=Announcements'>Edit</a></h3>
    <div class="annoc-item">
    <p id='ParaAnnounce'></p>
    </div>
    </section>`;

    var _this = this;
    //Checking user details in group
    checkUserinGroup("Admin", this.context.pageContext.user.email, function (result) {
      if (result == 1) {
        _this.userflag = true;
      }
      _this.getAnnouncements(_this.userflag);
    })
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
