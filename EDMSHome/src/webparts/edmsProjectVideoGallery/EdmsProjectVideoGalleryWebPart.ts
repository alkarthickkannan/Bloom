import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './EdmsProjectVideoGalleryWebPart.module.scss';
import * as strings from 'EdmsProjectVideoGalleryWebPartStrings';
import 'jquery';
require('bootstrap');
import { SPComponentLoader } from '@microsoft/sp-loader';
import { checkUserinGroup,readItems,DeleteFolder } from '../../commonJS';
declare var $;
declare var alertify;

export interface IEdmsProjectVideoGalleryWebPartProps {
  description: string;
}

export default class EdmsProjectVideoGalleryWebPart extends BaseClientSideWebPart<IEdmsProjectVideoGalleryWebPartProps> {

  public render(): void {
    SPComponentLoader.loadCss("https://cdnjs.cloudflare.com/ajax/libs/font-awesome/4.7.0/css/font-awesome.min.css");
    var siteURL = this.context.pageContext.legacyPageContext.webAbsoluteUrl;
    this.domElement.innerHTML =
    "<div class='breadcrumb'>" +
    "<ol>" +
      "<li><a href='" + siteURL + "/Pages/Home.aspx' title='Home'>Home</a></li>" +
      "<li><a href='" + siteURL + "/Pages/VideoGallery.aspx'>Video Gallery</a></li>" +
    "</ol>" +
  "</div>" +
  "<div class='title-section'>" +
    "<div class='button-field'>" +
      "<a href='" + siteURL + "/Pages/AddListItem.aspx?CName=Video Gallery' title='Add New' class='pointer' id='AddingButtons'><i class='icon-add'></i>Add New</a>" +
      "<a href='" + siteURL + "/Pages/Home.aspx' class='delete-icon pointer' title='Close'><img src='" + this.context.pageContext.site.absoluteUrl + "/_catalogs/masterpage/BloomHomepage/images/close-icon.png'>Close</a>" +
    "</div>" +
    "<h2>Video Gallery</h2>" +
  "</div>" +
    
    `
    <div class="gallery-listsec">
    </div>
    `;
    this.getItems();
    this.checkUserPermissionForDeletion();
  }


  async getItems() {
    var ImgHtml = "";
    var ImgSrc = "";
    var EventTitle = "";
    let columnArray: any = ["ID", "FileLeafRef", "FileSystemObjectType", "FileDirRef"];
    let picItems = await readItems("Video Gallery", columnArray, 5000, "ID");
    
    var itemLength = picItems.length;
    var arr = [];
    var Flag2 = 0;

  if(picItems.length > 0)
  {
    for (var i = 0; i < itemLength; i++) {
      var eventname = picItems[i].FileLeafRef;
      if (eventname != undefined) {
        if ($.inArray(eventname, arr) < 0) {
          arr.push(eventname);
        }
      }
    }
    var arrFlag = 0;
    for (var j = 0; j < arr.length; j++) {
      for (var k = 0; k < itemLength; k++) {
        if (arr[j] == picItems[k].FileLeafRef) {
          if (arrFlag == 0) {
            ImgSrc = picItems[k].FileDirRef + "/" + picItems[k].FileLeafRef;
            EventTitle = picItems[j].FileLeafRef;
            if(picItems[k].FileSystemObjectType == 1){
              ImgHtml += "<div class='col-lg-2 col-md-2 col-sm-4 col-xs-12'>" +
                "<div class='gallery-list'>" +
                "<a href='" + this.context.pageContext.web.absoluteUrl + "/Pages/VideoGalleryItems.aspx?imgeventid=" + arr[j] + "'title=''><img style='height:100px;width:100px;' src='" + this.context.pageContext.site.absoluteUrl + "/_catalogs/masterpage/BloomHomepage/images/folder-images-icon.png'>" +
                "</a>" +
                "<button style='margin-top: 17px;margin-right: -6px;color: white;background-color: grey;' class='deleteFolder'><i class='fa fa-trash'></i></button>" +
                "<h4>" + EventTitle + "</h4>" +
                "</div>" +
                "</div>";
                arrFlag++;
              }
            }
          }
        }
        
      arrFlag = 0;
    }
    $(".gallery-listsec").append(ImgHtml);
    $('.deleteFolder').hide();
  }
  else
  {
    ImgHtml += '<div class="col-lg-12 col-md-12 col-sm-12 col-xs-12">'+
                  '<h3>No Items To Display</h3>'+
                '</div>';
    $(".gallery-listsec").append(ImgHtml);

  }
    checkUserinGroup("Video Gallery",this.context.pageContext.user.loginName,function(result){
      if(result == 1){$(".deleteFolder").show();$('#AddingButtons').show();}else if(result == 0){$(".deleteFolder").hide();$('#AddingButtons').hide();}
    });

    // DELETE FOLDER - START

    $('.deleteFolder').click(function (event){
      event.preventDefault();
      let folderName = $(this).next().text();
      alertify.confirm("Are you sure you want to delete selected Folder ?", function (e) {
        if (e) {
          alertify.set('notifier', 'position', 'top-right');
          alertify.success("Folder Deleted Successfully");
           DeleteFolder("Video Gallery",folderName);
           location.reload();
        } else {}
      },function (e){if(e){}else{}}).set('closable', false).setHeader('Confirmation') ;      
    });
  }

  // USER PERMISSION CHECK - START

  checkUserPermissionForDeletion(){
    let email = this.context.pageContext.user.loginName;
    let compName = "Video Gallery";
    checkUserinGroup(compName,email,function(result){
        if (result == 1){
          $('.deleteFolder').show();
        }else{
          $('.deleteFolder').hide();
        }    
      });
    }

    // USER PERMISSION CHECK - END


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
