import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './EdmsProjectVideoDetailsWebPart.module.scss';
import * as strings from 'EdmsProjectVideoDetailsWebPartStrings';
import 'jquery';
require('bootstrap');
import { SPComponentLoader } from '@microsoft/sp-loader';
import { GetQueryStringParams,checkUserinGroup,readItems,deleteItem } from '../../commonJS';
declare var $;
declare var alertify;

export interface IEdmsProjectVideoDetailsWebPartProps {
  description: string;
}

export default class EdmsProjectVideoDetailsWebPart extends BaseClientSideWebPart<IEdmsProjectVideoDetailsWebPartProps> {

  public render(): void {
    SPComponentLoader.loadCss("https://cdnjs.cloudflare.com/ajax/libs/font-awesome/4.7.0/css/font-awesome.min.css");
    let siteURL = this.context.pageContext.web.absoluteUrl;
    this.domElement.innerHTML = "<div class='breadcrumb'>" +
    "<ol>" +
      "<li><a href='" + siteURL + "/Pages/Home.aspx' title='Home'>Home</a></li>" +
      "<li><a href='" + siteURL + "/Pages/VideoCollections.aspx'>Video Collections</a></li>" +
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
    
  }


  // TRIM SPACE IN QUERY STRING

  replaceAllSpaces(str) {   
    var arr = str.split('%20');
    var modifiedStr = arr.join(' ');
    return modifiedStr;
  }

  // TRIM PLUS IN QUERY STRING

  replaceAllPlus(str) {   
    var arr = str.split('+');
    var modifiedStr = arr.join(' ');
    return modifiedStr;
  }

  // DISPLAY IMAGE ITEMS

  /****** START ******/

  async getItems() {
    var q_imgeventid = GetQueryStringParams("imgeventid");
    var t_imgeventid = this.replaceAllSpaces(q_imgeventid);
    var t1_imgeventid = this.replaceAllPlus(t_imgeventid);
    var q_imgHtml = "";
    this.VidGalDetails(t1_imgeventid, q_imgHtml)
  }

  async VidGalDetails(imgeventid, ImgHtml) {
    let columnArray: any = ["ID", "Title", "FileRef", "FileLeafRef", "FileSystemObjectType", "FileDirRef", "LinkFilename","LinkURL"];
    var PageHeader = "";
    var VidHtml = "";
    var VidSrc = "";
    var VidItems = await readItems("Video Gallery", columnArray, 5000, "ID");
    var VidItemsLen = VidItems.length;
    var arr = [];
    var arrFile = [];
    var EventTitle = "";

      for (var i = 0; i < VidItemsLen; i++) {
        var eventname = VidItems[i].FileLeafRef;
        if (eventname != undefined) {
          if ($.inArray(eventname, arr) < 0) {
            arr.push(eventname);
          }
        }
      }

      for (var k = 0; k < VidItemsLen; k++) {
        if (VidItems[k].FileSystemObjectType == 1 && VidItems[k].LinkFilename == imgeventid) {
          PageHeader = VidItems[k].FileLeafRef;
        }
      }

      for (var i = 0; i < VidItemsLen; i++) {
        if (VidItems[i].FileSystemObjectType == 0) {
          let actFolderName = VidItems[i].FileDirRef;
          let urlFolderName = actFolderName.substr(actFolderName.lastIndexOf('/') + 1);
          arrFile.push(VidItems[i].FileSystemObjectType);
          if (urlFolderName == imgeventid && VidItems[i].FileSystemObjectType == 0) {
              arrFile.push(VidItems[i].FileSystemObjectType);
              VidSrc = VidItems[i].FileDirRef + "/" + VidItems[i].FileLeafRef;
              $(".page-title").text(PageHeader);
              EventTitle = VidItems[i].FileLeafRef;
              if(VidItems[i].LinkURL == null ){
                VidHtml += "<div class='col-lg-2 col-md-2 col-sm-4 col-xs-12'>" +
                          "<div class='gallery-list'>" +
                            "<a target='_blank' id=" + VidItems[i].ID + " href='" + this.context.pageContext.web.absoluteUrl + "/Video%20Gallery/Forms/AllItems.aspx?id=" + VidSrc + "&parent=" + VidItems[i].FileDirRef + "'>" + "<img style='height:100px;width:100px;' src='" + this.context.pageContext.site.absoluteUrl + "/_catalogs/masterpage/BloomHomepage/images/video-icon.jpg'>" +
                            "</a>" +
                            "<button style='margin-top: 17px;margin-right: -6px;color: white;background-color: grey;' class='deleteFolder'><i  class='fa fa-trash'></i></button>" +
                            "<h4>" + EventTitle + "<span></span></h4>" +
                          "</div>" +
                        "</div>";
              }
              else if(VidItems[i].LinkURL != null ){
                VidHtml += "<div class='col-lg-2 col-md-2 col-sm-4 col-xs-12'>" +
                          "<div class='gallery-list'>" +
                            "<a target='_blank' id=" + VidItems[i].ID + " href='"+ VidItems[i].LinkURL.Url + "'>" + "<img style='height:100px;width:100px;' src='" + this.context.pageContext.site.absoluteUrl + "/_catalogs/masterpage/BloomHomepage/images/video-icon.jpg'>" +
                            "</a>" +
                            "<button style='margin-top: 17px;margin-right: -6px;color: white;background-color: grey;' class='deleteFolder'><i  class='fa fa-trash'></i></button>" +
                            "<h4>" + EventTitle + "<span></span></h4>" +
                          "</div>" +
                        "</div>";
              }
            }
        }
        
      }
      if(jQuery.inArray(0, arrFile) == -1)
      {
        VidHtml += '<div class="col-lg-12 col-md-12 col-sm-12 col-xs-12">'+
                        '<h3>No Items To Display</h3>'+
                      '</div>';
      }
      $(".gallery-listsec").append(VidHtml);
      $('#deleteButtonField').hide();
    checkUserinGroup("Video Gallery",this.context.pageContext.user.loginName,function(result){
      if(result == 1){$("#deleteButtonField").show();$('#AddingButtons').show();}else if(result == 0){$("#deleteButtonField").hide();$('#AddingButtons').hide();}
    });

    $('.deleteFolder').click(function(){
        let itemId = $(this).parent().find('a').attr('id');
        alertify.confirm("Are you sure you want to delete selected Image ?", function (e) {
          if (e) {
            alertify.set('notifier', 'position', 'top-right');
            alertify.success("Video Deleted Successfully");
            deleteItem("Video Gallery",itemId);
            location.reload();
         } else {}
      },function (e){if(e){}else{}}).set('closable', false).setHeader('Confirmation') ;  
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
