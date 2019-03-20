import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';
import { SPComponentLoader } from '@microsoft/sp-loader';
import styles from './ImageGalleryDetailWebPart.module.scss';
import * as strings from 'ImageGalleryDetailWebPartStrings';
import { checkUserinGroup, readItems, GetQueryStringParams, deleteItem } from '../../commonJS';
import pnp from "sp-pnp-js";
import 'jquery';
require('bootstrap');
declare var $;
declare var alertify;
export interface IImageGalleryDetailWebPartProps {
  description: string;
}

export default class ImageGalleryDetailWebPart extends BaseClientSideWebPart<IImageGalleryDetailWebPartProps> {

  public render(): void {

    var siteURL = this.context.pageContext.site.absoluteUrl;
    SPComponentLoader.loadCss(siteURL + "/_catalogs/masterpage/BloomHomepage/css/style.css");
    SPComponentLoader.loadScript(siteURL + "/_catalogs/masterpage/BloomHomepage/js/jquery.min.js");
    SPComponentLoader.loadScript(siteURL + "/_catalogs/masterpage/BloomHomepage/js/jssor.slider.min.js");
    this.domElement.innerHTML =

      "<div class='breadcrumb'>" +
      "<ol>" +
      "<li><a href='" + siteURL + "/Pages/Home.aspx' title='Home'>Home</a></li>" +
      "<li><a href='" + siteURL + "/Pages/ImageGallery.aspx'>Image Gallery</a></li>" +
      "</ol>" +
      "</div>" +

      "<div class='title-section'>" +
      "<div class='button-field'>" +
      "<a href='ImageGallery.aspx' class='pointer' title='Close' style='background:#53545E;'><img src='" + siteURL + "/_catalogs/masterpage/BloomHomepage/images/close-icon.png'>Close</a>" +
      "</div>" +
      "<h2>Image Gallery</h2>" +
      "</div>" +

      `
    <br>

    <h3 class='page-title pageheader' style='margin-left:75px;margin:botton:30px;'></h3> 
    <div id="jssor_1" style="position:relative;margin:0 auto;top:0px;left:0px;width:980px;height:480px;overflow:hidden;visibility:hidden;"> 
    
    <!-- Loading Screen -->
    
    <div data-u='loading' class='jssorl-009-spin' style='position:absolute;top:0px;left:0px;width:100%;height:100%;text-align:center;background-color:rgba(0,0,0,0.7);'>
        <img style='margin-top:-19px;position:relative;top:50%;width:38px;height:38px;' src='/sites/spuat/_catalogs/masterpage/BloomHomepage/images/slider-loader.svg'/>
    </div>
    
    <div class="image-slides-cont-new" data-u="slides" style="cursor:default;position:relative;top:0px;left:0px;width:980px;height:380px;overflow:hidden;">
        
    </div>

    <!-- Thumbnail Navigator -->

    <div data-u="thumbnavigator" class="jssort101" style="position:absolute;left:0px;bottom:0px;width:980px;height:100px;background-color:#000;" data-autocenter="1" data-scale-bottom="0.75">
        <div data-u="slides">
            <div data-u="prototype" class="p" style="width:190px;height:90px;">
                <div data-u="thumbnailtemplate" class="t">
                </div>
                <svg viewbox="0 0 16000 16000" class="cv">
                    <circle class="a" cx="8000" cy="8000" r="3238.1"></circle>
                    <line class="a" x1="6190.5" y1="8000" x2="9809.5" y2="8000"></line>
                    <line class="a" x1="8000" y1="9809.5" x2="8000" y2="6190.5"></line>
                </svg>
            </div>
        </div>
    </div>

    <!-- Arrow Navigator -->

    <div data-u="arrowleft" class="jssora106" style="width:55px;height:55px;top:162px;left:30px;" data-scale="0.75">
        <svg viewbox="0 0 16000 16000" style="position:absolute;top:0;left:0;width:100%;height:100%;">
            <circle class="c" cx="8000" cy="8000" r="6260.9">
            
            </circle>
            <polyline class="a" points="7930.4,5495.7 5426.1,8000 7930.4,10504.3 "></polyline>
            <line class="a" x1="10573.9" y1="8000" x2="5426.1" y2="8000"></line>
        </svg>
    </div>

    <div data-u="arrowright" class="jssora106" style="width:55px;height:55px;top:162px;right:30px;" data-scale="0.75">
        <svg viewbox="0 0 16000 16000" style="position:absolute;top:0;left:0;width:100%;height:100%;">
            <circle class="c" cx="8000" cy="8000" r="6260.9"></circle>
            <polyline class="a" points="8069.6,5495.7 10573.9,8000 8069.6,10504.3 "></polyline>
            <line class="a" x1="5426.1" y1="8000" x2="10573.9" y2="8000"></line>
        </svg>
    </div>
  </div>
    `;
    this.checkUserPermissionForDeletion();
    this.getItems();
  }

  checkUserPermissionForDeletion() {
    let email = this.context.pageContext.user.loginName;
    let compName = "Image Gallery";
    checkUserinGroup(compName, email, function (result) {
      if (result == 1) {
        $('.deleteFolder').show();
        $('#AddingButtons').show();
        $('.delete-icon').show();
      } else {
        $('.deleteFolder').hide();
        $('#AddingButtons').hide();
        $('.delete-icon').hide();
      }
    });
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
    this.ImgGalDetails(t1_imgeventid, q_imgHtml)
  }

  async ImgGalDetails(imgeventid, ImgHtml) {
    let columnArray: any = ["ID", "Title", "FileRef", "FileLeafRef", "FileSystemObjectType", "FileDirRef", "LinkFilename"];
    var PageHeader = "";
    var picItems = await readItems("Image Gallery", columnArray, 5000, "ID");
    var picItemsLen = picItems.length;

    for (var k = 0; k < picItemsLen; k++) {
      if (picItems[k].FileSystemObjectType == 1 && picItems[k].LinkFilename == imgeventid) {
        PageHeader = picItems[k].FileLeafRef;
      }
    }
    for (var i = 0; i < picItemsLen; i++) {
      let actFolderName = picItems[i].FileRef;
      let urlFolderName = actFolderName.substr(actFolderName.lastIndexOf('/') + 1);
      if (picItems[i].FileSystemObjectType == 1 && urlFolderName == imgeventid) {
        let folderServerURL = picItems[i].FileRef;
        this.ImageGalleryFolderchecking(folderServerURL)
      }
      if (picItems[i].FileSystemObjectType == 0) {
        let actImageFolderName = picItems[i].FileDirRef;
        let urlImageFolderName = actImageFolderName.substr(actImageFolderName.lastIndexOf('/') + 1);
        if (urlImageFolderName == imgeventid) {
          $(".page-title").text(PageHeader);
          ImgHtml += "<div>" +
            "<img data-u='image' src='" + picItems[i].FileRef + "'><div id='deleteButtonField' class='button-field'><a class='delete-icon' title='Delete' id='" + picItems[i].ID + "'><img src='" + this.context.pageContext.site.absoluteUrl + "/_catalogs/masterpage/BloomHomepage/images/close-icon.png'>Delete</a></div></img>" +
            "<img data-u='thumb' src='" + picItems[i].FileRef + "' />" +
            "</div>";
        }
        this.checkUserPermissionForDeletion();
      }
    }

    $('.image-slides-cont-new').append(ImgHtml);
    $('#deleteButtonField').hide();
    checkUserinGroup("Image Gallery", this.context.pageContext.user.loginName, function (result) {
      if (result == 1) { $("#deleteButtonField").show(); } else if (result == 0) { $("#deleteButtonField").hide(); }
    });

    $('.delete-icon').click(function () {
      let itemId = $(this).parent().find('a').attr('id');
      alertify.confirm("Are you sure you want to delete selected Image ?", function (e) {
        if (e) {
          alertify.success("");
          deleteItem("Image Gallery", itemId);
          location.reload();
        } else { }
      }, function (e) { if (e) { alertify.error(""); } else { } }).set('closable', false).setHeader('Confirmation');
    });
    SPComponentLoader.loadScript(this.context.pageContext.site.absoluteUrl + "/_catalogs/masterpage/BloomHomepage/js/jssorScript.js");
  }

  // FOR NO ITEM DISPLAY VALIDATION

  ImageGalleryFolderchecking(folderName: string) {
    let siteUrl = this.context.pageContext.web.absoluteUrl;
    console.log(siteUrl);
    $.ajax
      ({
        url: siteUrl +"/_api/web/getfolderbyserverrelativeurl('" + folderName + "')/files?",
        type: 'GET',
        headers:
        {
          "Accept": "application/json;odata=verbose",
          "Content-Type": "application/json;odata=verbose",

        },
        cache: false,
        success: function (data) {

          if (data.d.results.length == 0) {

            $(".page-title").text("No Item to Display");
            $('#jssor_1,.pointer').hide();
          }
        },
        error: function (data) {
          console.log(data.responseJSON.error);
        }
      });

  }

  /****** END *****/

  // DELETE IMAGE 

  DeleteItem(itemId) { deleteItem("Image Gallery", itemId); }
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
