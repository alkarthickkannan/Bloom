import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';
import * as strings from 'EdmsMediaGalleryHomeWebPartStrings';
import { readItems, checkUserinGroup } from '../../commonJS';
import 'jquery';
// require('bootstrap');
declare var $;

export interface IEdmsMediaGalleryHomeWebPartProps {
  description: string;
}

export default class EdmsMediaGalleryHomeWebPart extends BaseClientSideWebPart<IEdmsMediaGalleryHomeWebPartProps> {
  userflag: boolean = false;
  public render(): void {
    var _this = this;
    //Checking user details in group
    checkUserinGroup("Media Gallery", this.context.pageContext.user.email, function (result) {
      //console.log(result);
      if (result == 1) {
        _this.userflag = true;
      }
      _this.MediaGallery();
    })
  }

  public MediaGallery() {
    var siteURL = this.context.pageContext.web.absoluteUrl;
    this.domElement.innerHTML = `
    <section class="banner-section ban-sec1">
      <div class="ban-section">
        <h3 class="tt-head">Media Gallery <a id="addEvents" class="pull-right" href="../Pages/ListView.aspx?CName=Media Gallery"> More </a></h3>
        <div id="carousel-banner" class="carousel carousel-fade" data-ride="carousel">
          <!-- Wrapper for slides -->
          <div id="carouselDataBind" class="carousel-inner" role="listbox">
          </div>
        </div>
      </div>
    </section>
    `;
    this.GetmediaGalleryItems(this.userflag);
  }

  async GetmediaGalleryItems(userflag) {
    let sliderHtml ="";
    let renderHtml = "";
    let renderliitems = "";
    let objResults = await readItems("Media Gallery", ["LinkURL", "Display", "MediaFileType", "Image", "Title"], 3, "Modified", "Display", 1);
    let objResultsLen = objResults.length;

    // VALIDATE IF GALLERY EMPTY
    if (objResultsLen == 0 ) {
      sliderHtml = 
        // <!-- Indicators -->
        "<ol class='carousel-indicators'>"+
          "<li data-target='#carousel-banner' data-slide-to='0' class='active'></li>"+
        "</ol>";
      renderHtml += 
        "<div class='item active'>" +
          "<img src='" + this.context.pageContext.site.absoluteUrl + "/Site Assets/ImageGallery/no_image_available.jpeg' alt='Slide' title='No Item to Display'/>" +
        "</div>"

    }
    // IF DISPLAYED ITEMS = 1

    else if (objResultsLen == 1 ) {

      sliderHtml = 
        // <!-- Indicators -->
        "<ol class='carousel-indicators'>"+
          "<li data-target='#carousel-banner' data-slide-to='0' class='active'></li>"+
        "</ol>";

      if (objResults[0].MediaFileType == "Image") {
        renderHtml += "<div class='item active'> <img src='" + objResults[0].Image.Url + "' alt='Slide' title='" + objResults[0].Title + "' /> </div>"
      } else if (objResults[0].MediaFileType == "Video") {
        renderHtml += "<div class='item active'>" +
          "<video width='100%' height='100%' controls poster='" + objResults[0].Image.Url + "_jpg.jpg'>" +
          "<source src='" + objResults[0].Image.Url + "' type='video/mp4'>Your browser does not support HTML5 video." +
          "</video>" +
          "</div>"
      } else if (objResults[0].MediaFileType == "Streams") {
        renderHtml += "<div class='item active'>" +
          "<a href='"+ objResults[0].LinkURL.Url +"' target='_blank'><img href='"+ objResults[0].LinkURL.Url +"' src='" + objResults[0].Image.Url + "' alt='Slide' title='" + objResults[0].Title + "' /></a>" +
          "</div>"
      }
    }

    // IF DISPLAYED ITEMS = 2

    else if (objResultsLen == 2 ){

        sliderHtml = 
        // <!-- Indicators -->
        "<ol class='carousel-indicators'>"+
          "<li data-target='#carousel-banner' data-slide-to='0' class='active'></li>"+
          "<li data-target='#carousel-banner' data-slide-to='1' ></li>"+
        "</ol>";

        // FIRST ITEM TO BE BINDED

        if (objResults[0].MediaFileType == "Image") {
          renderHtml += "<div class='item active'> <img src='" + objResults[0].Image.Url + "' alt='Slide' title='" + objResults[0].Title + "' /> </div>"
        } else if (objResults[0].MediaFileType == "Video") {
          renderHtml += "<div class='item active'>" +
            "<video width='100%' height='100%' controls poster='" + objResults[0].Image.Url + "_jpg.jpg'>" +
            "<source src='" + objResults[0].Image.Url + "' type='video/mp4'>Your browser does not support HTML5 video." +
            "</video>" +
            "</div>"
        } else if (objResults[0].MediaFileType == "Streams") {
          renderHtml += "<div class='item active'>" +
            "<a href='" + objResults[0].LinkURL.Url + "' target='_blank'><img href='" + objResults[0].LinkURL.Url + "' src='" + objResults[0].Image.Url + "' alt='Slide' title='" + objResults[0].Title + "' /></a>" +
            "</div>"
        }

        // SECOND ITEM TO BE BINDED

        if (objResults[1].MediaFileType == "Image") {
          renderHtml += "<div class='item'> <img src='" + objResults[1].Image.Url + "' alt='Slide' title='" + objResults[1].Title + "' /> </div>"
        } else if (objResults[1].MediaFileType == "Video") {
          renderHtml += "<div class='item'>" +
            "<video width='100%' height='100%' controls poster='" + objResults[1].Image.Url + "_jpg.jpg'>" +
            "<source src='" + objResults[1].Image.Url + "' type='video/mp4'>Your browser does not support HTML5 video." +
            "</video>" +
            "</div>"
        } else if (objResults[1].MediaFileType == "Streams") {
          renderHtml += "<div class='item'>" +
            "<a href='" + objResults[1].LinkURL.Url + "' target='_blank'><img href='" + objResults[1].LinkURL.Url + "' src='" + objResults[1].Image.Url + "' alt='Slide' title='" + objResults[1].Title + "' /></a>" +
            "</div>"
        }
       
      }
      // IF DISPLAYED ITEMS = 3

      else if (objResultsLen >= 3 ){
        
        sliderHtml = 
        // <!-- Indicators -->
        "<ol class='carousel-indicators'>"+
          "<li data-target='#carousel-banner' data-slide-to='0' class='active'></li>"+
          "<li data-target='#carousel-banner' data-slide-to='1' ></li>"+
          "<li data-target='#carousel-banner' data-slide-to='2' ></li>"+
        "</ol>";

        // FIRST ITEM TO BE BINDED

        if (objResults[0].MediaFileType == "Image") {
          renderHtml += "<div class='item active'> <img src='" + objResults[0].Image.Url + "' alt='Slide' title='" + objResults[0].Title + "' /> </div>"
        } else if (objResults[0].MediaFileType == "Video") {
          renderHtml += "<div class='item active'>" +
            "<video width='100%' height='100%' controls poster='" + objResults[0].Image.Url + "_jpg.jpg'>" +
            "<source src='" + objResults[0].Image.Url + "' type='video/mp4'>Your browser does not support HTML5 video." +
            "</video>" +
            "</div>"
        } else if (objResults[0].MediaFileType == "Streams") {
          renderHtml += "<div class='item active'>" +
            "<a href='" + objResults[0].LinkURL.Url + "' target='_blank'><img href='" + objResults[0].LinkURL.Url + "' src='" + objResults[0].Image.Url + "' alt='Slide' title='" + objResults[0].Title + "' /></a>" +
            "</div>"
        }

        // OTHER ITEMS

        for (var i = 1; i < objResultsLen; i++) {
          if (objResults[i].MediaFileType == "Image") {
            renderHtml += "<div class='item'> <img src='" + objResults[i].Image.Url + "' alt='Slide' title='" + objResults[i].Title + "' /> </div>"
          } else if (objResults[i].MediaFileType == "Video") {
            renderHtml += "<div class='item'>" +
              "<video width='100%' height='100%' controls poster='" + objResults[i].Image.Url + "_jpg.jpg'>" +
              "<source src='" + objResults[i].Image.Url + "' type='video/mp4'>Your browser does not support HTML5 video." +
              "</video>" +
              "</div>"
          } else if (objResults[i].MediaFileType == "Streams") {
            renderHtml += "<div class='item'>" +
              "<a href='" + objResults[i].LinkURL.Url + "' target='_blank'><img href='" + objResults[i].LinkURL.Url + "' src='" + objResults[i].Image.Url + "' alt='Slide' title='" + objResults[i].Title + "' /></a>" +
              "</div>"
          }
        }
      }

      $('#carouselDataBind').before(sliderHtml);
      $('#carouselDataBind').append(renderHtml);

    // VALIDATE USER 

    if (userflag == false) {
      $('#addEvents').hide();
    }
    else {
      $('#addEvents').show();
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
