import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';
import * as strings from 'EdmsProjectsQuickLinksWebPartStrings';
import pnp from 'sp-pnp-js';
import "jquery";
import { checkUserinGroup } from "../../commonJS";
declare var $;

export interface IEdmsProjectsQuickLinksWebPartProps {
  description: string;
}

export default class EdmsProjectsQuickLinksWebPart extends BaseClientSideWebPart<IEdmsProjectsQuickLinksWebPartProps> {
  userflag: boolean = false;
  public render(): void {
    var _this = this;
    checkUserinGroup("Quick Links", this.context.pageContext.user.email, function (result) {
      if (result == 1) {
        _this.userflag = true;
      }
      _this.QuickLinksDisplay();
    })
  }
  QuickLinksDisplay() {
    var siteURL = this.context.pageContext.web.absoluteUrl;
    this.domElement.innerHTML = `
    <section class="vertical-menu">
    <div class="panel-group panel-link" id="accordionMenu" role="tablist" aria-multiselectable="true">
    <div class="panel panel-default">
    <div class="panel-heading" role="tab" id="headingOne">
    <a href="../Pages/Listview.aspx?CName=ProjectQuickLinks&LinkType=Projects" class="panel-title" id="AProjects"> <i class="icon-home"></i>Projects </a>
    <a role="button" data-toggle="collapse" data-parent="#accordionMenu" href="#collapseOne" aria-expanded="false" aria-controls="collapseOne"></a>
    </div>
    <div id="collapseOne" class="panel-collapse collapse" role="tabpanel" aria-labelledby="headingOne">
    </div>
    </div>
    <div class="panel panel-default">
    <div class="panel-heading" role="tab" id="headingTwo">
    <a href="../Pages/Listview.aspx?CName=ProjectQuickLinks&LinkType=Documents" class="panel-title" id="ADocuments"><i class="icon-file"></i>Documents </a>
     <a class="collapsed" role="button" data-toggle="collapse" data-parent="#accordionMenu" href="#collapseTwo" aria-expanded="false" aria-controls="collapseTwo">  </a>
    </div>
    <div id="collapseTwo" class="panel-collapse collapse" role="tabpanel" aria-labelledby="headingTwo">
    </div>
    </div>
    
    <div class="panel panel-default">
    <div class="panel-heading" role="tab" id="headingThree">
    <a href="#" class="panel-title"> <i class="icon-gallery"></i>Photos </a>
    <a class="collapsed" role="button" data-toggle="collapse" data-parent="#accordionMenu" href="#collapseThree" aria-expanded="false" aria-controls="collapseThree"> </a>
    </div>
    <div id="collapseThree" class="panel-collapse collapse" role="tabpanel" aria-labelledby="headingThree">
    <div class="panel-body">
    <ul class="nav">
    <li><a href="../Pages/ImageGallery.aspx">Image Gallery</a></li>
    <li><a href="../Pages/VideoGallery.aspx">Video Gallery</a></li>
    </ul>
    </div>
    </div>
    </div>
    </div>
    </section>`;
    this.displayQuickLinks(this.userflag);
    // var divHeight = $('#right-side').height(); 
    // $('.vertical-menu').css('min-height', divHeight+'px');
    if(this.userflag == false){
      $("#AProjects,#ADocuments").removeAttr("href");
    }

  }

  async displayQuickLinks(userflag) {
    // GET SUBSITES
    let webUrl = this.context.pageContext.site.absoluteUrl + "/EDMS/Projects/"
    let subRenderhtml = "";
    // let subsiteList = await getListOfSubSites(webUrl);

    // let subsiteList = await readItems("ProjectQuickLinks", ["Title","ID", "LinkURL", "LinkType", "Display"], 50, "Modified", "LinkType", "Projects");

    let subsiteList = await pnp.sp.web.lists.getByTitle("ProjectQuickLinks").items.select("Title,ID,LinkURL,LinkType,Display").filter("Display eq 1 and LinkType eq 'Projects'").top(50).orderBy("Modified").get();

    let subsiteListLen = subsiteList.length;
    if (subsiteListLen == 0) {
      subRenderhtml += "<div class='panel-body'>" +
        "<ul class='nav'>" +
        "<li><a href='#'>No Sites found</a></li>" +
        "</ul>" +
        "</div>";
        $('#collapseOne').append(subRenderhtml);
    } else if (subsiteListLen > 0) {
      subRenderhtml += "<div class='panel-body'>";
      subRenderhtml += "<ul class='nav'>";
      for (let i = 0; i < subsiteListLen; i++) {
        subRenderhtml += "<li><a href='" + subsiteList[i].LinkURL.Url + "'>" + subsiteList[i].Title + "</a></li>"
      }
      subRenderhtml += "</ul>";
      subRenderhtml += "</div>";
      $('#collapseOne').append(subRenderhtml);
    }
    // GET DOC LIB UNDER THE SUBSITES

    var siteURL = this.context.pageContext.web.absoluteUrl;
    let docRenderHtml = "";
    // let docLibList = await getListOfDocLib(15, "LastItemModifiedDate");
    let docLibList = await pnp.sp.web.lists.getByTitle("ProjectQuickLinks").items.select("Title,ID,LinkURL,LinkType,Display").filter("Display eq 1 and LinkType eq 'Documents'").top(50).orderBy("Modified").get();
    let docLibListLen = docLibList.length;
    if (docLibListLen == 0) {
      docRenderHtml += "<div class='panel-body'>" +
        "<ul class='nav'>" +
        "<li><a href='#'>No libraries to show</a></li>" +
        "</ul>" +
        "</div>";

        $('#collapseTwo').append(docRenderHtml);
    } else if (docLibListLen > 0) {
      docRenderHtml += "<div class='panel-body'>";
      docRenderHtml += "<ul class='nav'>";
      for (let i = 0; i < docLibListLen; i++) {    
        // Change the next line to a proper logic
          // let transformedURL = docLibList[i].EntityTypeName.replace("_x0020_"," ").replace("_x0020_"," ");
          // docRenderHtml += "<li><a href='" + siteURL + "/" + transformedURL + "'>" + docLibList[i].Title + "</a></li>"
          docRenderHtml += "<li><a href='" + docLibList[i].LinkURL.Url + "'>" + docLibList[i].Title + "</a></li>"
        }
      docRenderHtml += "</ul>";
      docRenderHtml += "</div>";
      $('#collapseTwo').append(docRenderHtml);
    }
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
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
                PropertyPaneTextField("description", {
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