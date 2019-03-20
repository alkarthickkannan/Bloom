import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

//import styles from './ViewListItemWebPart.module.scss';
import * as strings from 'ViewListItemWebPartStrings';
import pnp from "sp-pnp-js";
import 'jquery';
require('../../ExternalRef/js/jquery.richtext.js');
import '../../ExternalRef/css/richtext.min.css';
import {
  GetQueryStringParams,
  formatDate,
  addItems,
  formString,
  updateItem,
  readItems,
  readItem,
  checkUserinGroup,
  deleteItem
} from '../../commonJS';
declare var $;
declare var alertify: any;
export interface IViewListItemWebPartProps {
  description: string;
}
var userflag = false;
export default class ViewListItemWebPart extends BaseClientSideWebPart<IViewListItemWebPartProps> {

  public render(): void {
    var listName = GetQueryStringParams("CName").replace("%20", " ");

    this.domElement.innerHTML = `
    <div class="breadcrumb">
        <ol id="ListBreadcrumbs">
            <li><a href="../Pages/Home.aspx" class='pointer' title="Home">Home</a></li>
            <li><a id="ListViewBC" class='pointer'  title="Home"></a></li>
            <li><span id="ViewListItemBC"></span></li>
        </ol>
        
    </div>
            <div class="title-section">
                <h2 id="DocumentTitle"></h2>
            </div>
            <div class="form-section"> 
                <div class="logo-cropsec">
                    <div class="row">
                      <div class="col-lg-12 col-md-12 col-sm-12 col-xs-12" id='Img-Part'>
                        <div class="form-imgsec">
                        </div>
                      </div>
                      <div class="col-lg-12 col-md-12 col-sm-12 col-xs-12"  id='Form-Part'>
                      </div>
                      <div class="col-lg-12 col-md-12 col-sm-12 col-xs-12" id="Announcement-Sec" style="display:none;">
                                <div class="card">
                                  <div id="Viewer-Tab" class="tab-content">
                                  </div>
                                </div>
                        </div>
                    </div>
                </div>
            </div> 
            <div class='modal'><!-- Place at bottom of page --></div>`;

    this.renderhtml();
    this.FetchListItems();
  }

  async FetchListItems() {
    var listName = GetQueryStringParams("CName").replace("%20", " ");
    var ItemID = GetQueryStringParams("CID");
    let columnArray: any = this.GetColumns(listName);
    let GetListItems = await readItems(listName, columnArray, 1, "Modified", "ID", ItemID);
    var _this = this;
    if (GetListItems.length > 0) {
      if (listName == "Banners") {
        $('#View-img').attr('src', GetListItems[0].Image.Url);
        $('#txtTitle').val(GetListItems[0].Title);
        //BannerContent: $('.richText-editor').html(),   
        $('.richText-editor').html(GetListItems[0].BannerContent);
        if (GetListItems[0].LinkURL != null) {
          $('#txtHyperlink').val(GetListItems[0].LinkURL.Url);
        }
        else {
          $('#DivHyperLink').hide();
        }
        $('#DocumentTitle').text(listName);
      } else if (listName == "Holiday") {
        $('#txtTitle').val(GetListItems[0].Title);
        $('#txtStartDate').val(new Date(GetListItems[0].EventDate).toLocaleDateString());
        if ((GetListItems[0].EndEventDate) != null) {
          GetListItems[0].EndEventDate = new Date(GetListItems[0].EndEventDate).toLocaleDateString();
        }
        else {
          GetListItems[0].EndEventDate = "";
        }
        $('#txtEndDate').val(GetListItems[0].EndEventDate);
        $('#DocumentTitle').text(listName);
      } else if (listName == "News") {
        $('#View-img').attr('src', GetListItems[0].Image.Url);
        $('#txtTitle').val(GetListItems[0].Title);
        $('#txtDescription').val(GetListItems[0].Explanation);
        $('#txtDate').val(new Date(GetListItems[0].Date).toLocaleDateString());
        $('#DocumentTitle').text(listName);
      } else if (listName == "Quick Links") {
        $('#txtTitle').val(GetListItems[0].Title);
        $('#txtHyperlink').val(GetListItems[0].LinkURL.Url);
        $('#DocumentTitle').text("Quick Links");
      } else if (listName == "Employee Corner") {
        var Doctype = GetListItems[0].DocumentFile.Url.split(".");
        Doctype = Doctype[Doctype.length - 1];
        $('#cropped-img').attr('src', this.GetDocImages(Doctype));
        $('a.Link').attr('href', GetListItems[0].DocumentFile.Url);
        $('#txtTitle').val(GetListItems[0].Title);
        $('#txtDate').val(new Date(GetListItems[0].Modified).toLocaleDateString());
        $('#DocumentTitle').text("Employee Corner");
      } else if (listName == "Organizational Policies") {
        var Doctype = GetListItems[0].DocumentFile.Url.split(".");
        Doctype = Doctype[Doctype.length - 1];
        $('#cropped-img').attr('src', this.GetDocImages(Doctype));
        $('a.Link').attr('href', GetListItems[0].DocumentFile.Url);
        $('#txtTitle').val(GetListItems[0].Title);
        //$('#txtDescription').val(GetListItems[0].Explanation);
        $('.richText-editor').html(GetListItems[0].Explanation),
          $('#txtDepartment').val(GetListItems[0].Departments);
        $('#DocumentTitle').text(listName);
      } else if (listName == "Corporate Discounts") {
        if (GetListItems[0].VendorLogo != null) {
          $('#View-img').attr('src', GetListItems[0].VendorLogo.Url);
        }
        else {
          $('#DivView-img').hide();
        }
        if (GetListItems[0].DocumentFile != null) {
          var Doctype = GetListItems[0].DocumentFile.Url.split(".");
          var DocName = GetListItems[0].DocumentFile.Url.split("/");
          Doctype = Doctype[Doctype.length - 1];
          DocName = DocName[DocName.length - 1];
          $('#cropped-img').attr('src', this.GetDocImages(Doctype));
          $('a.Link').attr('href', GetListItems[0].DocumentFile.Url);

        }
        else {
          $('#Divcropped-img').hide();
        }
        $('#txtTitle').val(GetListItems[0].Title);
        $('#txtSitelink').val(GetListItems[0].SiteLink.Url);
        $('#DocumentTitle').text(listName);
        $('#LblImage').text("Vendor Logo");
      } else if (listName == "Events") {
        $('#View-img').attr('src', GetListItems[0].Image.Url);
        $('#txtTitle').val(GetListItems[0].Title);
        $('.richText-editor').html(GetListItems[0].Explanation)
        $('#txtStartDate').val(new Date(GetListItems[0].StartDate).toLocaleDateString());
        if ((GetListItems[0].EndDate) == null || (GetListItems[0].EndDate) == "") {

          GetListItems[0].EndDate = "";
        }
        else {
          GetListItems[0].EndDate = new Date(GetListItems[0].EndDate).toLocaleDateString();
        }
        $('#txtEndDate').val(GetListItems[0].EndDate);
        $('#DocumentTitle').text(listName);
      } else if (listName == "Polls") {
        if (GetListItems[0].Options.split(';') != null) {
          let renderOptions = "";
          let arrOption = GetListItems[0].Options.split(';');
          arrOption = arrOption.filter(function (v) {
            return /\S/.test(v);
          });
          for (var j = 0; j < arrOption.length; j++) {
            renderOptions += arrOption[j] + "\n";
          }
          $('#txtQuestion').val(GetListItems[0].Question);
          $('#txtOptions').val(renderOptions);
          $('#DocumentTitle').text(listName);
        }
        else {
          $('#txtQuestion').val(GetListItems[0].Question);
          $('#txtOptions').val(GetListItems[0].Options);
          $('#DocumentTitle').text(listName);
        }
      } else if (listName == "Announcements") {
        $('#View-img').attr('src', GetListItems[0].Image.Url);
        $('#LblAnnounceTitle').text(GetListItems[0].Title);
        $('#LblAnnounceExpiryDate').text(new Date(GetListItems[0].Expires).toLocaleDateString());
        $('.richText-editor').html(GetListItems[0].Explanation),
          $('#DocumentTitle').text(listName);
        var ViewCount = await _this.GetViewCount(GetListItems[0].ViewedUsers);
        if (userflag == true) {
          $('.icon-eye').nextAll().remove();
          var node = $('.icon-eye').get(0).nextSibling;
          node.parentNode.removeChild(node);
          if (typeof (ViewCount) === undefined) {
            ViewCount = 1;
          }
          $('.icon-eye').after("Views <b>" + ViewCount + "</b>");

        }
      }
      if (listName == "Organizational Policies") {
        $('#ListViewBC').attr('href', this.context.pageContext.legacyPageContext.webAbsoluteUrl + '/Pages/OrganizationalPolicies.aspx');

      }
      else if (listName == "Corporate Discounts") {
        $('#ListViewBC').attr('href', this.context.pageContext.legacyPageContext.webAbsoluteUrl + '/Pages/Corporatediscounts.aspx');
      }
      else {
        $('#ListViewBC').attr('href', this.context.pageContext.legacyPageContext.webAbsoluteUrl + '/Pages/ListView.aspx?CName=' + listName);
      }
      $('#ListViewBC').text(listName + " List View");
      $('#ViewListItemBC').text("View " + listName);
    }
    else {
      var strconfirm = "There is no Data in the ID Specified.";
      alertify.confirm("Confirmation", strconfirm, function () {
        window.history.back();
      }, function () {
        //alertify.error("EVENT CANCELED");
      }).set('closable', false);
    }
    document.title = "View " + listName;
  }

  GetColumns(listName: string) {
    var Columns = [];
    switch (listName) {
      case "Announcements":
        Columns = ["Title", "Explanation", "Expires", "Image", "Display", "ViewedUsers"];
        break;
      case "Banners":
        Columns = ["Title", "Modified", "BannerContent", "Display", "Order", "Image", "LinkURL"];
        break;
      case "Polls":
        Columns = ["Title", "Modified", "Question", "Options"];
        break;
      case "Events":
        Columns = ["Title", "Modified", "StartDate", "EndDate", "Image", "Explanation"];
        break;
      case "Quick Links":
        Columns = ["Title", "Modified", "LinkURL", "Display"];
        break;
      case "Employee Corner":
        Columns = ["Title", "Modified", "Icon", "DocumentFile", "Display"];
        break;
      case "Organizational Policies":
        Columns = ["Title", "Modified", "DocumentFile", "Explanation", "Departments"];
        break;
      case "Corporate Discounts":
        Columns = ["Title", "Modified", "DocumentFile", "CorporateImage", "Hyperlink", "VendorLogo", "SiteLink"];
        break;
      case "Holiday":
        Columns = ["Title", "Modified", "EventDate", "EndEventDate", "Display"];
        break;
      case "News":
        Columns = ["Title", "Modified", "Date", "Image", "Explanation", "Display"];
        break;
    }
    return Columns;
  }

  GetDocImages(DocType: string) {
    var ImageURL;
    switch (DocType) {
      case "jpeg":
        ImageURL = this.context.pageContext.legacyPageContext.webAbsoluteUrl + "/_catalogs/masterpage/BloomHomepage/images/jpeg.png";
        break;
      case "ppt":
        ImageURL = this.context.pageContext.legacyPageContext.webAbsoluteUrl + "/_catalogs/masterpage/BloomHomepage/images/ppt.png";
        break;
      case "xls":
        ImageURL = this.context.pageContext.legacyPageContext.webAbsoluteUrl + "/_catalogs/masterpage/BloomHomepage/images/xls.png";
        break;
      case "doc":
        ImageURL = this.context.pageContext.legacyPageContext.webAbsoluteUrl + "/_catalogs/masterpage/BloomHomepage/images/doc.png";
        break;
      case "docx":
        ImageURL = this.context.pageContext.legacyPageContext.webAbsoluteUrl + "/_catalogs/masterpage/BloomHomepage/images/doc.png";
        break;
      case "pdf":
        ImageURL = this.context.pageContext.legacyPageContext.webAbsoluteUrl + "/_catalogs/masterpage/BloomHomepage/images/pdf.png";
        break;
    }
    return ImageURL;
  }

  renderhtml() {
    var listName = GetQueryStringParams("CName").replace("%20", " ");
    var renderhtmlImage = "";
    var renderimage = "";
    var renderDate = "";
    var renderTitle = "";
    var renderDescription = "";
    var renderEventDate = "";
    var renderHyperlink = "";
    var renderSitelink = "";
    var renderQuestion = "";
    var renderOptions = "";
    var renderAnnounceBtns = "";
    var renderExpiryDate = "";
    var renderAnnounceTitle = "";
    var renderAnnounceExpiry = "";
    var renderAnnounceLike = "";
    var renderAnnounceDesc = "";
    var renderAnnounceTabs = "";
    var renderAnnounceTabContent = "";
    var renderAnnounceCommentTab = "";
    var renderAnnounceCommentSubmit = "";
    var renderAnnounceCommentcontent = "";
    var renderDepartment = "";
    var requirednewrichTextEditor = "";
    var newrichTextEditor = "";
    var renderhtmlFile = "";
    var renderCorporateDiscountsearch = "";

    renderhtmlImage += '<div id="Divcropped-img" >' +
      '<img id="cropped-img" src="" class="img-responsive">' +
      '<div class="image-upload" id="Download-link">' +
      '<a class="Link pointer"  title="Delete" id="Dwnd-Link">' +
      '<i>Download</i>' +
      '</a>' +
      '</div>' +
      '</div>';
    renderhtmlFile += '<div id="Divcropped-img" >' +
      '<img id="cropped-img" src="" class="img-responsive">' +
      '<div class="image-upload" id="Download-link">' +
      '<a class="Link pointer"  title="Delete" id="Dwnd-Link">' +
      '<i>Download</i>' +
      '</a>' +
      '</div>' +
      '</div>';

    renderimage += '<div id="DivView-img" >' +
      '<label id="LblImage">Image</label>' +
      '<img id="View-img" src="" class="img-responsive">' +
      '</div>';

    renderDate += '<div class="input date">' +
      '<i class="icon-calenter"></i>' +
      '<label>Date</label>' +
      '<input class="form-control" id="txtDate" type="text" value="">' +
      '</div>';

    renderTitle += '<div class="input text">' +
      '<label>Title</label>' +
      '<input id="txtTitle" class="form-control" type="text" value="">' +
      '</div>';

    renderDepartment += '<div class="input text">' +
      '<label>Department</label>' +
      '<input id="txtDepartment" class="form-control" type="text" value="">' +
      '</div>';

    renderDescription += '<div class="input textarea">' +
      '<label>Description</label>' +
      '<textarea id="txtDescription" class="form-control"></textarea>' +
      '</div>';

    requirednewrichTextEditor += "<div class='textarea input' style='cursor:not-allowed;pointer-events:none;'>" +
      "<label class='control-label'>Description</label>" +
      "<textarea id='txtrequiredDescription' class='form-control content'></textarea>" +
      "</div>";

    newrichTextEditor += "<div class='textarea input' style='cursor:not-allowed;pointer-events:none;'>" +
      "<label>Description</label>" +
      "<textarea id='txtDescription' class='form-control content'></textarea>" +
      "</div>";

    renderEventDate += '<div class="input date">' +
      '<i class="icon-calenter"></i>' +
      '<label>StartDate</label>' +
      '<input class="form-control" id="txtStartDate" type="text" value="">' +
      '</div>' +
      '<div class="input date">' +
      '<i class="icon-calenter"></i>' +
      '<label>EndDate</label>' +
      '<input class="form-control" id="txtEndDate" type="text" value="">' +
      '</div>';

    renderHyperlink += '<div id="DivHyperLink" class="input text">' +
      '<label>Link URL</label>' +
      '<input id="txtHyperlink" class="form-control" type="text" value="">' +
      '</div>';


    renderSitelink += '<div class="input text">' +
      '<label>Site Link</label>' +
      '<input id="txtSitelink" class="form-control" type="text" value="">' +
      '</div>';

    renderQuestion += '<div class="input text">' +
      '<label>Question</label>' +
      '<input id="txtQuestion" class="form-control" type="text" value="">' +
      '</div>';

    renderOptions += '<div class="input textarea">' +
      '<label>Options</label>' +
      '<textarea id="txtOptions" class="form-control"></textarea>' +
      '</div>';

    renderAnnounceBtns += '<div class="button-field save-button pointer">' +
      '<a  class="delete-icon close-icon pointer" id="closeicon" title="Delete"><i class="commonicon-close"></i>Close</a>' +
      '</div>';

    renderAnnounceTitle += '<h3 id="LblAnnounceTitle"></h3>';

    renderAnnounceExpiry += '<h5 id="LblAnnounceExpiryDate"></h5>';

    renderAnnounceDesc += '<h5 id="LblAnnounceDesc"></h5>';

    renderAnnounceLike += '<div id="divHeartCheck" class="comment-div" style="padding-top:6px">' +
      '<a  class="pointer" id="aHeartCheck"><i class="icon-heart"></i> Liked </a></li>' +
      '</div>';

    renderExpiryDate += '<div class="input date">' +
      '<i class="icon-calenter"></i>' +
      '<label>Expiry Date</label>' +
      '<input class="form-control" id="txtExpiryDate" type="text" value="">' +
      '</div>';

    renderAnnounceTabs += '<ul class="nav nav-tabs" role="tablist">' +
      '<li role="presentation" class="active"><a class="pointer" href="#View" aria-controls="View" role="tab"><i class="icon-eye"></i> Views <b></b> </a></li>' +
      '<li role="presentation"><a class="pointer" href="#comments" aria-controls="comments" role="tab" data-toggle="tab"><i class="icon-comments"></i> Comments <b></b>  </a></li>' +
      '<li role="presentation"><a class="pointer" href="#like" aria-controls="like" role="tab" data-toggle="tab"><i class="icon-heart"></i>Likes <b></b> </a></li>' +
      '</ul>';

    renderAnnounceTabContent += '<div class="tab-content">' +
      '<div role="tabpanel" class="tab-pane active" id="View">' +
      '<ul id="View-Tab-Cmts">' +
      '</ul></div>' +
      '<div role="tabpanel" class="tab-pane" id="comments">' +
      '<ul id="Comments-Tab-Cmts">' +
      '</ul></div>' +
      '<div role="tabpanel" class="tab-pane" id="like">' +
      '<ul id="Like-Tab-Cmts">' +
      '</ul></div>' +
      '</div>';

    renderAnnounceCommentTab += '<div class="card col-md-12">' +
      '<ul class="nav nav-tabs" role="tablist">' +
      '<li role="presentation" class="active"><a href="#comments" aria-controls="comments" role="tab" data-toggle="tab"><i class="icon-comments"></i>Comments <b> 0 </b></a></li>' +
      '</ul>' +
      '<div class="tab-content">' +
      '<div role="tabpanel" class="tab-pane active" id="comments">' +
      '<ul id="AnnounceComments">' +
      '</ul>' +
      '</div>' +
      '</div>' +
      '</div>';

    renderAnnounceCommentSubmit += '<div role="tabpanel" class="tab-pane active" id="comments">' +
      '<div class="col-md-12"> <textarea style="height:100px !important;" id="txtAnnounceDesc" class="form-control form-group"></textarea>' +
      '</div>' +
      '<div class="col-md-12">' +
      '<div class="button-field save-button pull-right">' +
      '<a class="cmt-save pointer" id="Submit-Comments" title="Add New">Submit</a>' +
      '</div>' +
      '</div>' +
      '</div>';

    renderAnnounceCommentcontent += '<ul class="nav nav-tabs" role="tablist">' +
      '<li role="presentation" class="active"><a class="pointer" href="#comments" aria-controls="comments" role="tab" data-toggle="tab"><i class="icon-commentss"></i> Post Comments </a></li>' +
      '</ul>';

    renderCorporateDiscountsearch += '<input id="corporateSearch" class="CorporateDiscountsearch" type="text" placeholder="Search.." name="search">' +
      '<button type="submit"><i class="fa fa-search"></i></button>';

    if (listName == "Banners") {
      $('.form-imgsec').append(renderimage);
      $('#Form-Part').append(renderTitle + requirednewrichTextEditor + renderHyperlink);
      //$('#Download-link').hide();
    } else if (listName == "Holiday") {
      $('#Form-Part').append(renderTitle + renderEventDate);
      $('#Img-Part').hide();
    } else if (listName == "News") {
      $('.form-imgsec').append(renderimage);
      $('#Form-Part').append(renderDate + renderTitle + renderDescription);
    } else if (listName == "Quick Links") {
      $('#Form-Part').append(renderTitle + renderHyperlink);
      $('#Img-Part').hide();
    } else if (listName == "Employee Corner") {
      $('.form-imgsec').append(renderhtmlFile);
      $('#Form-Part').append(renderTitle + renderDate);
    } else if (listName == "Organizational Policies") {
      $('.form-imgsec').append(renderhtmlFile);
      $('#Form-Part').append(renderTitle + renderDepartment + newrichTextEditor);
    } else if (listName == "Corporate Discounts") {

      $('.form-imgsec').append(renderimage);
      $('#Form-Part').append(renderTitle + renderSitelink + renderhtmlFile);
    } else if (listName == "Events") {
      $('.form-imgsec').append(renderimage);
      $('#Form-Part').append(renderTitle + newrichTextEditor + renderEventDate);
    } else if (listName == "Polls") {
      $('#Form-Part').append(renderQuestion + renderOptions);
      $('#Img-Part').hide();
    } else if (listName == "Announcements") {
      $('.form-imgsec').append(renderimage);
      //$('#DocumentTitle').before(renderAnnounceBtns);
      $('#Form-Part').append(renderAnnounceTitle + renderAnnounceExpiry + requirednewrichTextEditor);
      $('#Announcement-Sec').show();

      var _this = this;
      var Email = this.context.pageContext.user.email;

      checkUserinGroup(listName, Email, function (result) {
        console.log(result);

        if (result == 1) {
          $('.card').append(renderAnnounceTabs + renderAnnounceTabContent);
          userflag = true;
          _this.GetLikesCount();
        }
        else {
          userflag = false;
          $('#DivView-img').after(renderAnnounceLike);
          $('#Viewer-Tab').before(renderAnnounceCommentcontent);
          $('#Viewer-Tab').append(renderAnnounceCommentSubmit + renderAnnounceCommentTab);

          _this.GetLikes();
          _this.GetLikesCount();
          let SubmitCmtevent = document.getElementById('Submit-Comments');
          SubmitCmtevent.addEventListener("click", (e: Event) => _this.SubmitComments());
          let SubmitLikeevent = document.getElementById('aHeartCheck');
          SubmitLikeevent.addEventListener("click", (e: Event) => _this.SubmitLikes());
        }

        _this.GetComments(userflag);
      })
    }
    $('#DocumentTitle').before(renderAnnounceBtns);
    $('#Form-Part :input').attr("disabled", "true");
    let Closeevent = document.getElementById('closeicon');
    Closeevent.addEventListener("click", (e: Event) => window.history.back());
    $('.content').richText();
    $('.richText-toolbar').css("display", "none");
  }


  SubmitComments() {


    var $body = $('body');

    if ($('.ajs-message').length > 0) {
      $('.ajs-message').remove();
    }

    if (this.AnnouncementValidation()) {
      var ItemID = GetQueryStringParams("CID");
      var listName = GetQueryStringParams("CName").replace("%20", " ");

      var siteURL = this.context.pageContext.web.absoluteUrl;
      var _this = this;
      var txtAnnounceDesc = $.trim($("#txtAnnounceDesc").val()).length;
      if (txtAnnounceDesc != 0) {
        let myobjHol = {
          Comments: $('#txtAnnounceDesc').val(),
          AnnouncementID: ItemID
        }
        $body.addClass("loading");
        let AddComments = addItems("AnnouncementComments", myobjHol);

        AddComments.then(result => {
          // if () {

          $('#AnnounceComments').html("");
          $('#txtAnnounceDesc').val("");
          _this.GetComments(false);
          $body.removeClass("loading");
          // } else {
          //   $body.removeClass("loading");
          //   console.log(result);
          // }

        });

      }
      else {

      }
    }

  }

  GetComments(filterKey?: boolean) {
    var ItemID = GetQueryStringParams("CID");
    var $body = $("body");
    var siteURL = this.context.pageContext.web.absoluteUrl;
    var Columns = ["Comments", "Editor/Title", "ID"];
    let GetComments = readItem("AnnouncementComments", Columns, 50, "Modified", "AnnouncementID", ItemID, "Editor");
    var Html = "";
    var HtmlDelComments = "";
    var _this = this;
    GetComments.then((items: any) => {
      for (var i = 0; i < items.length; i++) {
        Html += '<li><h6>' + items[i].Editor.Title + '</h6>' + items[i].Comments + '</li>';
        HtmlDelComments += '<li><h6>' + items[i].Editor.Title + '</h6>' + items[i].Comments + '<a  id="' + items[i].ID + '" data-value="' + items[i].ID + '" class="icon-delete pointer"></a></li>';
      }
      if (filterKey == false) {
        $('#AnnounceComments').append(Html);
      }
      else {
        //$('#View-Tab-Cmts').append(Html);
        $("#Comments-Tab-Cmts").append(HtmlDelComments);
        let DeleteCmtevent = document.getElementsByClassName('icon-delete');
        for (let i = 0; i < DeleteCmtevent.length; i++) {
          DeleteCmtevent[i].addEventListener("click", (e: Event) => _this.DeleteComments(DeleteCmtevent[i].id));
        }
      }

      $('.icon-comments').nextAll().remove();
      var node = $('.icon-comments').get(0).nextSibling;
      node.parentNode.removeChild(node);
      $('.icon-comments').after("Comments <b>" + items.length + "</b>");
    });

  }

  DeleteComments(id: string) {
    var strconfirm = "Are you sure you want to delete selected Comment ?";
    var _this = this;
    alertify.confirm('Confirmation', strconfirm, function () {
      //var _this = this;
      var ItemID = GetQueryStringParams("CID");
      var listName = GetQueryStringParams("CName").replace("%20", " ");
      var $body = $("body");
      let CommentItemID = parseInt(id, 10);
      let DeleteListItems = deleteItem("AnnouncementComments", CommentItemID);
      //window.location.href = _this.context.pageContext.legacyPageContext.webAbsoluteUrl + '/Pages/Viewlistitem.aspx?CName=' + listName + '&CID=' + ItemID;
      DeleteListItems.then(result => {
        $('#Comments-Tab-Cmts').html("");
        $('#txtAnnounceDesc').val("");
        _this.GetComments(true);
      });

    }, function () {
    }).set('closable', false);
  }




  async SubmitLikes() {
    var ItemID = GetQueryStringParams("CID");
    var listName = "AnnouncementsLikes";
    var $body = $("body");
    var _this = this;
    var Columns = ["User", "Liked", "AnnouncementID", "ID"];
    var matchColumns = formString(Columns);
    var filterValue = this.context.pageContext.user.email;
    let CommentItemID = parseInt(ItemID, 10);
    let GetLikes = await pnp.sp.web.lists.getByTitle(listName).items.select(matchColumns).filter("User eq '" + filterValue + "' and AnnouncementID eq '" + ItemID + "'").top(1).orderBy("Modified").get();
    if (GetLikes.length == 0) {
      $('#divHeartCheck').addClass("heart-check");
      $body.addClass("loading");
      let myobjHol = {
        User: this.context.pageContext.user.email,
        AnnouncementID: CommentItemID,
        Liked: true
      }


      let AddLikes = addItems(listName, myobjHol);

      AddLikes.then(result => {
        $('#divHeartCheck').addClass("heart-check");
        _this.GetLikesCount();
        $body.removeClass("loading");
      })
    }
    else {
      if (GetLikes[0].Liked == true) {

        $('#divHeartCheck').removeClass("heart-check");
        $body.addClass("loading");
        let myobjHol = {
          Liked: false
        }

        let updateLikes = await updateItem(listName, GetLikes[0].ID, myobjHol);


        if (updateLikes.data) {
          $('#divHeartCheck').removeClass("heart-check");
          _this.GetLikesCount();
          $body.removeClass("loading");
        } else {
          $body.removeClass("loading");
          console.log(updateLikes);
        }
      }
      else if (GetLikes[0].Liked == false) {
        $('#divHeartCheck').addClass("heart-check");
        $body.addClass("loading");
        let myobjHol = {
          Liked: true
        }

        let updateLikes = await updateItem(listName, GetLikes[0].ID, myobjHol);

        if (updateLikes.data) {
          $('#divHeartCheck').addClass("heart-check");
          _this.GetLikesCount();
          $body.removeClass("loading");
        } else {
          $body.removeClass("loading");
          console.log(updateLikes);
        }
      }
    }
  }

  async GetLikes() {
    var ItemID = GetQueryStringParams("CID");
    var listName = "AnnouncementsLikes";
    var Columns = ["User", "Liked", "AnnouncementID", "ID"];
    var matchColumns = formString(Columns);
    var filterValue = this.context.pageContext.user.email;
    let GetLikes = await pnp.sp.web.lists.getByTitle(listName).items.select(matchColumns).filter("User eq '" + filterValue + "' and AnnouncementID eq '" + ItemID + "'").top(1).orderBy("Modified").get();

    if (GetLikes.length != 0) {
      if (GetLikes[0].Liked == true) {
        $('#divHeartCheck').addClass("heart-check");
      }
      else {
        $('#divHeartCheck').removeClass("heart-check");
      }
    }
  }
  nullDateValidate(nullDate) {
    var exdate = new Date(nullDate);
    var day = ("0" + exdate.getDate()).slice(-2);
    var month = ("0" + (exdate.getMonth() + 1)).slice(-2);
    var expiredate = exdate.getFullYear() + "/" + (month) + "/" + (day);
    return expiredate;
  }
  async GetLikesCount() {
    var ItemID = GetQueryStringParams("CID");
    var listName = "AnnouncementsLikes";
    var Columns = ["User", "Liked", "AnnouncementID", "ID", "Editor/Title"];
    var matchColumns = formString(Columns);
    var Html = "";

    let GetLikeCount = await pnp.sp.web.lists.getByTitle(listName).items.select(matchColumns).expand("Editor").filter("AnnouncementID eq '" + ItemID + "' and Liked eq 1").orderBy("Modified").get();
    $('.icon-heart').nextAll().remove();
    var node = $('.icon-heart').get(0).nextSibling;
    node.parentNode.removeChild(node);
    $('.icon-heart').after("Likes <b>" + GetLikeCount.length + "</b>");

    for (var i = 0; i < GetLikeCount.length; i++) {
      Html += '<li><h6>' + GetLikeCount[i].Editor.Title + '</h6></li>';
    }
    $('#Like-Tab-Cmts').html("");
    $('#Like-Tab-Cmts').append(Html);
  }



  async GetViewCount(Users: string) {

    var Email = this.context.pageContext.user.email;
    var ItemID = parseInt(GetQueryStringParams("CID"));
    var _this = this;
    var $body = $("body");
    var usercount;
    var UserID = await pnp.sp.site.rootWeb.ensureUser(Email).then(result => {
      return result.data.Id;
    });

    var ViewedUsers;
    if (Users != null) {
      ViewedUsers = Users.split(",");
      if (ViewedUsers.indexOf(UserID.toString()) > -1) {
        return ViewedUsers.length;
      }
      else {
        var user = Users + "," + UserID.toString();
        usercount = user.split(",");
        let myobjHol = {
          ViewedUsers: user
        }

        let ViewCountRet = await updateItem("Announcements", ItemID, myobjHol);
        return ViewCountRet.length;
      }
    }
    else {
      var user1 = UserID.toString();
      usercount = 1;
      let myobjHol = {
        ViewedUsers: user1
      }

      let ViewCountRet = await updateItem("Announcements", ItemID, myobjHol);
      return usercount;
    }

  }


  AnnouncementValidation() {
    if (!$('#txtAnnounceDesc').val().trim()) {
      alertify.set('notifier', 'position', 'top-right');
      alertify.error("Please Enter Comments");
      return false;
    }
    return true;
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
