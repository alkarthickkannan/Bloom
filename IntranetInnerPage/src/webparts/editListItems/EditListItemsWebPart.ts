import {
  Version
} from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import {
  escape
} from '@microsoft/sp-lodash-subset';

//import styles from './EditListItemsWebPart.module.scss';
import * as strings from 'EditListItemsWebPartStrings';
import {
  SPComponentLoader
} from '@microsoft/sp-loader';

require('bootstrap');
import pnp from 'sp-pnp-js';
import '../../ExternalRef/css/cropper.min.css'
import '../../ExternalRef/css/cropper.css';

import '../../ExternalRef/js/cropper-main.js';
import '../../ExternalRef/js/cropper.min.js';

import '../../ExternalRef/css/bootstrap-datepicker.min.css';
require('../../ExternalRef/js/bootstrap-datepicker.min.js');
import {
  readItems,
  updateitems,
  GetQueryStringParams,
  additemsimage,
  base64ToArrayBuffer
} from '../../commonService';

import 'jquery';

export interface IEditListItemWebPartProps {
  description: string;
}
declare var $;
declare var alertify: any;
declare var datepicker: any;
var siteURL = this.context.pageContext.web.absoluteUrl;
export default class EditListItemWebPart extends BaseClientSideWebPart<IEditListItemWebPartProps> {
  strcropstorage = "";
  imageValue = 0;
  imgsrc;
  siteURL = "";
  public render(): void {
      this.siteURL = this.context.pageContext.web.absoluteUrl;
      var strLocalStorage = GetQueryStringParams("CName");
      strLocalStorage = strLocalStorage.split('%20').join(' ');
      var strLocalStorageBreadcrumb = GetQueryStringParams("CName");
      strLocalStorageBreadcrumb = strLocalStorageBreadcrumb.split("%20").join(' ');
      let sourceComponent = "";
      this.domElement.innerHTML =
          "<div class='breadcrumb'>" +
          "<ol>" +
          "<li><a href='" + this.siteURL + "/Pages/Home.aspx' title='Home'>Home</a></li>" +
          "<li><a class='pointer' id='breadTilte' title='" + strLocalStorage + "'>" + strLocalStorage + " List View</a></li>" +
          "<li><span>Edit " + strLocalStorage + "</span></li>" +
          "</ol>" +
          "</div>" +
          "<div class='title-section'>" +
          "<div class='button-field save-button'>" +
          "<a class='addbutton pointer' title='Update Item' id='UpdateItem'><i class='commonicon-save addbutton'></i>Save</a>" +
          "<a class='delete-icon close-icon pointer' class='closebutton'  title='Close' id='DelItem'><i class='commonicon-close closebutton'></i>Close</a>" +
          "</div>" +
          "<h2 id='ComponentName'>Announcements</h2>" +
          "</div>" +
          "<div class='form-section required'>" +
          "</div>" +
          "<div class='modal'><!-- Place at bottom of page --></div>";

      document.title = 'Edit' + strLocalStorage;
      document.getElementById("ComponentName").innerHTML = GetQueryStringParams("CName").split("%20").join(" ");
      // var strLocalStorage = GetQueryStringParams("CName");
      //strLocalStorage = strLocalStorage.split('%20').join(' ');
      var strComponentId = GetQueryStringParams("CID");
      this.renderhtml(strComponentId);
      let Addevent = document.getElementById('UpdateItem');
      // Addevent.addEventListener("click", (e: Event) => this.UpdateItem(siteURL, strComponentId));
      //for (let i = 0; i < Addevent.length; i++) {
      Addevent.addEventListener("click", (e: Event) => this.UpdateItem(this.siteURL, strLocalStorage, strComponentId));
      //}
      let breadTilte = document.getElementById('breadTilte');
      //for (let i = 0; i < Addevent.length; i++) {
      breadTilte.addEventListener("click", (e: Event) => this.pageBack());

      let Closeevent = document.getElementById('DelItem');
      // Addevent.addEventListener("click", (e: Event) => this.UpdateItem(siteURL, strComponentId));
      //for (let i = 0; i < Closeevent.length; i++) {
      Closeevent.addEventListener("click", (e: Event) => this.pageBack());
  }

  // protected get dataVersion(): Version {
  //     return Version.parse('1.0');
  // }
  pageBack() {
      window.history.back();
  }

  DateChecker() {
      if (Date.parse($("#txtEvDate").val()) > Date.parse($("#txtEEDate").val())) {
          alertify.set('notifier', 'position', 'top-right');
          alertify.error("Start date must be less than End date");
          return false;
      } else {
          return true;
      }
  }
  EventDateChecker() {
      if ($('#txtEndDate').val() == "") {
          return true;
      }
      else if (Date.parse($("#txtStartDate").val()) > Date.parse($("#txtEndDate").val())) {
          alertify.set('notifier', 'position', 'top-right');
          alertify.error("Start date must be less than End date");
          return false;
      }
      return true;

  }
  announcementsValidtion() {
      if (!$('#txtExpires').val()) {
          alertify.set('notifier', 'position', 'top-right');
          alertify.error("Please Select the Date");
          return false;
          //isAllfield = false;
      } else if (!$('#txtTitle').val().trim()) {
          alertify.set('notifier', 'position', 'top-right');
          alertify.error("Please Enter Title");
          return false;
          // isAllfield = false;
      } else if (!$('#txtrequiredDescription').val().trim()) {
          alertify.set('notifier', 'position', 'top-right');
          alertify.error("Please Enter Description");
          return false;
          //isAllfield = false;
      }
      return true;
  }
  holidaysValidtion() {
      if (!$('#txtTitle').val().trim()) {
          alertify.set('notifier', 'position', 'top-right');
          alertify.error("Please Enter Title");
          return false;
      } else if (!$('#txtEvDate').val()) {
          alertify.set('notifier', 'position', 'top-right');
          alertify.error("Please Select Start Date");
          return false;
      } else if (Date.parse($("#txtEvDate").val()) > Date.parse($("#txtEEDate").val())) {
          alertify.set('notifier', 'position', 'top-right');
          alertify.error("Start Date must be less than End Date");
          return false;
      }

      return true;
  }
  quickLinksValidation() {
      var regexp = /^(https?|s?ftp):\/\/(((([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(%[\da-f]{2})|[!\$&'\(\)\*\+,;=]|:)*@)?(((\d|[1-9]\d|1\d\d|2[0-4]\d|25[0-5])\.(\d|[1-9]\d|1\d\d|2[0-4]\d|25[0-5])\.(\d|[1-9]\d|1\d\d|2[0-4]\d|25[0-5])\.(\d|[1-9]\d|1\d\d|2[0-4]\d|25[0-5]))|((([a-z]|\d|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(([a-z]|\d|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])*([a-z]|\d|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])))\.)+(([a-z]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(([a-z]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])*([a-z]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])))\.?)(:\d*)?)(\/((([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(%[\da-f]{2})|[!\$&'\(\)\*\+,;=]|:|@)+(\/(([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(%[\da-f]{2})|[!\$&'\(\)\*\+,;=]|:|@)*)*)?)?(\?((([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(%[\da-f]{2})|[!\$&'\(\)\*\+,;=]|:|@)|[\uE000-\uF8FF]|\/|\?)*)?(#((([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(%[\da-f]{2})|[!\$&'\(\)\*\+,;=]|:|@)|\/|\?)*)?$/i

      if (!$('#txtTitle').val().trim()) {
          alertify.set('notifier', 'position', 'top-right');
          alertify.error("Please Enter Title");
          return false;
      } else if (!$('#txtHyper').val().trim()) {
          alertify.set('notifier', 'position', 'top-right');
          alertify.error("Please Enter Link URL");
          return false;
      } else if (!regexp.test($('#txtHyper').val().trim())) {
          $('#txtHyper').focus();
          alertify.set('notifier', 'position', 'top-right');
          alertify.error("Please Enter Link URL Correctly");
          return false;
      }
      return true
  }
  nullDateValidate(nullDate) {
      var exdate = new Date(nullDate);
      var day = ("0" + exdate.getDate()).slice(-2);
      var month = ("0" + (exdate.getMonth() + 1)).slice(-2);
      var expiredate = exdate.getFullYear() + "/" + (month) + "/" + (day);
      return expiredate;
  }
  newsValidation() {
      if (!$('#txtExpires').val()) {
          alertify.set('notifier', 'position', 'top-right');
          alertify.error("Please Select the Date");
          return false;
      } else if (!$('#txtTitle').val().trim()) {
          alertify.set('notifier', 'position', 'top-right');
          alertify.error("Please Enter Title");
          return false;
      } else if (!$('#txtrequiredDescription').val().trim()) {
          alertify.set('notifier', 'position', 'top-right');
          alertify.error("Please Enter Description");
          return false;
      }
      return true;
  }
  quickReadsValidation() {
      if (!$('#txtTitle').val().trim()) {
          alertify.set('notifier', 'position', 'top-right');
          alertify.error("Please Enter Title");
          return false;
      } else if (!$('#uploadFile').val()) {
          $('#uploadFile').focus();
          alertify.set('notifier', 'position', 'top-right');
          alertify.error("Please Select Document File");
          return false;

      }
      return true
  }
  eventsValidation() {
      if (!$('#txtTitle').val().trim()) {
          alertify.set('notifier', 'position', 'top-right');
          alertify.error("Please Enter Title");
          return false;
      }
      else if (!$('#txtrequiredDescription').val().trim()) {
          alertify.set('notifier', 'position', 'top-right');
          alertify.error("Please Enter Description");
          return false;
      } else if (!$('#txtStartDate').val()) {
          alertify.set('notifier', 'position', 'top-right');
          alertify.error("Please Select Start Date");
          return false;
      } else if (Date.parse($("#txtStartDate").val()) > Date.parse($("#txtEndDate").val())) {
          alertify.set('notifier', 'position', 'top-right');
          alertify.error("Start date must be less than End date");
          return false;
      }
      return true;

  }
  orgpolicyValidation(isAllfield) {
      if (!$('#txtTitle').val().trim()) {
          alertify.set('notifier', 'position', 'top-right');
          alertify.error("Please Enter Title");
          isAllfield = false;
      } else if (!$('#txtDescription').val().trim()) {
          alertify.set('notifier', 'position', 'top-right');
          alertify.error("Please Enter Description");
          isAllfield = false;
      }
  }
  bannersValidation() {
      if (!$('#inputImage').val()) {
          alertify.set('notifier', 'position', 'top-right');
          alertify.error("Please Select Image");
          return false;
          // isAllfield = false;

      } else if (!$('#txtTitle').val().trim()) {
          alertify.set('notifier', 'position', 'top-right');
          alertify.error("Please Enter the Title");
          return false;
          //isAllfield = false;
      }
      return true;
  }
  pollsValidation() {
      var optionseperate = $('#txtOptions').val();
      var resultarray = optionseperate.split(";");
      var newArray = resultarray.filter(function (v) {
          return v !== ' '
      });
      if (!$('#txtQuestion').val().trim()) {
          alertify.set('notifier', 'position', 'top-right');
          alertify.error("Please Enter the Question");
          return false;
      } else if (!$('#txtOptions').val().trim() || newArray.length <= 1) {
          alertify.set('notifier', 'position', 'top-right');
          alertify.error("Please Enter the Answers Correctly");
          return false;
      }
      else if (!$('#txtOptions').val().trim() || newArray.length >= 5) {
          alertify.set('notifier', 'position', 'top-right');
          alertify.error("Maximum Five Answers only Allowed");
          return false;
      }
      return true;
  }
  corporationValidation() {
      var regexp = /^(https?|s?ftp):\/\/(((([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(%[\da-f]{2})|[!\$&'\(\)\*\+,;=]|:)*@)?(((\d|[1-9]\d|1\d\d|2[0-4]\d|25[0-5])\.(\d|[1-9]\d|1\d\d|2[0-4]\d|25[0-5])\.(\d|[1-9]\d|1\d\d|2[0-4]\d|25[0-5])\.(\d|[1-9]\d|1\d\d|2[0-4]\d|25[0-5]))|((([a-z]|\d|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(([a-z]|\d|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])*([a-z]|\d|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])))\.)+(([a-z]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(([a-z]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])*([a-z]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])))\.?)(:\d*)?)(\/((([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(%[\da-f]{2})|[!\$&'\(\)\*\+,;=]|:|@)+(\/(([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(%[\da-f]{2})|[!\$&'\(\)\*\+,;=]|:|@)*)*)?)?(\?((([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(%[\da-f]{2})|[!\$&'\(\)\*\+,;=]|:|@)|[\uE000-\uF8FF]|\/|\?)*)?(#((([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(%[\da-f]{2})|[!\$&'\(\)\*\+,;=]|:|@)|\/|\?)*)?$/i
      if (!$("#txtTitle").val().trim()) {
          alertify.set('notifier', 'position', 'top-right');
          alertify.error("Please Enter the Title");
          return false;
      } else if (!$("#txtsitelink").val().trim()) {
          alertify.set('notifier', 'position', 'top-right');
          alertify.error("Please Enter the Site Link");
          return false;
      } else if (!regexp.test($('#txtsitelink').val().trim())) {
          $('#txtsitelink').focus();
          alertify.set('notifier', 'position', 'top-right');
          alertify.error("Please Enter SiteLink Correctly");
          return false;


      }
      return true
  }
  imagecropperChecking() {
      if ($('#canvasdisplay').css('display') == 'block') {
          alertify.set('notifier', 'position', 'top-right');
          alertify.error("Please save the cropper Image First");
          return false;
      }
      return true;
  }
  UpdateItem(siteURL, strLocalStorage, strComponentId) {
      var $body = $('body');
      if ($('.ajs-message').length > 0) {
          $('.ajs-message').remove();
      }
      var that = this;
      let strcrop = localStorage.getItem("crop");
      var count;
      // this.imageValue=0;
      let objResults;
      var $body = $("body");
      var isAllfield = true;

      if (strLocalStorage == "Announcements") {
          var files = <HTMLInputElement>document.getElementById("inputImage");
          let file = files.files[0];

          if (strcrop == "1" && files.files.length == 0) {
              var saveData = {
                  Title: $("#txtTitle").val(),
                  Explanation: $("#txtrequiredDescription").val(),
                  Expires: $('#txtExpires').val(),
                  Image: {
                      "__metadata": {
                          "type": "SP.FieldUrlValue"
                      },
                      Url: siteURL + "/_catalogs/masterpage/Bloom/images/logo.png"
                  }

              };
              isAllfield = this.announcementsValidtion();
              let cropchecker = this.imagecropperChecking();
              if (isAllfield && cropchecker) {
                  $body.addClass("loading");
                  updateitems("Announcements", strComponentId, saveData, function (e) {

                      if (e.data) {
                          $body.removeClass("loading");
                          that.pageBack();
                          localStorage.clear();

                          //window.location.href = "" + siteURL + "/Pages/ListView.aspx?CName=" + strLocalStorage + "";

                      } else {
                          $body.removeClass("loading");
                          console.log(e);
                      }
                  });
                  strcrop = "0";
              }

          }
          else if (files.files.length == 0) {

              //   this.announcementsValidtion(isAllfield)
              var saveDatas = {
                  Title: $("#txtTitle").val(),
                  Explanation: $("#txtrequiredDescription").val(),
                  Expires: $('#txtExpires').val(),

              };
              isAllfield = this.announcementsValidtion();
              let cropchecker = this.imagecropperChecking();
              if (isAllfield && cropchecker) {
                  $body.addClass("loading");
                  updateitems("Announcements", strComponentId, saveDatas, function (e) {

                      if (e.data) {
                          $body.removeClass("loading");
                          that.pageBack();
                          //window.location.href = "" + siteURL + "/Pages/ListView.aspx?CName=" + strLocalStorage + "";

                      } else {
                          $body.removeClass("loading");
                          console.log(e);
                      }
                  });
                  strcrop = "0";
              }

          }

          else {
              isAllfield = this.announcementsValidtion();
              let cropchecker = this.imagecropperChecking();
              if (isAllfield && cropchecker) {
                  var fileURL = window.location.origin;
                  $body.addClass("loading");
                  //var filename = $('#inputImage').val().replace(/C:\\fakepath\\/i, '');
                  var uniquename = Math.random().toString(36).substr(2, 9) + ".png";
                  var file1 = $('#cropped-img').attr('src').split("base64,");
                  var blob = base64ToArrayBuffer(file1[1]);
                  pnp.sp.web.getFolderByServerRelativeUrl("PublishingImages").files.add(file.name, blob, true)
                      .then(function (result) {
                          pnp.sp.web.lists.getByTitle("Announcements").items.getById(strComponentId).update({
                              ID: strComponentId,
                              Title: $('#txtTitle').val().trim(),
                              Expires: new Date($('#txtExpires').val()),
                              Explanation: $("#txtrequiredDescription").val(),
                              Image: {
                                  "__metadata": {
                                      "type": "SP.FieldUrlValue"
                                  },
                                  Url: fileURL + result.data.ServerRelativeUrl
                              }
                          }).then(r => {
                              $body.removeClass("loading");
                              that.pageBack();
                              //window.location.href = "" + siteURL + "/Pages/ListView.aspx?CName=" + strLocalStorage + "";
                          });

                      });
              }
          }
          // strcrop="0";
      } else if (strLocalStorage == "Holiday") {
          isAllfield = this.holidaysValidtion();

          let myobjHol = {
              Title: $("#txtTitle").val(),
              EndEventDate: $("#txtEEDate").val(),
              EventDate: $("#txtEvDate").val()
          }
          let isDateChecker = this.DateChecker();
          this.imagecropperChecking();
          if (isAllfield && isDateChecker) {
              $body.addClass("loading");
              updateitems("Holiday", strComponentId, myobjHol, function (e) {
                  if (e.data) {
                      $body.removeClass("loading");
                      that.pageBack();
                      //window.location.href = "" + siteURL + "/Pages/ListView.aspx?CName=" + strLocalStorage + "";

                  } else {
                      $body.removeClass("loading");
                      console.log(e);
                  }
              });
          }
      } else if (strLocalStorage == "Quick Links") {
          isAllfield = this.quickLinksValidation();
          let myobjQl = {
              Title: $("#txtTitle").val(),
              LinkURL: {
                  "__metadata": {
                      "type": "SP.FieldUrlValue"
                  },
                  Url: $('#txtHyper').val()
              }
          }
          this.imagecropperChecking();
          if (isAllfield) {
              $body.addClass("loading");
              isAllfield = this.quickLinksValidation();
              updateitems("Quick Links", strComponentId, myobjQl, function (e) {

                  if (e.data) {
                      $body.removeClass("loading");
                      that.pageBack();
                      //window.location.href = "" + siteURL + "/Pages/ListView.aspx?CName=" + strLocalStorage + "";

                  } else {
                      $body.removeClass("loading");
                      console.log(e);
                  }
              });
          }
      } else if (strLocalStorage == "News") {
          var files = <HTMLInputElement>document.getElementById("inputImage");
          let file = files.files[0];
          if (strcrop == "1" && files.files.length == 0) {
              isAllfield = this.newsValidation()
              let saveData = {
                  Title: $("#txtTitle").val(),
                  Date: $("#txtExpires").val(),
                  Explanation: $("#txtrequiredDescription").val(),
                  Image: {
                      "__metadata": {
                          "type": "SP.FieldUrlValue"
                      },
                      Url: siteURL + "/_catalogs/masterpage/Bloom/images/logo.png"
                  }
              }
              let cropchecker = this.imagecropperChecking();
              if (isAllfield && cropchecker) {
                  $body.addClass("loading");
                  updateitems("News", strComponentId, saveData, function (e) {

                      if (e.data) {
                          $body.removeClass("loading");
                          that.pageBack();
                          //window.location.href = "" + siteURL + "/Pages/ListView.aspx?CName=" + strLocalStorage + "";
                          localStorage.clear();
                      } else {
                          $body.removeClass("loading");
                          console.log(e);
                      }
                  });
              }
          }
          else if (files.files.length == 0) {
              isAllfield = this.newsValidation()
              let saveData = {
                  Title: $("#txtTitle").val(),
                  Date: $("#txtExpires").val(),
                  Explanation: $("#txtrequiredDescription").val(),

              }
              let cropchecker = this.imagecropperChecking();
              if (isAllfield && cropchecker) {
                  $body.addClass("loading");
                  updateitems("News", strComponentId, saveData, function (e) {

                      if (e.data) {
                          $body.removeClass("loading");
                          that.pageBack();
                          //window.location.href = "" + siteURL + "/Pages/ListView.aspx?CName=" + strLocalStorage + "";

                      } else {
                          $body.removeClass("loading");
                          console.log(e);
                      }
                  });
              }
          } else {
              if (!$('#inputImage').val().trim()) {
                  alertify.set('notifier', 'position', 'top-right');
                  alertify.error("Please Enter Title");
                  isAllfield = false;
              }
              isAllfield = this.newsValidation()
              let cropchecker = this.imagecropperChecking();
              if (isAllfield && cropchecker) {
                  $body.addClass("loading");
                  //var filename = $('#inputImage').val().replace(/C:\\fakepath\\/i, '');
                  var uniquename = Math.random().toString(36).substr(2, 9) + ".png";
                  var file1 = $('#cropped-img').attr('src').split("base64,");
                  var blob = base64ToArrayBuffer(file1[1]);

                  pnp.sp.web.getFolderByServerRelativeUrl("News").files.add(uniquename, blob, true)
                      .then(function (result) {
                          pnp.sp.web.lists.getByTitle("News").items.getById(strComponentId).update({
                              ID: strComponentId,
                              Title: $("#txtTitle").val(),
                              Date: $("#txtExpires").val(),
                              Explanation: $("#txtrequiredDescription").val(),
                              Image: {
                                  "__metadata": {
                                      "type": "SP.FieldUrlValue"
                                  },
                                  Url: siteURL + "/" + strLocalStorage + "/" + uniquename
                              },
                          }).then(r => {
                              $body.removeClass("loading");
                              that.pageBack();
                              //window.location.href = "" + siteURL + "/Pages/ListView.aspx?CName=" + strLocalStorage + "";
                          });

                      });
              }
          }
      } else if (strLocalStorage == "Employee Corner") {

          var files = <HTMLInputElement>document.getElementById("uploadFile");
          let file = files.files[0];
          if (files.files.length == 0) {
              //  isAllfield = this.quickReadsValidation();
              if (!$('#txtTitle').val().trim()) {
                  // $('#txtTitle').focus();
                  alertify.set('notifier', 'position', 'top-right');
                  alertify.error("Please Enter Title");
                  isAllfield = false;
              }

              let saveData = {
                  Title: $("#txtTitle").val()
              }
              if (isAllfield) {
                  $body.addClass("loading");
                  updateitems("Employee Corner", strComponentId, saveData, function (e) {

                      if (e.data) {
                          $body.removeClass("loading");
                          that.pageBack();
                          //window.location.href = "" + siteURL + "/Pages/ListView.aspx?CName=" + strLocalStorage + "";

                      } else {
                          $body.removeClass("loading");
                          console.log(e);
                      }
                  });

              }
          } /*else {
              isAllfield = this.quickReadsValidation();
              if (isAllfield) {
                  var files = <HTMLInputElement>document.getElementById("uploadFile");
                  let file = files.files[0];
                 // var uniquename = Math.random().toString(36).substr(2, 9) + "." + file.name.substring(file.name.lastIndexOf(".") + 1, file.name.length);
                  $body.addClass("loading");
                  pnp.sp.web.getFolderByServerRelativeUrl("Shared Documents").files.add(file.name, file, true)
                      .then(function (result) {
                          pnp.sp.web.lists.getByTitle("Employee Corner").items.getById(strComponentId).update({
                              ID: strComponentId,
                              Title: $("#txtTitle").val(),
                              DocumentFile: {
                                  "__metadata": {
                                      "type": "SP.FieldUrlValue"
                                  },
                                  Url: siteURL + "/" + strLocalStorage + "/" + file.name
                              }

                          }).then(r => {
                              $body.removeClass("loading");
                              console.log("Employee Corner Updated Successfully...!");
                              //window.location.href = "" + siteURL + "/Pages/ListView.aspx?CName=" + strLocalStorage + "";
                          });
                      });
              }
          }*/
          else {
              isAllfield = this.quickReadsValidation();
              if (isAllfield) {
                  var fileURL = window.location.origin;
                  var files = <HTMLInputElement>document.getElementById("uploadFile");
                  let file = files.files[0];
                  // var uniquename = Math.random().toString(36).substr(2, 9) + "." + file.name.substring(file.name.lastIndexOf(".") + 1, file.name.length);
                  $body.addClass("loading");
                  pnp.sp.web.getFolderByServerRelativeUrl("Shared Documents").files.add(file.name, file, true)
                      .then(function (result) {
                          pnp.sp.web.lists.getByTitle("Employee Corner").items.getById(strComponentId).update({
                              ID: strComponentId,
                              Title: $("#txtTitle").val(),
                              DocumentFile: {
                                  "__metadata": {
                                      "type": "SP.FieldUrlValue"
                                  },//result.data.ServerRelativeUrl
                                  Url: fileURL + result.data.ServerRelativeUrl
                              }

                          }).then(r => {
                              $body.removeClass("loading");
                              that.pageBack();
                              //window.location.href = "" + siteURL + "/Pages/ListView.aspx?CName=" + strLocalStorage + "";
                          });
                      });
              }
          }
      } else if (strLocalStorage == "Events") {
          var files = <HTMLInputElement>document.getElementById("inputImage");
          let file = files.files[0];
          if (($('#cropped-img')[0].src == siteURL + '/_catalogs/masterpage/BloomHomepage/images/prof-img.jpg')) {
              isAllfield = this.eventsValidation();
              let saveEvents = {
                  Title: $("#txtTitle").val(),
                  StartDate: $('#txtStartDate').val(),
                  EndDate: new Date($('#txtEndDate').val()),
                  Explanation: $('#txtrequiredDescription').val(),
                  Image: {
                      "__metadata": {
                          "type": "SP.FieldUrlValue"
                      },
                      Url: siteURL + "/_catalogs/masterpage/Bloom/images/logo.png"
                  }
              }
              let isDateChecker = this.EventDateChecker();
              let cropchecker = this.imagecropperChecking();
              if (isAllfield && isDateChecker && cropchecker) {
                  $body.addClass("loading");
                  updateitems("Events", strComponentId, saveEvents, function (e) {
                      if (e.data) {
                          $body.removeClass("loading");
                          window.history.back();
                          //window.location.href = "" + siteURL + "/Pages/ListView.aspx?CName=" + strLocalStorage + "";
                          //localStorage.clear();
                      } else {
                          $body.removeClass("loading");
                          console.log(e);
                      }
                  });
              }
              strcrop = "0";
          }
          else if (strcrop == "1" && files.files.length == 0) {
              isAllfield = this.eventsValidation();
              let saveEvents = {
                  Title: $("#txtTitle").val(),
                  StartDate: $('#txtStartDate').val(),
                  EndDate: new Date($('#txtEndDate').val()),
                  Explanation: $('#txtrequiredDescription').val(),
                  Image: {
                      "__metadata": {
                          "type": "SP.FieldUrlValue"
                      },
                      Url: siteURL + "/_catalogs/masterpage/Bloom/images/logo.png"
                  }

              }
              let isDateChecker = this.EventDateChecker();
              let cropchecker = this.imagecropperChecking();
              if (isAllfield && isDateChecker && cropchecker) {
                  $body.addClass("loading");
                  updateitems("Events", strComponentId, saveEvents, function (e) {
                      if (e.data) {
                          $body.removeClass("loading");
                          window.history.back();
                          //window.location.href = "" + siteURL + "/Pages/ListView.aspx?CName=" + strLocalStorage + "";
                          localStorage.clear();
                      } else {
                          $body.removeClass("loading");
                          console.log(e);
                      }
                  });

              }
              strcrop = "0";
          }

          else if (files.files.length == 0) {
              isAllfield = this.eventsValidation();
              let saveEvents = {
                  Title: $("#txtTitle").val(),
                  StartDate: $('#txtStartDate').val(),
                  EndDate: new Date($('#txtEndDate').val()),
                  Explanation: $('#txtrequiredDescription').val(),

              }
              let isDateChecker = this.EventDateChecker();
              let cropchecker = this.imagecropperChecking();
              if (isAllfield && isDateChecker && cropchecker) {
                  $body.addClass("loading");
                  updateitems("Events", strComponentId, saveEvents, function (e) {

                      if (e.data) {
                          $body.removeClass("loading");
                          that.pageBack();
                          //window.location.href = "" + siteURL + "/Pages/ListView.aspx?CName=" + strLocalStorage + "";

                      } else {
                          $body.removeClass("loading");
                          console.log(e);
                      }
                  });

              }
              strcrop = "0";
              /* if ($("#inputImage").val() == undefined || $("#inputImage").val() == null || $("#inputImage").val() == '') {
                 $.ajax({
                   url: "https://zsl85.sharepoint.com/sites/BloomHolding/_api/web/getfilebyserverrelativeurl('/sites/BloomHolding/_catalogs/masterpage/Bloom/images/logo.png')/openbinarystream",
                   type: "GET",
                   success: function (data) {
                     var name = 'Sample1.jpg'
                     pnp.sp.web.getFolderByServerRelativeUrl("Events").files.add(name, data, true)
                       .then(function (result) {
                         pnp.sp.web.lists.getByTitle("Events").items.getById(strComponentId).update({
                           Title: $("#txtTitle").val(),
                           StartDate: $('#txtStartDate').val(),
                           EndDate: $('#txtEndDate').val(),
                           Explanation: $('#txtDescription').val(),
                           Image: {
                             "__metadata": { "type": "SP.FieldUrlValue" },
                             Url: siteURL + "/_catalogs/masterpage/Bloom/images/logo.png"
                           }
                         }).then(r => {
                           $('#UpdateItem').prop('disabled',true);
                           console.log(name + " properties updated successfully!");
                           //window.location.href = "" + siteURL + "/Pages/ListView.aspx?CName=" + strLocalStorage + "";
                         });
                       })
                   },
                   error: function (data) {
                     console.log(data);
                   },

                 });
               }*/

          } else {
              //  var files = <HTMLInputElement>document.getElementById("inputImage");
              //  let file = files.files[0];
              var fileURL = window.location.origin;
              isAllfield = this.eventsValidation();
              let isDateChecker = this.EventDateChecker();
              let cropchecker = this.imagecropperChecking();
              if (isAllfield && isDateChecker && cropchecker) {
                  $body.addClass("loading");
                  // var filename = $('#inputImage').val().replace(/C:\\fakepath\\/i, '');
                  // var uniquename = Math.random().toString(36).substr(2, 9) + ".png";
                  var file1 = $('#cropped-img').attr('src').split("base64,");
                  var blob = base64ToArrayBuffer(file1[1]);
                  pnp.sp.web.getFolderByServerRelativeUrl("PublishingImages").files.add(file.name, blob, true)
                      .then(function (result) {
                          pnp.sp.web.lists.getByTitle("Events").items.getById(strComponentId).update({
                              ID: strComponentId,
                              Title: $("#txtTitle").val().trim(),
                              StartDate: new Date($('#txtStartDate').val()),
                              EndDate: new Date($('#txtEndDate').val()),
                              Explanation: $('#txtrequiredDescription').val(),
                              Image: {
                                  "__metadata": {
                                      "type": "SP.FieldUrlValue"
                                  },
                                  Url: fileURL + result.data.ServerRelativeUrl
                              }
                          }).then(r => {
                              $body.removeClass("loading");
                              that.pageBack();
                              //window.location.href = "" + siteURL + "/Pages/ListView.aspx?CName=" + strLocalStorage + "";
                          });

                      });
              }
              strcrop = "0";
          }
      } else if (strLocalStorage == "Organizational Policies") {

          var files = <HTMLInputElement>document.getElementById("uploadFile");
          let file = files.files[0];
          if (files.files.length == 0) {
              this.orgpolicyValidation(isAllfield);
              let saveData = {
                  Title: $("#txtTitle").val(),
                  Explanation: $("#txtDescription").val(),
              }
              if (isAllfield) {
                  $body.addClass("loading");
                  updateitems("Organizational Policies", strComponentId, saveData, function (e) {
                      if (e.data) {
                          $body.removeClass("loading");
                          that.pageBack();
                          //window.location.href = "" + siteURL + "/Pages/OrganizationalPolicies.aspx";
                      } else {
                          $body.removeClass("loading");
                          console.log(e);
                      }
                  });
              }
          } else {
              this.orgpolicyValidation(isAllfield);
              if (isAllfield) {
                  var fileURL = window.location.origin;
                  $body.addClass("loading");
                  pnp.sp.web.getFolderByServerRelativeUrl("Shared Documents").files.add(file.name, file, true)
                      .then(function (result) {
                          pnp.sp.web.lists.getByTitle("Organizational Policies").items.getById(strComponentId).update({
                              ID: strComponentId,
                              Title: $("#txtTitle").val(),
                              Explanation: $("#txtDescription").val(),
                              DocumentFile: {
                                  "__metadata": {
                                      "type": "SP.FieldUrlValue"
                                  },
                                  Url: fileURL + result.data.ServerRelativeUrl
                              }

                          }).then(r => {
                              $body.removeClass("loading");
                              that.pageBack();
                              //window.location.href = "" + siteURL + "/Pages/OrganizationalPolicies.aspx";

                          });

                      });
              }
          }

      } else if (strLocalStorage == "Banners") {
          //var regexp = /^(https?|s?ftp):\/\/(((([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(%[\da-f]{2})|[!\$&'\(\)\*\+,;=]|:)*@)?(((\d|[1-9]\d|1\d\d|2[0-4]\d|25[0-5])\.(\d|[1-9]\d|1\d\d|2[0-4]\d|25[0-5])\.(\d|[1-9]\d|1\d\d|2[0-4]\d|25[0-5])\.(\d|[1-9]\d|1\d\d|2[0-4]\d|25[0-5]))|((([a-z]|\d|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(([a-z]|\d|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])*([a-z]|\d|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])))\.)+(([a-z]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(([a-z]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])*([a-z]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])))\.?)(:\d*)?)(\/((([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(%[\da-f]{2})|[!\$&'\(\)\*\+,;=]|:|@)+(\/(([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(%[\da-f]{2})|[!\$&'\(\)\*\+,;=]|:|@)*)*)?)?(\?((([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(%[\da-f]{2})|[!\$&'\(\)\*\+,;=]|:|@)|[\uE000-\uF8FF]|\/|\?)*)?(#((([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(%[\da-f]{2})|[!\$&'\(\)\*\+,;=]|:|@)|\/|\?)*)?$/i
          var files = <HTMLInputElement>document.getElementById("inputImage");
          let file = files.files[0];
          if ($('#cropped-img')[0].src == siteURL + '/_catalogs/masterpage/BloomHomepage/images/prof-img.jpg') {
              alertify.set('notifier', 'position', 'top-right');
              alertify.error("Please Select the Image");
              isAllfield = false;
          }
          else if (files.files.length == 0) {
              if (!$('#txtTitle').val().trim()) {
                  alertify.set('notifier', 'position', 'top-right');
                  alertify.error("Please Enter Title");
                  isAllfield = false;

              } else if (!$("#txtrequiredDescription").val().trim()) {
                  alertify.set('notifier', 'position', 'top-right');
                  alertify.error("Please Enter Title");
                  isAllfield = false;

              }
              let saveData = {
                  Title: $("#txtTitle").val(),
                  BannerContent: $("#txtrequiredDescription").val(),
                  LinkURL: {
                      "__metadata": {
                          "type": "SP.FieldUrlValue"
                      },
                      Url: $('#txtHyper').val().trim(),
                  }
              }
              this.imagecropperChecking();
              if (isAllfield) {
                  $body.addClass("loading");
                  updateitems("Banners", strComponentId, saveData, function (e) {

                      if (e.data) {
                          $body.removeClass("loading");
                          that.pageBack();
                          //window.location.href = "" + siteURL + "/Pages/ListView.aspx?CName=" + strLocalStorage + "";

                      } else {
                          $body.removeClass("loading");
                          console.log(e);
                      }
                  });

              }
          } else {
              if (!$('#inputImage').val()) {
                  alertify.set('notifier', 'position', 'top-right');
                  alertify.error("Please Select Image");
                  isAllfield = false;

              } else if (!$('#txtTitle').val().trim()) {
                  alertify.set('notifier', 'position', 'top-right');
                  alertify.error("Please Enter Title");
                  isAllfield = false;

              } else if (!$("#txtrequiredDescription").val().trim()) {
                  alertify.set('notifier', 'position', 'top-right');
                  alertify.error("Please Enter Description");
                  isAllfield = false;
              }
              var uniquename = Math.random().toString(36).substr(2, 9) + ".png";
              var file1 = $('#cropped-img').attr('src').split("base64,");
              var blob = base64ToArrayBuffer(file1[1]);
              // isAllfield=this.bannersValidation();
              this.imagecropperChecking();
              if (isAllfield) {
                  var fileURL = window.location.origin;
                  $body.addClass("loading");
                  pnp.sp.web.getFolderByServerRelativeUrl("Images").files.add(uniquename, blob, true)
                      .then(function (result) {
                          pnp.sp.web.lists.getByTitle("Banners").items.getById(strComponentId).update({
                              ID: strComponentId,
                              Title: $("#txtTitle").val(),
                              BannerContent: $("#txtDescription").val(),
                              Image: {
                                  "__metadata": {
                                      "type": "SP.FieldUrlValue"
                                  },
                                  Url: fileURL + result.data.ServerRelativeUrl
                                  // Url: siteURL + "/" + strLocalStorage + "/" + uniquename
                              },
                              LinkURL: {
                                  "__metadata": {
                                      "type": "SP.FieldUrlValue"
                                  },
                                  Url: $('#txtHyper').val().trim(),
                              }
                          }).then(r => {
                              $body.removeClass("loading");
                              that.pageBack();
                              //window.location.href = "" + siteURL + "/Pages/ListView.aspx?CName=" + strLocalStorage + "";

                          });

                      });
              }
          }
      } else if (strLocalStorage == "Polls") {
          isAllfield = this.pollsValidation()

          let myobjPols = {
              Question: $("#txtQuestion").val(),
              Options: $("#txtOptions").val()

          }
          if (isAllfield) {
              $body.addClass("loading");
              updateitems("Polls", strComponentId, myobjPols, function (e) {

                  if (e.data) {
                      $body.removeClass("loading");
                      that.pageBack();
                      //window.location.href = "" + siteURL + "/Pages/ListView.aspx?CName=" + strLocalStorage + "";

                  } else {
                      $body.removeClass("loading");
                      console.log(e);
                  }
              });

          }
      } else if (strLocalStorage == "Corporate Discounts") {
          var fileURL = window.location.origin;
          var docfiles = <HTMLInputElement>document.getElementById("uploadFile");
          let docfile = docfiles.files[0];
          var files = <HTMLInputElement>document.getElementById("inputImage");
          let file = files.files[0];
          if (files.files.length == 0 && docfiles.files.length == 0) {
              isAllfield = this.corporationValidation();
              let saveData = {
                  Title: $("#txtTitle").val(),
                  SiteLink: {
                      "__metadata": {
                          "type": "SP.FieldUrlValue"
                      },
                      Url: $("#txtsitelink").val(),
                  }
              }
              if (isAllfield) {
                  $body.addClass("loading");
                  updateitems("Corporate Discounts", strComponentId, saveData, function (e) {

                      if (e.data) {
                          $body.removeClass("loading");
                          that.pageBack();
                          //window.location.href = "" + siteURL + "/Pages/ListView.aspx?CName=" + strLocalStorage + "";

                      } else {
                          $body.removeClass("loading");
                          console.log(e);
                      }
                  });
              }
          } else if (files.files.length > 0 && docfiles.files.length > 0) {
              var fileURL = window.location.origin;
              var file1 = $('#cropped-img').attr('src').split("base64,");
              var blob = base64ToArrayBuffer(file1[1]);
              //var filename = $('#inputImage').val().replace(/C:\\fakepath\\/i, '');
              var uniquename = Math.random().toString(36).substr(2, 9) + ".png";
              pnp.sp.web.getFolderByServerRelativeUrl("PublishingImages").files.add(file.name, blob, true)
                  .then(function (result) {
                      pnp.sp.web.getFolderByServerRelativeUrl("Shared Documents").files.add(docfile.name, docfile, true)
                          .then(function (datafile) {
                              pnp.sp.web.lists.getByTitle("Corporate Discounts").items.getById(strComponentId).update({
                                  ID: strComponentId,
                                  Title: $("#txtTitle").val(),
                                  SiteLink: {
                                      "__metadata": {
                                          "type": "SP.FieldUrlValue"
                                      },
                                      Url: $("#txtsitelink").val(),
                                  },
                                  VendorLogo: {
                                      "__metadata": {
                                          "type": "SP.FieldUrlValue"
                                      },
                                      Url: fileURL + result.data.ServerRelativeUrl
                                  },
                                  DocumentFile: {
                                      "__metadata": {
                                          "type": "SP.FieldUrlValue"
                                      },
                                      Url: fileURL + datafile.data.ServerRelativeUrl
                                  }
                              }).then(r => {
                                  $body.removeClass("loading");
                                  window.history.back();
                                  //window.location.href = "" + siteURL + "/Pages/ListView.aspx?CName=" + strLocalStorage + "";
                              });

                          });
                  });

          } else if (files.files.length > 0) {

              isAllfield = this.corporationValidation();
              if (isAllfield) {
                  $body.addClass("loading");
                  var fileURL = window.location.origin;
                  //var filename = $('#inputImage').val().replace(/C:\\fakepath\\/i, '');
                  var uniquename = Math.random().toString(36).substr(2, 9) + ".png";
                  var file1 = $('#cropped-img').attr('src').split("base64,");
                  var blob = base64ToArrayBuffer(file1[1]);
                  pnp.sp.web.getFolderByServerRelativeUrl("PublishingImages").files.add(uniquename, blob, true)
                      .then(function (result) {
                          pnp.sp.web.lists.getByTitle("Corporate Discounts").items.getById(strComponentId).update({
                              ID: strComponentId,
                              Title: $("#txtTitle").val(),
                              SiteLink: {
                                  "__metadata": {
                                      "type": "SP.FieldUrlValue"
                                  },
                                  Url: $("#txtsitelink").val(),
                              },
                              VendorLogo: {
                                  "__metadata": {
                                      "type": "SP.FieldUrlValue"
                                  },
                                  Url: fileURL + result.data.ServerRelativeUrl
                              }
                          }).then(r => {
                              $body.removeClass("loading");
                              that.pageBack();
                              //window.location.href = "" + siteURL + "/Pages/ListView.aspx?CName=" + strLocalStorage + "";
                          });
                      });
              }
          }
          else {

              isAllfield = this.corporationValidation();
              if (isAllfield) {
                  $body.addClass("loading");
                  var fileURL = window.location.origin;
                  pnp.sp.web.getFolderByServerRelativeUrl("Shared Documents").files.add(docfile.name, docfile, true)
                      .then(function (result) {
                          pnp.sp.web.lists.getByTitle("Corporate Discounts").items.getById(strComponentId).update({
                              ID: strComponentId,
                              Title: $("#txtTitle").val(),
                              SiteLink: {
                                  "__metadata": {
                                      "type": "SP.FieldUrlValue"
                                  },
                                  Url: $("#txtsitelink").val(),
                              },
                              DocumentFile: {
                                  "__metadata": {
                                      "type": "SP.FieldUrlValue"
                                  },
                                  Url: fileURL + result.data.ServerRelativeUrl
                              }
                          }).then(r => {
                              $body.removeClass("loading");
                              that.pageBack();
                              //window.location.href = "" + siteURL + "/Pages/ListView.aspx?CName=" + strLocalStorage + "";
                          });
                      });
              }
          }
      }

  }

  renderhtml(strComponentId) {
      var renderhtml = "<ul>";
      var renderhtmlImage = "";
      var rendercrop = "";
      var rendertext = "";
      var renderdate = "";
      var renderDescription = "";
      var renderEventDate = "";
      var renderHyperlink = "";
      var renderHyperSitelink = "";
      var renderUploadfile = "";
      var renderCorpUploadfile = "";
      var renderRequiredDescription = "";
      var renderUploadOrganization = "";
      var renderSiteLink = "";
      var renderStartEndDate = "";
      var renderhtmlImageEvents = "";
      var renderhtmlImageBanners = "";
      var renderhtmlCorporateImage = "";
      var renderQuestion = "";
      var renderAnswers = "";
      var renderDropdown = "";
      var renderNews = "";
      var sSiteURL = this.context.pageContext.web.absoluteUrl;
      var strLocalStorage = GetQueryStringParams("CName");
      strLocalStorage = strLocalStorage.split('%20').join(' ');
      var renderfileuploadwithlogo;
      var strComponentMode = GetQueryStringParams("CMode");

      renderhtmlImage += "<div class='form-imgsec'>" +
          "<div class='themelogo-upload'>" +
          "<label class='control-label'>Image</label>" +
          "<img class='crapImages' src=''/>" +
          "<div class='image-upload'>" +
          "<div class='custom-upload'>" +
          "<input type='file' id='inputImage' name='file' accept='.jpg,.jpeg,.png' multiple='' class='file' />" +
          "<button class='btn btn-primary' type='button'><i class='icon-camera'></i></button>" +
          "</div>" +
          "<a href='#'id='clearImage' title='Clear' id='image-delete'>delete<i class='icon-delete'></i></a>" +
          "</div>" +
          "</div>";

      rendercrop += "<div class='col-lg-6 col-md-6 col-sm-6 col-xs-12' id='canvasdisplay' style='display:none'>" +
          "<h4>Image Preview </h4>" +
          "<div class='btn-group-crop'>" +
          "<button type='button' class='btn btn-primary'id='btnCrop' ><i class='commonicon-save'></i>Save</button>" +
          "<button class='btn btn-primary crop-cancel' id='btnRestore' type='button'><i class='commonicon-close'></i>Cancel</button>" +

          "<canvas id='canvas'>" +
          "</canvas>" +
          "</div>" +
          "</div>";

      renderhtmlImageEvents += "<div class='form-imgsec'>" +
          "<div class='themelogo-upload'>" +
          "<label>Image</label>" +
          "<img id='cropped-img' class='crapImagesevent crop-imagedisplay' src=''/>" +
          "<div class='image-upload'>" +
          "<div class='custom-upload'>" +
          "<input type='file' id='inputImage' name='file' accept='.jpg,.jpeg,.png,.gif,.bmp,.tiff' multiple='' class='file' />" +
          "<button class='btn btn-primary' type='button'><i class='icon-camera'></i></button>" +
          "</div>" +
          "<a href='#' title='Delete' id='image-delete'>" +
          "<i class='icon-delete'></i>" +
          "</a>" +
          "</div>" +
          "</div>" +
          "</div>";

      renderhtmlImageBanners += "<div class='form-imgsec'>" +
          "<div class='themelogo-upload'>" +
          "<label class='control-label'>Image</label>" +
          "<img id='cropped-img' class='crapImagesevent crop-imagedisplay' src=''/>" +
          "<div class='image-upload'>" +
          "<div class='custom-upload'>" +
          "<input type='file' id='inputImage' name='file' accept='.jpg,.jpeg,.png,.gif,.bmp,.tiff' multiple='' class='file' />" +
          "<button class='btn btn-primary' type='button'><i class='icon-camera'></i></button>" +
          "</div>" +
          "<a href='#' title='Delete' id='image-delete'>" +
          "<i class='icon-delete'></i>" +
          "</a>" +
          "</div>" +
          "</div>" +
          "</div>";

      renderhtmlCorporateImage += "<div class='form-imgsec'>" +
          "<div class='themelogo-upload'>" +
          "<label>Vendor Logo</label>" +
          "<img id='cropped-img' class='crapImagesevent crop-imagedisplay' src=''/>" +
          "<div class='image-upload'>" +
          "<div class='custom-upload'>" +
          "<input type='file' id='inputImage' name='file' accept='.jpg,.jpeg,.png,.gif,.bmp,.tiff' multiple='' class='file' />" +
          "<button class='btn btn-primary' type='button'><i class='icon-camera'></i></button>" +
          "</div>" +
          "<a href='#' title='Delete' id='image-delete'>" +
          "<i class='icon-delete'></i>" +
          "</a>" +
          "</div>" +
          "</div>" +
          "</div>";

      rendertext += "<div id='renderText' class='input text'>" +
          "<label class='control-label'>Title</label>" +
          "<input class='form-control' type='text' value='' id='txtTitle' /></div>";

      renderdate += "<div class='input date'><i class='icon-calenter'></i>" +
          "<label class='control-label'>Date</label>" +
          "<input class='form-control date-selector' type='text' value='' id='txtExpires' /></div>";

      renderDescription += "<div class='input textarea'><label >Description</label><textarea class='form-control' id='txtDescription'></textarea></div>";

      renderRequiredDescription += "<div id='rrdescription' class='input textarea'><label class='control-label'>Description</label><textarea class='form-control' id='txtrequiredDescription'></textarea></div>";

      renderEventDate += "<div class='input date'>" +
          "<i class='icon-calenter'></i>" +
          "<label class='control-label'>Start Date</label>" +
          "<input class='form-control date-selector' type='text' value='' id='txtEvDate' />" +
          "</div>" +
          "<div class='input date'>" +
          "<i class='icon-calenter'></i>" +
          "<label>End Date</label>" +
          "<input class='form-control date-selector' type='text' value='' id='txtEEDate' />" + "</div>";

      renderHyperlink += "<div class='input text'>" +
          "<label class='control-label'>Hyperlink</label>" +
          "<input class='form-control' type='text' value='' id='txtHyper' />" +
          "<span>Please enter the Hyperlink in the following format : https://www.bloomholding.com</span>" +
          "</div>";

      renderHyperSitelink += "<div class='input text'>" +
          "<label>Link URL</label>" +
          "<input class='form-control' type='text' value='' id='txtHyper' />" +
          "<label>Please given valid Announcements or Events URL</label>" +
          "</div>";

      renderUploadfile += "<div class='form-imgsec'>" +
          "<a id='filetype' href='' download><img id='fileimg' src=''></a>" +
          "<div class='themelogo-upload' style='display: block;'>" +
          "<div class='custom-upload banner-upload'>" +
          "<label class='control-label'>Document File</label>" +
          "<input type='file' id='uploadFile' name='file' accept='.doc,.docx,.xls,.ppt,.pdf,.jpg' multiple='' class='file'>" +
          "<div class='input-group'>" +
          "<span class='input-group-btn input-group-sm'>" +
          "<button type='button' class='btn btn-fab btn-fab-mini'>Browse</button>" +
          "</span>" +
          "<input type='text' class='form-control' placeholder='Upload Files'>" +
          "</div>" +
          "</div>" +
          "</div>" +
          "</div>";

      renderCorpUploadfile += "<div class='form-imgsec'>" +
          "<a id='filetype' href='' download><img id='fileimg' src=''></a>" +
          "<div class='themelogo-upload' style='display: block;'>" +
          "<div class='custom-upload banner-upload'>" +
          "<label>Document File</label>" +
          "<input type='file' id='uploadFile' name='file' accept='.doc,.docx,.xls,.ppt,.pdf,.jpg' multiple='' class='file'>" +
          "<div class='input-group'>" +
          "<span class='input-group-btn input-group-sm'>" +
          "<button type='button' class='btn btn-fab btn-fab-mini'>Browse</button>" +
          "</span>" +
          "<input type='text' class='form-control' placeholder='Upload Files'>" +
          "</div>" +
          "</div>" +
          "</div>" +
          "</div>";
      renderUploadOrganization += "<div class='form-imgsec'>" +
          "<a id='filetype' href='' download><img id='fileimg' src=''></a>" +
          "<div class='themelogo-upload' style='display: block;'>" +
          "<div class='custom-upload banner-upload'>" +
          "<input type='file' id='inputImage' name='file' accept='.pdf,.doc,.docx' multiple='' class='file'>" +
          "<div class='input-group'>" +
          "<span class='input-group-btn input-group-sm'>" +
          "<button type='button' class='btn btn-fab btn-fab-mini'>Browse</button>" +
          "</span>" +
          "<input type='text' class='form-control' placeholder='Upload Files'>" +
          "</div>" +
          "</div>" +
          "</div>" +
          "</div>";

      renderNews += "<div class='form-imgsec'>" +
          "<div class='themelogo-upload'>" +
          "<label class='control-label'>Image</label>" +
          "<img id='cropped-img' src='' />" +
          "<div class='image-upload'>" +
          "<div class='custom-upload'>" +
          "<input type='file' id='inputImage' name='file' accept='.jpg,.jpeg,.png,.gif,.bmp,.tiff' multiple='' class='file' />" +
          "<button class='btn btn-primary' type='button'>Browse</button>" +
          "</div>" +
          "<a href='#' title='Delete' id='image-delete'>" +
          "<i class='icon-delete'></i>" +
          "</a>" +
          "</div>" +
          "</div>" +
          "</div>";
      renderSiteLink += "<div id='siteLink' class='input text'>" +
          "<i class=''></i>" +
          "<label class='control-label'>Site Link</label>" +
          "<input class='form-control' type='text' value='' id='txtsitelink'/>" +
          "<span>Please enter the Site Link in the following format : https://www.bloomholding.com</span>" +
          "</div>";
      renderStartEndDate += "<div class='input date'>" +
          "<i class='icon-calenter'></i>" +
          "<label class='control-label'>Start Date</label>" +
          "<input class='form-control date-selector' type='text' value='' id='txtStartDate' />" +
          "</div>" +
          "<div class='input date'>" +
          "<i class='icon-calenter'></i>" +
          "<label>End Date</label>" +
          "<input class='form-control date-selector' type='text' value='' id='txtEndDate' />" + "</div>";

      renderQuestion += "<div class='input textarea'>" +
          "<label class='control-label'>Question</label>" +
          "<textarea class='form-control' id='txtQuestion'></textarea>" +
          "</div>";
      renderAnswers += "<div class='input text'>" +
          "<i class=''></i>" +
          "<label class='control-label'>Answers</label>" +
          "<input class='form-control' type='text' value='' id='txtOptions'/>" +
          "<span>Please Enter more than one answers with Semicolon ( ; ) Maximum Four answers</span>" +
          "</div>";
      renderfileuploadwithlogo += "<div id='filewithLogo'></div>"



      this.getListItems(strComponentId);
      $('.appendsec').append(renderhtml);
      console.log(strLocalStorage);
      var date = new Date();
      var today = new Date(date.getFullYear(), date.getMonth(), date.getDate());

      if (strLocalStorage == 'Announcements') {

          $('.form-section').append(renderhtmlImageBanners);
          $('.form-imgsec').after(rendercrop);
          $('#canvasdisplay').after(renderdate);
          $('.date').after(rendertext);
          $('.text').after(renderRequiredDescription);
          $('#txtExpires').datepicker({
              format: "mm/dd/yyyy"

          });
          /* $('#txtExpires').datepicker({ dateFormat: 'yy-mm-dd' }).bind("change",function(){
      var minValue = $(this).val();
      minValue = $('#txtExpires').datepicker.parseDate("yy-mm-dd", minValue);
      minValue.setDate(minValue.getDate());
      $('#txtExpires').datepicker( "option", "minDate", minValue );
  })*/
          $(document).on('change', '.file', function () {
              $(this).parent('.custom-upload').find('.form-control').val($(this).val().replace(/C:\\fakepath\\/i, ''));
          });
          this.ViewMode(strComponentMode);
      } else if (strLocalStorage == 'Holiday') {
          $('.form-section').append(rendertext);
          $('.text').after(renderEventDate);
          $('#txtEvDate').datepicker({
              format: "mm/dd/yyyy",
              //startDate: today
          });
          $('#txtEEDate').datepicker({
              format: "mm/dd/yyyy",
              //startDate: today
          });
          this.ViewMode(strComponentMode);
      } else if (strLocalStorage == 'News') {
          $('.form-section').append(renderhtmlImageEvents);
          $('.form-imgsec').after(rendercrop);
          $('#canvasdisplay').after(renderdate);
          $('.date').after(rendertext);
          $('.text').after(renderRequiredDescription);
          $('#txtExpires').datepicker({
              format: "mm/dd/yyyy"
              // startDate: today
          });
          this.ViewMode(strComponentMode);
      } else if (strLocalStorage == 'Quick Links') {
          $('.form-section').append(rendertext);
          $('.text').after(renderHyperlink);
          this.ViewMode(strComponentMode);
      } else if (strLocalStorage == 'Employee Corner') {
          $('.form-section').append(rendertext);
          $('.text').after(renderUploadfile);
          //  $('.form-imgsec').after(renderHyperlink);
          $(document).on('change', '.file', function () {
              if ($.inArray($(this).val().split('.').pop().toLowerCase(), ['doc', 'docx', 'xls', 'csv', 'ppt', 'pdf']) == -1) {
                  alertify.set('notifier', 'position', 'top-right');
                  alertify.error("Please Select Valid file Format");
                  $("#uploadFile").val("");
              } else {
                  $(this).parent('.custom-upload').find('.form-control').val($(this).val().replace(/C:\\fakepath\\/i, ''));
              }
          });
          this.ViewMode(strComponentMode);
      } else if (strLocalStorage == 'Organizational Policies') {
          $('.form-section').append(rendertext);
          $('.text').after(renderUploadfile);
          $('.form-imgsec').after(renderDescription);
          $(document).on('change', '.file', function () {
              if ($.inArray($(this).val().split('.').pop().toLowerCase(), ['doc', 'docx', 'xls', 'csv', 'ppt', 'pdf']) == -1) {
                  alertify.set('notifier', 'position', 'top-right');
                  alertify.error("Please Select Valid file Format");
                  $("#uploadFile").val("");
              } else {
                  $(this).parent('.custom-upload').find('.form-control').val($(this).val().replace(/C:\\fakepath\\/i, ''));
              }
          });
          this.ViewMode(strComponentMode);
      } else if (strLocalStorage == 'Banners') {

          $('.form-section').append(renderhtmlImageBanners);
          $('.form-imgsec').after(rendercrop);
          $('#canvasdisplay').after(rendertext);
          $('.text').after(renderRequiredDescription);
          $('#rrdescription').after(renderHyperSitelink);
          //$('.textarea').after(renderDropdown);
          this.ViewMode(strComponentMode);

      } else if (strLocalStorage == 'Corporate Discounts') {
          $('.form-section').append(renderhtmlCorporateImage);
          $('.form-imgsec').after(rendercrop);
          $('#canvasdisplay').after(rendertext);
          $('.text').after(renderSiteLink);
          $('#siteLink').after(renderCorpUploadfile);

          this.ViewMode(strComponentMode);

      } else if (strLocalStorage == 'Events') {

          $('.form-section').append(renderhtmlImageEvents);
          $('.form-imgsec').after(rendercrop);
          $('#canvasdisplay').after(rendertext);
          $('#renderText').after(renderRequiredDescription);
          $('.textarea').after(renderStartEndDate);
          $('#txtStartDate').datepicker({
              format: "mm/dd/yyyy",
              //  startDate: today
          });
          $('#txtEndDate').datepicker({
              format: "mm/dd/yyyy",
              //startDate: today
          });
          this.ViewMode(strComponentMode);

      } else if (strLocalStorage == 'Polls') {
          $('.form-section').append(renderQuestion);
          $('.textarea').after(renderAnswers);
          this.ViewMode(strComponentMode);
      }
      $('.date-selector').on('changeDate', function (ev) {
          $(this).datepicker('hide');
      });
      $("#txtStartDate").keypress(function (evt) {

          var keycode = evt.charCode || evt.keyCode;
          if (keycode == 13) { //Enter key's keycode
              return false;
          } else {
              evt.preventDefault();
          }
      });
      $("#txtEndDate").keypress(function (evt) {

          var keycode = evt.charCode || evt.keyCode;
          if (keycode == 13) { //Enter key's keycode
              return false;
          } else {
              evt.preventDefault();
          }
      });
      if ($('#uploadFile').length > 0) {
          $(document).on('change', '#uploadFile', function () {
              var docname = $(this).val().split('.');
              docname = docname[docname.length - 1].toLowerCase();
              if ($.inArray(docname, ['doc', 'docx', 'xls', 'csv', 'ppt', 'pdf']) == -1) {
                  alertify.set('notifier', 'position', 'top-right');
                  // var msg=alertify.error("Please Select Valid File Format");

                  alertify.error("Please Select Valid File Format");
                  $("#uploadFile").val("");
              } else {
                  $(this).parent('.custom-upload').find('.form-control').val($(this).val().replace(/C:\\fakepath\\/i, ''));
              }
          });
      }
      $('#image-delete').click(function () {
          if ($('#cropped-img')[0].src == sSiteURL + '/_catalogs/masterpage/BloomHomepage/images/prof-img.jpg') {
              alertify.set('notifier', 'position', 'top-right');
              alertify.error("Please upload the Image File");
          }
          else {
              $('#cropped-img').removeClass("crop-imagedisplay");
              $('.image-upload').css('width', '103px');
              $("#cropped-img").attr('src', sSiteURL + '/_catalogs/masterpage/BloomHomepage/images/prof-img.jpg');
              $('#inputImage').val("");
          }


      });
      if ($('#inputImage').length > 0) {
          
          var canvas = $("#canvas"),
              context = canvas.get(0).getContext("2d"),
              $result = $('#cropped-img');

          $('#inputImage').on('change', function () {
              var iscropflag = true;
              var docname = $(this).val().split('.');
              docname = docname[docname.length - 1].toLowerCase();
              if ($.inArray(docname, ['jpg', 'jpeg', 'png', 'gif', 'bmp', 'tiff']) == -1) {
                  alertify.set('notifier', 'position', 'top-right');
                  alertify.error("Please Select Valid file Format");
                  $("#inputImage").val("");
                  iscropflag = false
              }
              if (iscropflag) {
                  canvas.cropper('destroy');
                  if (this.files && this.files[0]) {
                      if (this.files[0].type.match(/^image\//)) {
                          var reader = new FileReader();
                          reader.onload = function (evt) {
                              var img = new Image();
                              img.onload = function () {
                                  context.canvas.height = img.height;
                                  context.canvas.width = img.width;
                                  context.drawImage(img, 0, 0);
                                  var cropper = canvas.cropper({
                                      aspectRatio: 16 / 9
                                  });

                              };
                              //img.src = evt.target.result;
                              img.src = evt.target['result'];
                              $('#canvasdisplay').css('display', 'block');
                          };
                          reader.readAsDataURL(this.files[0]);
                      } else {
                          // alert("Invalid file type! Please select an image file.");
                      }
                  } else {
                      //alert('No file(s) selected.');
                  }
              }
          });
          $('#btnCrop').click(function () {
              // Get a string base 64 data url
              $result.empty();
              var croppedImageDataURL = canvas.cropper('getCroppedCanvas').toDataURL("image/png");
              $result.attr('class', 'crop-imagedisplay');
              //  $('.image-upload').css('width', '42%');
              $result.attr('src', croppedImageDataURL);
              $('#canvasdisplay').css('display', 'none');
              canvas.cropper('destroy');
          });

          $('#btnRestore').click(function () {
            //  var siteURL = this.context.pageContext.web.absoluteUrl;
              canvas.cropper('reset');
              $result.empty();
              $result.attr('src', siteURL + '/_catalogs/masterpage/BloomHomepage/images/prof-img.jpg');
              $('#canvasdisplay').css('display', 'none');
              $('#inputImage').val("");
          });
      }

  }

  public ViewMode(strComponentMode) {
      if (strComponentMode == 'ViewMode') {
          $('#UpdateItem').hide();
          $('.image-upload').hide();
          $('.form-section :input').prop("disabled", true);
      }

  }
  public getListItems(strComponentId) {
      var count = 5;
      var strLocalStorage = GetQueryStringParams("CName");
      strLocalStorage = strLocalStorage.split('%20').join(' ');
      let objResults;


      if (strLocalStorage == "Announcements") {
          objResults = readItems("Announcements", ["Title", "Explanation", "Expires", "Image", "Display"], count, "Modified", "ID", strComponentId)
          objResults.then((items: any) => {
              $('.crapImagesevent').attr("src", items[0].Image.Url);
              $('#txtTitle').val(items[0].Title);
              var eedate = "";
              if ((items[0].Expires) != null) {
                  eedate = this.nullDateValidate(items[0].Expires);
              }
              // 

              $('#txtExpires').datepicker('setDate', new Date(eedate));
              // $('#txtExpires').val(eedate);
              var description = items[0].Explanation.replace(/<[^>]*>/g, '');
              var splitdesc = description.split('&#160;').join(' ')
              $('#txtrequiredDescription').val(splitdesc)

          })
      } else if (strLocalStorage == "Holiday") {
          objResults = readItems("Holiday", ["Title", "Modified", "EventDate", "EndEventDate", "Display"], count, "Modified", "ID", strComponentId)
          objResults.then((items: any[]) => {
              $('#txtTitle').val(items[0].Title);
              var eedate = "";
              if ((items[0].EventDate) != null) {
                  eedate = this.nullDateValidate(items[0].EventDate);
              }
              // $('#txtEvDate').val(eedate);
              $('#txtEvDate').datepicker('setDate', new Date(eedate));

              if ((items[0].EndEventDate) != null) {
                  eedate = this.nullDateValidate(items[0].EndEventDate);
              }
              //  $('#txtEEDate').val(eedate);
              $('#txtEEDate').datepicker('setDate', new Date(eedate));

          })
      } else if (strLocalStorage == "News") {
          objResults = readItems("News", ["Title", "Modified", "Date", "Image", "Explanation", "Display",], count, "Modified", "ID", strComponentId)
          objResults.then((items: any[]) => {
              $('.crapImagesevent').attr("src", items[0].Image.Url);
              //  $('#txtHyper').val(items[0].Hyperlink.Url)
              $('#txtTitle').val(items[0].Title);
              $('#txtrequiredDescription').val(items[0].Explanation);
              var eedate = "";
              if ((items[0].Date) != null) {
                  eedate = this.nullDateValidate(items[0].Date);
              }
              $('#txtExpires').datepicker('setDate', new Date(eedate));

          })
      } else if (strLocalStorage == "Quick Links") {
          objResults = readItems("Quick Links", ["Title", "Modified", "LinkURL", "Display"], count, "Modified", "ID", strComponentId)
          objResults.then((items: any[]) => {
              $('#txtTitle').val(items[0].Title);
              $('#txtHyper').val(items[0].LinkURL.Url)
          })
      } else if (strLocalStorage == "Employee Corner") {
          debugger;
          objResults = readItems("Employee Corner", ["Title", "Modified", "Icon", "DocumentFile", "Display"], count, "Modified", "ID", strComponentId)
          objResults.then((items: any[]) => {
              $('#txtTitle').val(items[0].Title);
              var strfileType = items[0].DocumentFile.Url.substring(items[0].DocumentFile.Url.lastIndexOf(".") + 1);

              var ftype = this.siteURL + "/_catalogs/masterpage/BloomHomepage/images/";
              if (strfileType == "xls" || strfileType == "xlsx" || strfileType == "csv") {
                  $('#fileimg').attr("src", ftype + "xls.png");
              }
              else if (strfileType == "pdf") {
                  $('#fileimg').attr("src", ftype + "pdf.png");
              }
              else if (strfileType == "doc" || strfileType == "docx") {
                  $('#fileimg').attr("src", ftype + "doc.png");
              }
              else if (strfileType == "ppt") {
                  $('#fileimg').attr("src", ftype + "ppt.png");
              }
              $('#filetype').attr("href", items[0].DocumentFile.Url);

          })
      } else if (strLocalStorage == "Organizational Policies") {
          objResults = readItems("Organizational Policies", ["Title", "Modified", "DocumentFile", "Explanation",], count, "Modified", "ID", strComponentId)
          objResults.then((items: any[]) => {
              $("#txtTitle").val(items[0].Title);
              var strfileType = items[0].DocumentFile.Url.substring(items[0].DocumentFile.Url.lastIndexOf(".") + 1);

              var ftype = this.siteURL + "/_catalogs/masterpage/BloomHomepage/images/";
              if (strfileType == "xls" || strfileType == "xlsx" || strfileType == "csv") {
                  $('#fileimg').attr("src", ftype + "xls.png");
              }
              else if (strfileType == "pdf") {
                  $('#fileimg').attr("src", ftype + "pdf.png");
              }
              else if (strfileType == "doc" || strfileType == "docx") {
                  $('#fileimg').attr("src", ftype + "doc.png");
              }
              else if (strfileType == "ppt") {
                  $('#fileimg').attr("src", ftype + "ppt.png");
              }
              $('#filetype').attr("href", items[0].DocumentFile.Url);

              var description = items[0].Explanation.replace(/<[^>]*>/g, '');
              var splitdesc = description.split('&#160;').join(' ');
              $('#txtDescription').val(splitdesc);
              //$('#txtUrl').val(items[0].DocumentFile.Url);
          })
      } else if (strLocalStorage == "Banners") {
          objResults = readItems("Banners", ["Title", "Modified", "BannerContent", "Display", "LinkURL", "Order", "Image"], count, "Modified", "ID", strComponentId)
          objResults.then((items: any[]) => {
              $('.crapImagesevent').attr("src", items[0].Image.Url);
              $('#txtTitle').val(items[0].Title);
              var bannerContent = items[0].BannerContent.replace(/<[^>]*>/g, '');
              var splitContent = bannerContent.split('&#160;').join(' ');
              $('#txtrequiredDescription').val(splitContent);
              if (items[0].LinkURL == null) {
                  $('#txtHyper').val('');
              } else {
                  $('#txtHyper').val(items[0].LinkURL.Url)
              }
          })
      } else if (strLocalStorage == "Corporate Discounts") {
          objResults = readItems("Corporate Discounts", ["Title", "Modified", "VendorLogo", "SiteLink", "DocumentFile"], count, "Modified", "ID", strComponentId)
          objResults.then((items: any[]) => {
              $('.crapImagesevent').attr("src", items[0].VendorLogo.Url);

              if (items[0].DocumentFile == null) {
                  $('#filetype').hide();
              } else {
                  var strfileType = items[0].DocumentFile.Url.substring(items[0].DocumentFile.Url.lastIndexOf(".") + 1);

                  var ftype = this.siteURL + "/_catalogs/masterpage/BloomHomepage/images/";
                  if (strfileType == "xls" || strfileType == "xlsx" || strfileType == "csv") {
                      $('#fileimg').attr("src", ftype + "xls.png");
                  }
                  else if (strfileType == "pdf") {
                      $('#fileimg').attr("src", ftype + "pdf.png");
                  }
                  else if (strfileType == "doc" || strfileType == "docx") {
                      $('#fileimg').attr("src", ftype + "doc.png");
                  }
                  else if (strfileType == "ppt") {
                      $('#fileimg').attr("src", ftype + "ppt.png");
                  }
                  $('#filetype').attr("href", items[0].DocumentFile.Url);
              }

              $('#txtTitle').val(items[0].Title);
              $('#txtsitelink').val(items[0].SiteLink.Url);
          })
      } else if (strLocalStorage == "Events") {
          objResults = readItems("Events", ["Title", "Modified", "StartDate", "EndDate", "Image", "Explanation"], count, "Modified", "ID", strComponentId)
          objResults.then((items: any[]) => {
              $('.crapImagesevent').attr("src", items[0].Image.Url);
              $('#txtTitle').val(items[0].Title);
              var description = items[0].Explanation.replace(/<[^>]*>/g, '');
              var splitdesc = description.split('&#160;').join(' ');
              $('#txtrequiredDescription').val(splitdesc);
              // $('#txtsitelink').val(items[0].Image.Url);
              var sdate = "";
              var eedate = "";
              if ((items[0].StartDate) != null) {
                  sdate = this.nullDateValidate(items[0].StartDate);
              }

              $('#txtStartDate').datepicker('setDate', new Date(sdate));

              if ((items[0].EndDate) != null) {
                  eedate = this.nullDateValidate(items[0].EndDate);
              }
              //$('#txtEndDate').val(eedate);
              $('#txtEndDate').datepicker('setDate', new Date(eedate));

          })
      } else if (strLocalStorage == "Polls") {
          objResults = readItems("Polls", ["Title", "Modified", "Question", "Options"], count, "Modified", "ID", strComponentId)
          objResults.then((items: any[]) => {
              //$('.crapImages').attr("src", items[0].VendorLogo.Url);
              $('#txtQuestion').val(items[0].Question);
              $('#txtOptions').val(items[0].Options);
          })
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
