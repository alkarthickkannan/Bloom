import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';
import pnp from 'sp-pnp-js';
import { sp } from "sp-pnp-js";
import styles from './AddItemWebPart.module.scss';
import * as strings from 'AddItemWebPartStrings';
import { readItems,addItems,GetQueryStringParams,additemsimage,base64ToArrayBuffer} from '../../commonJS';

export interface IAddItemWebPartProps {
  description: string;
}

import 'jquery';
import '../../ExternalRef/css/bootstrap-datepicker.min.css';
import '../../ExternalRef/css/alertify.min.css';
import '../../ExternalRef/css/cropper.min.css';
require('bootstrap');
require('../../ExternalRef/js/alertify.min.js');
require('../../ExternalRef/js/bootstrap-datepicker.min.js');
require('../../ExternalRef/js/cropper-main.js');
require('../../ExternalRef/js/cropper.min.js');

declare var $;
declare var alertify: any;


export default class AddItemWebPart extends BaseClientSideWebPart<IAddItemWebPartProps> {

  public render(): void {
    
    var siteweburl = this.context.pageContext.site.absoluteUrl;
    var strLocalStorage = GetQueryStringParams("CName");
    strLocalStorage = strLocalStorage.split("%20").join(' ');
    this.domElement.innerHTML =

      "<div class='breadcrumb'>" +
      "<ol>" +
      "<li><a href='" + siteweburl + "/Pages/Home.aspx' title='Home'>Home</a></li>" +
      "<li><a class='pointer' id='breadTilte' title='" + strLocalStorage + "'>" + strLocalStorage + "</a></li>" +
      "<li><span>Add " + strLocalStorage + "</span></li>" +
      "</ol>" +
      "</div>" +
      "<div class='title-section'>" +
      "<div class='button-field save-button'>" +
      "<a  title='Save' class='addbutton pointer' id='AddItem'><i class='commonicon-save addbutton'></i>Save</a>" +
      "<a class='delete-icon close-icon pointer deletebutton' class='closebutton' title='Close' id='DelItem'><i class='commonicon-close deletebutton'></i>Close</a>" +
      "</div>" +
      "<h2 id='ComponentName'>Announcements</h2>" +
      "</div>" +
      "<div  class='form-section required'>" +
      "</div>" +
      "<div class='modal'><!-- Place at bottom of page --></div>";

    document.title = 'Add' + strLocalStorage;
    //For Design Load
    //var _this = this;
    document.getElementById("ComponentName").innerHTML = GetQueryStringParams("CName").split('%20').join(" ");
    this.AddListItems();
    //For Radio Button
    $("input[name='selectionradio']").click(function () {
      var radioValue = $("input[name='selectionradio']:checked").val();
      if (radioValue == "Holiday") {
        $('.themelogo-upload').hide();
        $('#cropped-img').attr("src", '../../../_catalogs/masterpage/BloomHomepage/images/prof-img.jpg');
        $('#rrdescription').hide();
        $("#txtTitle,#txtEndDate,#txtrequiredDescription,#inputImage").val("");
        $('#txtStartDate').datepicker({
          Format: 'yy/dd/mm'

        }).datepicker('setDate', 'new Date()');

      } else {
        $('.themelogo-upload').show();
        $('#rrdescription').show();
      }

    });




    $("input[name='selectionradioImage']").click(function () {
      var radioValue = $("input[name='selectionradioImage']:checked").val();
      if (radioValue == "Upload") {
        $('#divHyperLink').hide();
        $('.banner-upload').show();

      } else {
        $('.banner-upload').hide();
        $('#divHyperLink').show();
      }

    });




    if (strLocalStorage == "Holiday") {
      $("#radio-2").attr('checked', 'checked');
      $('.themelogo-upload').hide();
      $('#cropped-img').attr("src", '../../../_catalogs/masterpage/BloomHomepage/images/prof-img.jpg');
      $('#rrdescription').hide();
      $("#txtTitle,#txtEndDate,#txtrequiredDescription,#inputImage").val("");
      $('#txtStartDate').datepicker({
        Format: 'yy/dd/mm'

      }).datepicker('setDate', 'new Date()');
    } else {
      $("#radio-1").attr('checked', 'checked');
      $('.themelogo-upload').show();
      $('#rrdescription').show();
    }


    $("#radio-3").attr('checked', 'checked');
    //For Save
    let Addevent = document.getElementById('AddItem');
    //for (let i = 0; i < Addevent.length; i++) {
    Addevent.addEventListener("click", (e: Event) => this.AddItem(siteweburl, e));
    //}
    let breadTilte = document.getElementById('breadTilte');
    //for (let i = 0; i < Addevent.length; i++) {
    breadTilte.addEventListener("click", (e: Event) => this.pageBack());

    let Closeevent = document.getElementById('DelItem');
    // Addevent.addEventListener("click", (e: Event) => this.UpdateItem(siteweburl, strComponentId));
    //for (let i = 0; i < Closeevent.length; i++) {
    Closeevent.addEventListener("click", (e: Event) => this.pageBack());
    // }

    this.datepickerkeyTypeBlocker();
    //$("#txtExpires").datepicker("setDate", '12/12/2018');
  }

  datepickerkeyTypeBlocker() {
    $("#txtExpires,#txtStartDate,#txtEndDate,#txtEvDate,#txtEEDate").keypress(
      function (event) {
        event.preventDefault();
      });
  }
  pageBack() {
    window.history.back();
  }

  async AddDepartments() {
    let myobjPols = {
      Departments: $("#txtDepartment").val().trim()
    }
    await addItems("Departments", myobjPols);
  }

 async AddItem(siteweburl, e) {
    var $body = $('body');
    e.preventDefault();
    $(this).prop('disabled', true);
    var dropOptionValue = $("input[name='selectionradio']:checked").val()
    var dropOptionImageValue = $("input[name='selectionradio']:checked").val()
    var $body = $("body");
    var strLocalStorage = GetQueryStringParams("CName");
    strLocalStorage = strLocalStorage.split("%20").join(' ');
    var isAllfield = true;
    // var siteweburl = this.context.pageContext.web.serverRelativeUrl;

    //Add Announcements Part
    if (strLocalStorage == "Announcements") {
      var files = <HTMLInputElement>document.getElementById("inputImage");
      let file = files.files[0];

      if (!$('#txtExpires').val().trim()) {
        alertify.set('notifier', 'position', 'top-right');
        alertify.error("Please Select the Date");
        isAllfield = false;
      } else if (!$('#txtTitle').val().trim()) {
        alertify.set('notifier', 'position', 'top-right');
        alertify.error("Please Enter Title");
        isAllfield = false;
      } else if (!$('#txtrequiredDescription').val().trim()) {
        alertify.set('notifier', 'position', 'top-right');
        alertify.error("Please Enter Description");
        isAllfield = false;
      }
      else if ($('#canvasdisplay').css('display') == 'block') {
        alertify.set('notifier', 'position', 'top-right');
        alertify.error("Please Save Cropped Image");
        isAllfield = false;

      }
      if (isAllfield) {
        $body.addClass("loading");
        if ($("#inputImage").val() == undefined || $("#inputImage").val() == null || $("#inputImage").val() == '') {
          $.ajax({
            url: siteweburl + "/_api/web/getfilebyserverrelativeurl('" + siteweburl + "/_catalogs/masterpage/Bloom/images/logo.png')/openbinarystream",
            type: "GET",
            success: function (data) {
              var name = Math.random().toString(36).substr(2, 9) + ".png";
              pnp.sp.web.getFolderByServerRelativeUrl("Announcements").files.add(name, data, true)
                .then(function (result) {
                  result.file.listItemAllFields.get().then((listItemAllFields) => {
                    pnp.sp.web.lists.getByTitle("Announcements").items.getById(listItemAllFields.Id).update({
                      Title: $("#txtTitle").val().trim(),
                      Expires: new Date($('#txtExpires').val()),
                      Explanation: $("#txtrequiredDescription").val(),
                      Image: {
                        "__metadata": {
                          "type": "SP.FieldUrlValue"
                        },
                        Url: siteweburl + "/_catalogs/masterpage/Bloom/images/logo.png"
                      }
                    }).then(r => {
                      $body.removeClass("loading");
                      $('.addbutton').prop('disabled', true);
                      //this.pageBack();
                      window.history.back();
                      //window.location.href = "" + siteweburl + "/Pages/ListView.aspx?CName=" + strLocalStorage + "";
                      console.log(name + " properties updated successfully!");

                    });
                  });
                })
            },
            error: function (data) {
              $body.removeClass("loading");
              console.log(data);
            },
          });
        } else {
          var uniquename = Math.random().toString(36).substr(2, 9) + ".png";
          var file1 = $('#cropped-img').attr('src').split("base64,");
          var blob = base64ToArrayBuffer(file1[1]);
          let myobjQl = {
            Title: $('#txtTitle').val().trim(),
            Expires: new Date($('#txtExpires').val()),
            Explanation: $("#txtrequiredDescription").val(),
            Image: {
              "__metadata": {
                "type": "SP.FieldUrlValue"
              },
              Url: siteweburl + "/" + strLocalStorage + "/" + uniquename
            }
          }

           await additemsimage(strLocalStorage, uniquename, blob, myobjQl);
          
            $body.addClass("loading");
            if (e.data) {
              $body.removeClass("loading");
              $('.addbutton').prop('disabled', true);
              // this.pageBack();
              window.history.back();
              //   window.location.href = "" + siteweburl + "/Pages/ListView.aspx?CName=" + strLocalStorage + "";

            } else {
              $body.removeClass("loading");
              console.log(e);
            }
          
        }
      }
    }   //Add Holiday Part
    else if (dropOptionValue == "Holiday") {
      if (!$('#txtTitle').val().trim()) {
        alertify.set('notifier', 'position', 'top-right');
        alertify.error("Please Enter Title");
        isAllfield = false;
      } else if (!$('#txtStartDate').val().trim()) {
        //$('#txtEvDate').focus();
        alertify.set('notifier', 'position', 'top-right');
        alertify.error("Please Enter txtStartDate");
        isAllfield = false;
      } else if (Date.parse($("#txtStartDate").val().trim()) > Date.parse($("#txtEndDate").val().trim())) {
        alertify.set('notifier', 'position', 'top-right');
        alertify.error("Start date must be less than End date");
        return false;
      }
      let myobjHol = {
        Title: $("#txtTitle").val().trim(),
        EventDate: new Date($("#txtStartDate").val()),
        EndEventDate: new Date($("#txtEndDate").val()),

      }
      if (isAllfield) {
        $body.addClass("loading");
       await addItems("Holiday", myobjHol);
          $body.addClass("loading");
          if (e.data) {
            $body.removeClass("loading");
            $('.addbutton').prop('disabled', true);
            // this.pageBack();
            window.history.back();
            // window.location.href = "" + siteweburl + "/Pages/ListView.aspx?CName=" + strLocalStorage + "";

          } else {
            $body.removeClass("loading");
            console.log(e);
          }

      }
    }     //Add QuickLinks Part
    else if (strLocalStorage == "Quick Links") {
      var regexp = /^(https?|s?ftp):\/\/(((([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(%[\da-f]{2})|[!\$&'\(\)\*\+,;=]|:)*@)?(((\d|[1-9]\d|1\d\d|2[0-4]\d|25[0-5])\.(\d|[1-9]\d|1\d\d|2[0-4]\d|25[0-5])\.(\d|[1-9]\d|1\d\d|2[0-4]\d|25[0-5])\.(\d|[1-9]\d|1\d\d|2[0-4]\d|25[0-5]))|((([a-z]|\d|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(([a-z]|\d|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])*([a-z]|\d|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])))\.)+(([a-z]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(([a-z]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])*([a-z]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])))\.?)(:\d*)?)(\/((([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(%[\da-f]{2})|[!\$&'\(\)\*\+,;=]|:|@)+(\/(([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(%[\da-f]{2})|[!\$&'\(\)\*\+,;=]|:|@)*)*)?)?(\?((([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(%[\da-f]{2})|[!\$&'\(\)\*\+,;=]|:|@)|[\uE000-\uF8FF]|\/|\?)*)?(#((([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(%[\da-f]{2})|[!\$&'\(\)\*\+,;=]|:|@)|\/|\?)*)?$/i

      if (!$('#txtTitle').val().trim()) {
        //$('#txtTitle').focus();
        alertify.set('notifier', 'position', 'top-right');
        alertify.error("Please Enter Title");
        isAllfield = false;
      } else if (!$('#txtHyper').val().trim()) {
        //$('#txtHyper').focus();
        alertify.set('notifier', 'position', 'top-right');
        alertify.error("Please Enter Link URL");
        isAllfield = false;
      } else if (!regexp.test($('#txtHyper').val().trim())) {
        //$('#txtHyper').focus();
        alertify.set('notifier', 'position', 'top-right');
        alertify.error("Please Enter Link URL Correctly");
        isAllfield = false;
      }
      let myobjQl = {
        Title: $("#txtTitle").val().trim(),
        LinkURL: {
          "__metadata": {
            "type": "SP.FieldUrlValue"
          },
          Url: $('#txtHyper').val().trim(),
        }
        // ,Display: true
      }
      if (isAllfield) {
        $body.addClass("loading");
        await addItems("Quick Links", myobjQl);
          $body.addClass("loading");
          if (e.data) {
            $body.removeClass("loading");
            $('.addbutton').prop('disabled', true);
            // this.pageBack();
            window.history.back();
            // window.location.href = "" + siteweburl + "/Pages/ListView.aspx?CName=" + strLocalStorage + "";

          } else {
            $body.removeClass("loading");
            console.log(e);
          }
       
      }
    }  //Add News Part
    else if (strLocalStorage == "News") {
      var files = <HTMLInputElement>document.getElementById("inputImage");
      let file = files.files[0];
      if (!$('#inputImage').val().trim()) {
        alertify.set('notifier', 'position', 'top-right');
        alertify.error("Please Select the Image");
        isAllfield = false;
      } else if (!$('#txtExpires').val().trim()) {
        alertify.set('notifier', 'position', 'top-right');
        alertify.error("Please Select the Date");
        isAllfield = false;
      } else if (!$('#txtTitle').val().trim()) {
        alertify.set('notifier', 'position', 'top-right');
        alertify.error("Please Enter Title");
        isAllfield = false;
      } else if (!$('#txtrequiredDescription').val().trim()) {
        alertify.set('notifier', 'position', 'top-right');
        alertify.error("Please Enter Description");
        isAllfield = false;
      }
      else if ($('#canvasdisplay').css('display') == 'block') {
        alertify.set('notifier', 'position', 'top-right');
        alertify.error("Please Save Cropped Image");
        isAllfield = false;

      }
      if (isAllfield) {
        $body.addClass("loading");
        var uniquename = Math.random().toString(36).substr(2, 9) + ".png";
        var file1 = $('#cropped-img').attr('src').split("base64,");
        var blob = base64ToArrayBuffer(file1[1]);
        let myobjQl = {
          Title: $("#txtTitle").val().trim(),
          Date: new Date($("#txtExpires").val()),
          Explanation: $("#txtrequiredDescription").val(),
          Image: {
            "__metadata": {
              "type": "SP.FieldUrlValue"
            },
            Url: siteweburl + "/" + strLocalStorage + "/" + uniquename
          }
        }
        await additemsimage(strLocalStorage, uniquename, blob, myobjQl);
          if (e.data) {
            $body.removeClass("loading");
            $('.addbutton').prop('disabled', true);
            // this.pageBack();
            window.history.back();
            //  window.location.href = "" + siteweburl + "/Pages/ListView.aspx?CName=" + strLocalStorage + "";

          } else {
            console.log(e);
          }
      
      }
    }       //Add QuickReads Part
    else if (strLocalStorage == "Employee Corner") {
      var files = <HTMLInputElement>document.getElementById("uploadFile");
      let file = files.files[0];
      // var uniquename = Math.random().toString(36).substr(2, 9) + "." + file.name.substring(file.name.lastIndexOf(".") + 1, file.name.length);

      if (!$('#txtTitle').val().trim()) {
        // $('#txtTitle').focus();
        alertify.set('notifier', 'position', 'top-right');
        alertify.error("Please Enter Title");
        isAllfield = false;
      } else if (!$('#uploadFile').val().trim()) {
        //$('#uploadFile').focus();
        alertify.set('notifier', 'position', 'top-right');
        alertify.error("Please Select Document File");
        isAllfield = false;
      }
      let myobjQl = {
        Title: $("#txtTitle").val().trim(),
        DocumentFile: {
          "__metadata": {
            "type": "SP.FieldUrlValue"
          },
          Url: siteweburl + "/" + strLocalStorage + "/" + file.name
        }
        //,Display: true
      }
      if (isAllfield) {

        $body.addClass("loading");
        await additemsimage(strLocalStorage, file.name, file, myobjQl);
          if (e.data) {
            $body.removeClass("loading");
            $('.addbutton').prop('disabled', true);
            // this.pageBack();
            window.history.back();
            // window.location.href = "" + siteweburl + "/Pages/ListView.aspx?CName=" + strLocalStorage + "";
          } else {
            console.log(e);
          }
        
      }

    }       ////Add Banners Part
    else if (strLocalStorage == "Banners") {
      // var regexp = /^(https?|s?ftp):\/\/(((([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(%[\da-f]{2})|[!\$&'\(\)\*\+,;=]|:)*@)?(((\d|[1-9]\d|1\d\d|2[0-4]\d|25[0-5])\.(\d|[1-9]\d|1\d\d|2[0-4]\d|25[0-5])\.(\d|[1-9]\d|1\d\d|2[0-4]\d|25[0-5])\.(\d|[1-9]\d|1\d\d|2[0-4]\d|25[0-5]))|((([a-z]|\d|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(([a-z]|\d|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])*([a-z]|\d|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])))\.)+(([a-z]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(([a-z]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])*([a-z]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])))\.?)(:\d*)?)(\/((([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(%[\da-f]{2})|[!\$&'\(\)\*\+,;=]|:|@)+(\/(([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(%[\da-f]{2})|[!\$&'\(\)\*\+,;=]|:|@)*)*)?)?(\?((([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(%[\da-f]{2})|[!\$&'\(\)\*\+,;=]|:|@)|[\uE000-\uF8FF]|\/|\?)*)?(#((([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(%[\da-f]{2})|[!\$&'\(\)\*\+,;=]|:|@)|\/|\?)*)?$/i
      var files = <HTMLInputElement>document.getElementById("inputImage");
      let file = files.files[0];
      var siteImageURL = window.location.origin;
      if ($('#cropped-img')[0].src == '../../../_catalogs/masterpage/BloomHomepage/images/prof-img.jpg') {
        alertify.set('notifier', 'position', 'top-right');
        alertify.error("Please Select the Image");
        isAllfield = false;
      } else if (!$('#txtTitle').val().trim()) {
        alertify.set('notifier', 'position', 'top-right');
        alertify.error("Please Enter Title");
        isAllfield = false;
      }
      else if (!$('#txtrequiredDescription').val().trim()) {
        alertify.set('notifier', 'position', 'top-right');
        alertify.error("Please Enter Description");
        isAllfield = false;
      }
      else if ($('#canvasdisplay').css('display') == 'block') {
        alertify.set('notifier', 'position', 'top-right');
        alertify.error("Please Save Cropped Image");
        isAllfield = false;
      }
      var uniquename = Math.random().toString(36).substr(2, 9) + ".png";
      var file1 = $('#cropped-img').attr('src').split("base64,");
      var blob = base64ToArrayBuffer(file1[1]);
      var siteurl = this.context.pageContext.web.absoluteUrl;
      let myobjQl = {
        Title: $("#txtTitle").val().trim(),
        BannerContent: $("#txtrequiredDescription").val(),
        Image: {
          "__metadata": {
            "type": "SP.FieldUrlValue"
          },
          Url: siteurl + "/" + strLocalStorage + "/" + uniquename
        },
        LinkURL: {
          "__metadata": {
            "type": "SP.FieldUrlValue"
          },
          Url: $('#txtHyper').val(),
        }
      }//      additemsimage(strLocalStorage, uniquename, blob, myobjQl, function (e)
      if (isAllfield) {
        $body.addClass("loading");
        await additemsimage(strLocalStorage, uniquename, blob, myobjQl);
        $body.removeClass("loading");
        $('.addbutton').prop('disabled', true);
        window.history.back();
      }
    }     ////Add New ImageGallery Part

    else if (strLocalStorage == "Image Gallery") {

      var regexp = /^(https?|s?ftp):\/\/(((([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(%[\da-f]{2})|[!\$&'\(\)\*\+,;=]|:)*@)?(((\d|[1-9]\d|1\d\d|2[0-4]\d|25[0-5])\.(\d|[1-9]\d|1\d\d|2[0-4]\d|25[0-5])\.(\d|[1-9]\d|1\d\d|2[0-4]\d|25[0-5])\.(\d|[1-9]\d|1\d\d|2[0-4]\d|25[0-5]))|((([a-z]|\d|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(([a-z]|\d|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])*([a-z]|\d|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])))\.)+(([a-z]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(([a-z]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])*([a-z]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])))\.?)(:\d*)?)(\/((([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(%[\da-f]{2})|[!\$&'\(\)\*\+,;=]|:|@)+(\/(([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(%[\da-f]{2})|[!\$&'\(\)\*\+,;=]|:|@)*)*)?)?(\?((([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(%[\da-f]{2})|[!\$&'\(\)\*\+,;=]|:|@)|[\uE000-\uF8FF]|\/|\?)*)?(#((([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(%[\da-f]{2})|[!\$&'\(\)\*\+,;=]|:|@)|\/|\?)*)?$/i
      var files = <HTMLInputElement>document.getElementById("uploadImageFile");
      let file = files.files[0];
      if (!$('#uploadImageFile').val().trim()) {
        alertify.set('notifier', 'position', 'top-right');
        alertify.error("Please Select the Image");
        isAllfield = false;
      } else if (!$('#txtFolderName').val().trim()) {
        alertify.set('notifier', 'position', 'top-right');
        alertify.error("Please Enter Title");
        isAllfield = false;
      }

      if (isAllfield) {
        $body.addClass("loading");

        pnp.sp.web.lists.getByTitle("Image Gallery").rootFolder.folders.add($('#txtFolderName').val())
          .then(data => {

            pnp.sp.web.getFolderByServerRelativeUrl("Image Gallery" + "/" + $('#txtFolderName').val()).files.add(file.name, file, true)
              .then(function (result: any) {

                $body.removeClass("loading");
                window.history.back();
                //  window.location.href = "" + siteweburl + "/Pages/ListView.aspx?CName=" + strLocalStorage + "";

              })

          })

      }
    }
    else if (strLocalStorage == "Video Gallery") {

      var regexp = /^(https?|s?ftp):\/\/(((([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(%[\da-f]{2})|[!\$&'\(\)\*\+,;=]|:)*@)?(((\d|[1-9]\d|1\d\d|2[0-4]\d|25[0-5])\.(\d|[1-9]\d|1\d\d|2[0-4]\d|25[0-5])\.(\d|[1-9]\d|1\d\d|2[0-4]\d|25[0-5])\.(\d|[1-9]\d|1\d\d|2[0-4]\d|25[0-5]))|((([a-z]|\d|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(([a-z]|\d|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])*([a-z]|\d|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])))\.)+(([a-z]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(([a-z]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])*([a-z]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])))\.?)(:\d*)?)(\/((([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(%[\da-f]{2})|[!\$&'\(\)\*\+,;=]|:|@)+(\/(([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(%[\da-f]{2})|[!\$&'\(\)\*\+,;=]|:|@)*)*)?)?(\?((([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(%[\da-f]{2})|[!\$&'\(\)\*\+,;=]|:|@)|[\uE000-\uF8FF]|\/|\?)*)?(#((([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(%[\da-f]{2})|[!\$&'\(\)\*\+,;=]|:|@)|\/|\?)*)?$/i
      var files = <HTMLInputElement>document.getElementById("uploadImageFile");
      let file = files.files[0];
      var radioValue = $("input[name='selectionradioImage']:checked").val();

      if(radioValue == "Upload")
      {
        if(!$('#txtFolderName').val().trim()){
          alertify.set('notifier', 'position', 'top-right');
          alertify.error("Please enter the Folder name");
          isAllfield = false;
        }
        else if(!$('#txtTitle').val().trim()){
          alertify.set('notifier', 'position', 'top-right');
          alertify.error("Please enter the Title");
          isAllfield = false;
        }
        else if (!$('#uploadImageFile').val().trim()) {
          alertify.set('notifier', 'position', 'top-right');
          alertify.error("Please Select the video File");
          isAllfield = false;
        }
        
      }
      if(radioValue == "Stream")
      {
        if(!$('#txtFolderName').val().trim()){
          alertify.set('notifier', 'position', 'top-right');
          alertify.error("Please enter the Folder name");
          isAllfield = false;
        }
        else if(!$('#txtTitle').val().trim()){
          alertify.set('notifier', 'position', 'top-right');
          alertify.error("Please enter the Title");
          isAllfield = false;
        }
        else if (!$('#txtHyper').val().trim()) {
          alertify.set('notifier', 'position', 'top-right');
          alertify.error("Please Enter the URL");
          isAllfield = false;
        }
        
      }

      if (isAllfield) {
        
        $body.addClass("loading");

        pnp.sp.web.lists.getByTitle("Video Gallery").rootFolder.folders.add($('#txtFolderName').val())
          .then(data => {

        if(radioValue == "Upload")
        {
          var VideoTitle = { Title: $('#txtTitle').val().trim() };
          pnp.sp.web.getFolderByServerRelativeUrl("Video Gallery"+ "/" + $('#txtFolderName').val()).files.add(file.name, file, true)
            .then(({ file }) => file.getItem())
            .then(item => item.update(VideoTitle))
            .then(function (result: any) {
                   $body.removeClass("loading");
                   window.history.back();
            });
        }
        else{
          var Videojson = {
              Title: $("#txtTitle").val(),
              LinkURL: {
              "__metadata": { "type": "SP.FieldUrlValue" },
              Url: $('#txtHyper').val().trim()
              }
          };

          if ($("#uploadImageFile").val() == undefined || $("#uploadImageFile").val() == null || $("#uploadImageFile").val() == '') {
            $.ajax({
            url: this.context.pageContext.site.absoluteUrl + "/_api/web/getfilebyserverrelativeurl('/sites/BloomHolding/_catalogs/masterpage/Bloom/images/logo.png')/openbinarystream",
            type: "GET",
            success: function (data) {
            var name = $("#txtTitle").val().trim()+'.jpg';
              pnp.sp.web.getFolderByServerRelativeUrl("Video Gallery"+ "/" + $('#txtFolderName').val()).files.add(name, data, true)
              .then(({ file }) => file.getItem())
              .then(item => item.update(Videojson))
              .then(function (result: any) {
                    $body.removeClass("loading");
                    window.history.back();
              });
            },
            error: function (data) {
            console.log(data);
            },
            });
            }

        }
      });

      }
    }
    //Add Organizational Policies Part

    else if (strLocalStorage == "Organizational Policies") {
      var files = <HTMLInputElement>document.getElementById("uploadFile");
      let file = files.files[0];
      if (!$('#txtTitle').val().trim()) {
        // $('#txtTitle').focus();
        alertify.set('notifier', 'position', 'top-right');
        alertify.error("Please Enter Title");
        isAllfield = false;
      } else if ($('#ddlDepartment').val() == "Select") {
        //$('#uploadFile').focus();
        alertify.set('notifier', 'position', 'top-right');
        alertify.error("Please Select Department");
        isAllfield = false;
      } else if (!$('#txtDepartment').val().trim()) {
        //$('#txtDepartment').focus();
        alertify.set('notifier', 'position', 'top-right');
        alertify.error("Please Enter the Department Name");
        isAllfield = false;
      } else if (!$('#uploadFile').val().trim()) {
        //$('#uploadFile').focus();
        alertify.set('notifier', 'position', 'top-right');
        alertify.error("Please Select File");
        isAllfield = false;
      }
      let myobjQl = {
        Title: $("#txtTitle").val(),
        Departments: $("#txtDepartment").val(),
        Explanation: $("#txtDescription").val().trim(),
        DocumentFile: {
          "__metadata": {
            "type": "SP.FieldUrlValue"
          },
          Url: siteweburl + "/" + strLocalStorage + "/" + file.name
        }
      }
      if (isAllfield) {
        //  var uniquename = Math.random().toString(36).substr(2, 9) + "." + file.name.substring(file.name.lastIndexOf(".") + 1, file.name.length);
        $body.addClass("loading");
        var _thiss = this;
        await additemsimage(strLocalStorage, file.name, file, myobjQl);
          if (e.data) {
            $body.removeClass("loading");
            $('.addbutton').prop('disabled', true);
            // this.pageBack();
            if ($('#ddlDepartment').val() == "Others") {
              _thiss.AddDepartments();
            }
            else {
              window.history.back();
            }
            //window.location.href = "" + siteweburl + "/Pages/OrganizationalPolicies.aspx";
          } else {
            console.log(e);
          }
        
      }

    }     //Add Corporate Discounts Part
    else if (strLocalStorage == "Corporate Discounts") {
      var docfiles = <HTMLInputElement>document.getElementById("uploadFile");
      let docfile = docfiles.files[0];
      var files = <HTMLInputElement>document.getElementById("inputImage");
      let file = files.files[0];
      var regexp = /^(https?|s?ftp):\/\/(((([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(%[\da-f]{2})|[!\$&'\(\)\*\+,;=]|:)*@)?(((\d|[1-9]\d|1\d\d|2[0-4]\d|25[0-5])\.(\d|[1-9]\d|1\d\d|2[0-4]\d|25[0-5])\.(\d|[1-9]\d|1\d\d|2[0-4]\d|25[0-5])\.(\d|[1-9]\d|1\d\d|2[0-4]\d|25[0-5]))|((([a-z]|\d|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(([a-z]|\d|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])*([a-z]|\d|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])))\.)+(([a-z]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(([a-z]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])*([a-z]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])))\.?)(:\d*)?)(\/((([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(%[\da-f]{2})|[!\$&'\(\)\*\+,;=]|:|@)+(\/(([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(%[\da-f]{2})|[!\$&'\(\)\*\+,;=]|:|@)*)*)?)?(\?((([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(%[\da-f]{2})|[!\$&'\(\)\*\+,;=]|:|@)|[\uE000-\uF8FF]|\/|\?)*)?(#((([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(%[\da-f]{2})|[!\$&'\(\)\*\+,;=]|:|@)|\/|\?)*)?$/i
      if (!$('#txtTitle').val().trim()) {
        //$('#txtTitle').focus();
        alertify.set('notifier', 'position', 'top-right');
        alertify.error("Please Enter Title");
        isAllfield = false;
      } else if (!$('#txtsitelink').val().trim()) {
        //$('#txtsitelink').focus();
        alertify.set('notifier', 'position', 'top-right');
        alertify.error("Please Enter SiteLink");
        isAllfield = false;
      } else if (!regexp.test($('#txtsitelink').val().trim())) {
        //$('#txtsitelink').focus();
        alertify.set('notifier', 'position', 'top-right');
        alertify.error("Please Enter SiteLink Correctly");
        isAllfield = false;
      }
      else if ($('#canvasdisplay').css('display') == 'block') {
        alertify.set('notifier', 'position', 'top-right');
        alertify.error("Please Save Cropped Image");
        isAllfield = false;

      }
      if (isAllfield) {

        var fileURL = window.location.origin;
        $body.addClass("loading");

        if (files.files.length == 0 && docfiles.files.length > 0) {

          var name = Math.random().toString(36).substr(2, 9) + ".png";
          pnp.sp.web.getFolderByServerRelativeUrl("Shared Documents").files.add(docfile.name, docfile, true)
            .then(function (datefile) {
              let CorDocOnly = {
                Title: $("#txtTitle").val().trim(),
                SiteLink: {
                  "__metadata": {
                    "type": "SP.FieldUrlValue"
                  },
                  Url: $('#txtsitelink').val().trim(),
                },
                VendorLogo: {
                  "__metadata": {
                    "type": "SP.FieldUrlValue"
                  },
                  Url: siteweburl + "/_catalogs/masterpage/Bloom/images/logo.png"
                },
                DocumentFile: {
                  "__metadata": {
                    "type": "SP.FieldUrlValue"
                  },
                  Url: fileURL + datefile.data.ServerRelativeUrl
                }
              }
              $body.addClass("loading");
               additemsimage(strLocalStorage, name, docfile, CorDocOnly);
                if (e.data) {
                  $body.removeClass("loading");
                  $('.addbutton').prop('disabled', true);
                  // this.pageBack();
                  window.history.back();
                  //  window.location.href = "" + siteweburl + "/Pages/ListView.aspx?CName=" + strLocalStorage + "";

                } else {
                  console.log(e);
                }
              });
            
        }
        else if (docfiles.files.length == 0 && files.files.length > 0) {
          var file1 = $('#cropped-img').attr('src').split("base64,");
          var blob = base64ToArrayBuffer(file1[1]);
          //var filename = $('#inputImage').val().replace(/C:\\fakepath\\/i, '');
          var uniquename = Math.random().toString(36).substr(2, 9) + ".png";
          // pnp.sp.web.getFolderByServerRelativeUrl("Corporate Discounts").files.add(file.name, blob, true)
          // .then(function (result)
          //  {
          let myobjQl = {
            Title: $("#txtTitle").val().trim(),
            SiteLink: {
              "__metadata": {
                "type": "SP.FieldUrlValue"
              },
              Url: $('#txtsitelink').val().trim(),
            },
            VendorLogo: {
              "__metadata": {
                "type": "SP.FieldUrlValue"
              },
              Url: siteweburl + "/" + strLocalStorage + "/" + uniquename
            },

          }
          $body.addClass("loading");
          await additemsimage(strLocalStorage, uniquename, blob, myobjQl);
            if (e.data) {
              $body.removeClass("loading");
              $('.addbutton').prop('disabled', true);
              // this.pageBack();
              window.history.back();
              //  window.location.href = "" + siteweburl + "/Pages/ListView.aspx?CName=" + strLocalStorage + "";

            } else {
              console.log(e);
            }
          

        } else if (files.files.length == 0 && docfiles.files.length == 0) {

          let myobjQl = {
            Title: $("#txtTitle").val().trim(),
            SiteLink: {
              "__metadata": {
                "type": "SP.FieldUrlValue"
              },
              Url: $('#txtsitelink').val().trim(),
            },
            VendorLogo: {
              "__metadata": {
                "type": "SP.FieldUrlValue"
              },
              Url: siteweburl + "/_catalogs/masterpage/Bloom/images/logo.png"
            },
          }
          $body.addClass("loading");
          await additemsimage(strLocalStorage, uniquename, blob, myobjQl);
            if (e.data) {
              $body.removeClass("loading");
              $('.addbutton').prop('disabled', true);
              // this.pageBack();
              window.history.back();
              //  window.location.href = "" + siteweburl + "/Pages/ListView.aspx?CName=" + strLocalStorage + "";

            } else {
              console.log(e);
            }
          

        }

        else {
          var file1 = $('#cropped-img').attr('src').split("base64,");
          var blob = base64ToArrayBuffer(file1[1]);
          //var filename = $('#inputImage').val().replace(/C:\\fakepath\\/i, '');
          var uniquename = Math.random().toString(36).substr(2, 9) + ".png";
          pnp.sp.web.getFolderByServerRelativeUrl("Corporate Discounts").files.add(uniquename, blob, true)
            .then(function (result) {
              result.file.listItemAllFields.get().then((listItemAllFields) => {
                pnp.sp.web.getFolderByServerRelativeUrl("Shared Documents").files.add(docfile.name, docfile, true)
                  .then(function (datefile) {
                    pnp.sp.web.lists.getByTitle("Corporate Discounts").items.getById(listItemAllFields.Id).update({
                      Title: $("#txtTitle").val().trim(),
                      SiteLink: {
                        "__metadata": {
                          "type": "SP.FieldUrlValue"
                        },
                        Url: $('#txtsitelink').val().trim(),
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
                        Url: fileURL + datefile.data.ServerRelativeUrl
                      },
                    }).then(r => {
                      $body.removeClass("loading");
                      $('.addbutton').prop('disabled', true);
                      window.history.back();
                    });

                  });
              });
            });

        }
      }
    }     //Add Events Part
    else if (dropOptionValue == "Events") {
      var files = <HTMLInputElement>document.getElementById("inputImage");
      let file = files.files[0];
      if (!$('#txtTitle').val().trim()) {
        //$('#txtTitle').focus();
        alertify.set('notifier', 'position', 'top-right');
        alertify.error("Please Enter Title");
        isAllfield = false;
      } else if (!$('#txtrequiredDescription').val().trim()) {
        alertify.set('notifier', 'position', 'top-right');
        alertify.error("Please Enter Description");
        isAllfield = false;
      } else if (!$('#txtStartDate').val().trim()) {
        // $('#txtStartDate').focus();
        alertify.set('notifier', 'position', 'top-right');
        alertify.error("Please Select StartDate");
        isAllfield = false;
      } else if (Date.parse($("#txtStartDate").val().trim()) > Date.parse($("#txtEndDate").val().trim())) {
        alertify.set('notifier', 'position', 'top-right');
        alertify.error("Start date must be less than End date");
        return false;
      }
      else if ($('#canvasdisplay').css('display') == 'block') {
        alertify.set('notifier', 'position', 'top-right');
        alertify.error("Please Save Cropped Image");
        isAllfield = false;

      }
      if (isAllfield) {
        var fileURL = window.location.origin;
        $body.addClass("loading");
        if (files.files.length > 0) {

          var name = Math.random().toString(36).substr(2, 9) + ".png";
          var file1 = $('#cropped-img').attr('src').split("base64,");
          var blob = base64ToArrayBuffer(file1[1]);
          pnp.sp.web.getFolderByServerRelativeUrl("Events").files.add(name, blob, true)
            .then(function (result) {
              result.file.listItemAllFields.get().then((listItemAllFields) => {
                pnp.sp.web.lists.getByTitle("Events").items.getById(listItemAllFields.Id).update({
                  Title: $("#txtTitle").val().trim(),
                  StartDate: new Date($('#txtStartDate').val()),
                  EndDate: new Date($('#txtEndDate').val()),
                  Explanation: $('#txtrequiredDescription').val(),
                  Image: {
                    "__metadata": {
                      "type": "SP.FieldUrlValue"
                    },
                    Url: siteweburl + "/" + "Events" + "/" + name
                  }
                }).then(r => {
                  $body.removeClass("loading");
                  $('.addbutton').prop('disabled', true);
                  // this.pageBack();
                  window.history.back();
                  //   window.location.href = "" + siteweburl + "/Pages/ListView.aspx?CName=" + strLocalStorage + "";
                  //console.log(name + " properties updated successfully!");
                });
              });
            })

        } else {
          $body.addClass("loading");
          let myobjQl = {
            Title: $("#txtTitle").val().trim(),
            StartDate: new Date($('#txtStartDate').val()),
            EndDate: new Date($('#txtEndDate').val()),
            Explanation: $('#txtrequiredDescription').val(),
            Image: {
              "__metadata": {
                "type": "SP.FieldUrlValue"
              },
              Url: siteweburl + "/_catalogs/masterpage/Bloom/images/logo.png"
            }
          }

          await addItems(strLocalStorage, myobjQl);
            if (e.data) {
              $body.removeClass("loading");
              $('.addbutton').prop('disabled', true);
              // this.pageBack();
              window.history.back();
              // window.location.href = "" + siteweburl + "/Pages/ListView.aspx?CName=" + strLocalStorage + "";

            } else {
              console.log(e);
            }
          
        }
      }

    }    //Add Polls Part
    else if (strLocalStorage == "Polls") {
      var optionseperate = $('#txtOptions').val();
      var resultarray = optionseperate.split(";");
      var newArray = resultarray.filter(function (v) {
        return v !== ' '
      });

      if (!$('#txtQuestion').val().trim()) {
        //$('#txtQuestion').focus();
        alertify.set('notifier', 'position', 'top-right');
        alertify.error("Please Enter Question");
        isAllfield = false;
      } else if (!$('#txtOptions').val().trim() || newArray.length <= 1) {
        // $('#txtOptions').focus();
        alertify.set('notifier', 'position', 'top-right');
        alertify.error("Please Enter Answers Correctly");
        isAllfield = false;
      } else if (!$('#txtOptions').val().trim() || newArray.length >= 5) {
        alertify.set('notifier', 'position', 'top-right');
        alertify.error("Maximum Five Answers only Allowed");
        return false;
      }
      let myobjPols = {
        Question: $("#txtQuestion").val().trim(),
        Options: $("#txtOptions").val().trim(),
      }
      if (isAllfield) {
        $body.addClass("loading");
        await addItems("Polls", myobjPols);
          $body.addClass("loading");
          if (e.data) {
            $body.removeClass("loading");
            $('.addbutton').prop('disabled', true);
            // this.pageBack();
            window.history.back();
            // window.location.href = "" + siteweburl + "/Pages/ListView.aspx?CName=" + strLocalStorage + "";

          } else {
            $body.removeClass("loading");
            console.log(e);
          }
        

      }
    }
  }



  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }
  
  public bindgalleryImage() {
    pnp.sp.web.folders.getByName('ImageGalleryDocLib').folders.get().then(function (data) {
      let arrayname = [];
      for (var i = 0; i < data.length; i++) {
        arrayname.push(data[i].Name)

      }
      let SearchResult = arrayname.concat(arrayname);
      var uniqueArray = SearchResult.filter(function (elem, pos, arr) {
        return arr.indexOf(elem) == pos;
      });
      $("#txtTitle").autocomplete({ source: uniqueArray });
    }).catch(function (data) {

    });

  }

  public bindorgDept() {
    let objResults = readItems("Departments", ["ID", "Departments"], 1000, "Modified", "Display", 1);
    objResults.then((items: any[]) => {
      let arrayname = [];
      var DeptHTML = "";
      DeptHTML += "<option id='0' disabled selected>Select</option>";
      for (var i = 0; i < items.length; i++) {
        arrayname.push({
          "dept": items[i].Departments,
          "id": items[i].ID
        })
        DeptHTML += "<option id='" + items[i].ID + "'>" + items[i].Departments + "</option>";
      }
      DeptHTML += "<option id='" + i + 1 + "'>Others</option>";

      $("#ddlDepartment").append(DeptHTML);
    })

    $("#DivDepartment").hide();

    $('#ddlDepartment').change(function () {
      $("#txtDepartment").val($(this).val());
      if ($(this).val() == "Others") {
        $("#DivDepartment").show();
        $("#txtDepartment").val("");
      }
      else {
        $("#DivDepartment").hide();
      }
    });
  }

  public AddListItems() {

      var strLocalStorage = GetQueryStringParams("CName");
      strLocalStorage = strLocalStorage.split("%20").join(' ');
      var renderUploadImagefile = "";
      var renderUploadVideofile = "";
      var renderhtmlImage = "";
      var rendertext = "";
      var renderFoldertext = "";
      var renderdate = "";
      var renderDescription = "";
      var renderEventDate = "";
      var renderHyperlink = "";
      var renderHyperSitelink = "";
      var renderUploadfile = "";
      var renderUploadOrganization = "";
      var renderSiteLink = "";
      var renderStartEndDate = "";
      var renderhtmlImageEvents = "";
      var renderQuestion = "";
      var renderAnswers = "";
      var renderDepartment = "";
      var rendercrop = "";
      var renderRequiredDescription = "";
      var renderNews = "";
      var renderhtmlCorporateImage = "";
      var renderCorpUploadfile = "";
      var renderOptionImageGallery = "";
      var renderdropdown = "";
      var renderOption = "";
      var renderDepartmentddl = "";
      var strLocalStorage = GetQueryStringParams("CName");
      strLocalStorage = strLocalStorage.split("%20").join(' ');
  
      var radioValue = $("input[name='selectionradio']:checked").val();
      if (radioValue == "Holiday") {
        strLocalStorage = "Holiday";
  
      } else if (radioValue == "Events") {
        strLocalStorage = "Events";
      }
      var siteweburl = this.context.pageContext.web.absoluteUrl;
      renderOption += "<div class='radio-btn appendOption'>" +
        "<div class='col-md-12 form-group'>" +
        "<label>Choose Component</label>" +
        "<div class='radio col-md-6'>" +
        "<input id='radio-1' name='selectionradio' type='radio' value='Events'>" +
        "<label for='radio-1' class='radio-label'>Events</label>" +
        "</div>" +
        "<div class='radio col-md-6'>" +
        "<input id='radio-2' name='selectionradio' type='radio' value='Holiday'>" +
        "<label for='radio-2' class='radio-label'>Holidays</label>" +
        "</div>" +
        "</div>" +
        "</div>";
  
      renderOptionImageGallery += "<div class='radio-btn appendOptionImage'>" +
        "<div class='col-md-12 form-group'>" +
        "<label>Choose Component</label>" +
        "<div class='radio col-md-6'>" +
        "<input id='radio-3' name='selectionradioImage' type='radio' value='Upload'>" +
        "<label for='radio-3' class='radio-label'>Upload</label>" +
        "</div>" +
        "<div class='radio col-md-6'>" +
        "<input id='radio-4' name='selectionradioImage' type='radio' value='Stream'>" +
        "<label for='radio-4' class='radio-label'>Stream</label>" +
        "</div>" +
        "</div>" +
        "</div>";
  
      renderhtmlCorporateImage += "<div class='form-imgsec'>" +
        "<div class='themelogo-upload'>" +
        "<label>Vendor Logo</label>" +
        "<img id='cropped-img' class='crapImagesevent crop-imagedisplay' src= '../../../_catalogs/masterpage/BloomHomepage/images/prof-img.jpg'>" +
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
  
      renderhtmlImage += "<div class='form-imgsec'>" +
        "<div class='themelogo-upload'>" +
        "<label class='control-label'>Image</label>" +
        "<img src='../../../_catalogs/masterpage/BloomHomepage/images/prof-img.jpg'>" +
        "<div class='image-upload' id='imagerestrict'>" +
        "<div class='custom-upload'>" +
        "<input type='file' id='inputImage' name='file' accept='.jpg,.jpeg,.png,.gif,.bmp,.tiff' multiple='' class='file' />" +
        "<button class='btn btn-primary' type='button'><i class='icon-camera'></i></button>" +
        "</div>" +
        "<a href='#' title='Delete'>" +
        "<i class='icon-delete'></i></a>" +
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
        "<img id='cropped-img'  src='../../../_catalogs/masterpage/BloomHomepage/images/prof-img.jpg'>" +
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
  
      renderNews += "<div class='form-imgsec'>" +
        "<div class='themelogo-upload'>" +
        "<label class='control-label'>Image</label>" +
        "<img id='cropped-img' src='../../../_catalogs/masterpage/BloomHomepage/images/prof-img.jpg' />" +
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
  
      rendertext += "<div class='input text'>" +
        "<label class='control-label'>Title</label>" +
        "<input class='form-control' type='text' value='' maxlength='30' id='txtTitle' /></div>";
  
      renderFoldertext += "<div id='foldername' class='input text'>" +
        "<label class='control-label'>Folder Name</label>" +
        "<input class='form-control' type='text' value='' id='txtFolderName' /></div>";
  
      renderdate += "<div class='input date'>" + "<i class='icon-calenter'></i>" +
        "<label class='control-label'>Date</label>" +
        "<input class='form-control date-selector' type='text' value='' id='txtExpires' /></div>";
  
      renderDescription += "<div id='descrption' class='input textarea'><label>Description</label><textarea class='form-control' id='txtDescription'></textarea></div>";
  
      renderEventDate += "<div class='input date'>" +
        "<i class='icon-calenter'></i>" +
        "<label class='control-label'>Start Date</label>" +
        "<input class='form-control date-selector' type='text' value='' id='txtEvDate' />" +
        "</div>" +
        "<div class='input date'>" +
        "<i class='icon-calenter'></i>" +
        "<label>End Date</label>" +
        "<input class='form-control date-selector' type='text' value='' id='txtEEDate' />" + "</div>";
  
      renderHyperlink += "<div class='input text' id='divHyperLink'>" +
        "<label class='control-label'>Link URL</label>" +
        "<input class='form-control' type='text' value='' id='txtHyper' />" +
        "<span>Please enter the Link URL in the following format : https://www.bloomholding.com</span>" +
        "</div>";
  
      renderHyperSitelink += "<div class='input text'>" +
        "<label>Link URL</label>" +
        "<input class='form-control' type='text' value='' id='txtHyper' />" +
        "<label>Please given valid Announcements or Events URL</label>" +
        "</div>";
  
      /* renderUploadfile += "<div class='form-imgsec'>" +
         "<div class='themelogo-upload' style='display: block;'>" +
         "<label class='control-label'>DocumentFile</label>" +
         "<div class='custom-upload banner-upload'>" +
         "<input type='file' id='inputImage' name='file' accept='.doc,.docx,.xls,.ppt,.pdf' multiple='' class='file'>" +
         "<div class='input-group'>" +
         "<span class='input-group-btn input-group-sm'>" +
         "<button type='button' class='btn btn-fab btn-fab-mini'>Browse</button>" +
         "</span>" +
         "<input type='text' readonly='' class='form-control' placeholder='Upload Files'>" +
         "</div>" +
         "</div>" +
         "</div>" +
         "</div>";*/
      renderUploadfile += "<div class='form-imgsec'>" +
        "<div class='themelogo-upload' style='display: block;'>" +
        "<div class='custom-upload banner-upload'>" +
        "<label class='control-label'>Document File</label>" +
        "<input type='file' id='uploadFile' name='file' accept='.doc,.docx,.xls,.ppt,.pdf' multiple='' class='file'>" +
        "<div class='input-group'>" +
        "<span class='input-group-btn input-group-sm'>" +
        "<button type='button' class='btn btn-fab btn-fab-mini'>Browse</button>" +
        "</span>" +
        "<input type='text' readonly='' class='form-control' placeholder='Upload Files'>" +
        "</div>" +
        "</div>" +
        "</div>" +
        "</div>";
  
      renderUploadImagefile += "<div class='form-imgsec'>" +
        "<div class='themelogo-upload' style='display: block;'>" +
        "<div class='custom-upload banner-upload'>" +
        "<label class='control-label'>Upload Image File</label>" +
        "<input type='file' id='uploadImageFile' name='file' accept='.jpg,.jpeg,.png,.gif,.bmp,.tiff' multiple='' class='file'>" +
        "<div class='input-group'>" +
        "<span class='input-group-btn input-group-sm'>" +
        "<button type='button' class='btn btn-fab btn-fab-mini'>Browse</button>" +
        "</span>" +
        "<input type='text' readonly='' class='form-control' placeholder='Upload Files'>" +
        "</div>" +
        "</div>" +
        "</div>" +
        "</div>";
  
      renderUploadVideofile += "<div class='form-imgsec'>" +
        "<div class='themelogo-upload' style='display: block;'>" +
        "<div class='custom-upload banner-upload'>" +
        "<label class='control-label'>Upload Video File</label>" +
        "<input type='file' id='uploadImageFile' name='file' accept='video/mp4,video/x-m4v,video/*' multiple='' class='file'>" +
        "<div class='input-group'>" +
        "<span class='input-group-btn input-group-sm'>" +
        "<button type='button' class='btn btn-fab btn-fab-mini'>Browse</button>" +
        "</span>" +
        "<input type='text' readonly='' class='form-control' placeholder='Upload Files'>" +
        "</div>" +
        "</div>" +
        "</div>" +
        "</div>";
  
      renderUploadOrganization += "<div class='form-imgsec'>" +
        "<div class='themelogo-upload' style='display: block;'>" +
        "<label class='control-label'>DocumentFile</label>" +
        "<div class='custom-upload banner-upload'>" +
        "<input type='file' id='inputImage' name='file' accept='.pdf,.doc,.docx' multiple='' class='file'>" +
        "<div class='input-group'>" +
        "<span class='input-group-btn input-group-sm'>" +
        "<button type='button' class='btn btn-fab btn-fab-mini'>Browse</button>" +
        "</span>" +
        "<input type='text' readonly='' class='form-control' placeholder='Upload Files'>" +
        "</div>" +
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
        "<span>Please Enter more than one answers with Semicolon ( ; ) Maximun Four answers</span>" +
        "</div>";
      renderDepartment += "<div class='input text' id='DivDepartment'>" +
        "<i class=''></i>" +
        "<label class='control-label'>New Department</label>" +
        "<input class='form-control' type='text' value='' id='txtDepartment' autocomplete='off'/>" +
        "</div>";
      renderDepartmentddl += '<div class="input text">' +
        '<label class="control-label">Department</label>' +
        '<select id="ddlDepartment" class="form-control">' +
        '</select>' +
        '</div>';
      renderRequiredDescription += "<div id='rrdescription' class='input textarea'><label class='control-label'>Description</label><textarea class='form-control' id='txtrequiredDescription'></textarea></div>";
      var date = new Date();
      var today = new Date(date.getFullYear(), date.getMonth(), date.getDate());
  
      if (strLocalStorage == 'Announcements') {
        $('.form-section').append(renderhtmlImageEvents);
        $('.form-imgsec').after(rendercrop);
        $('#canvasdisplay').after(renderdate);
        $('.date').after(rendertext);
        $('.text').after(renderRequiredDescription);
        $('#txtExpires').datepicker({
          format: "mm/dd/yyyy",
          startDate: today
  
        }).datepicker('setDate', 'new Date()');
      } else if (strLocalStorage == 'Holiday') {
        /*$('.form-section').append(rendertext);
        $('.text').after(renderEventDate);
        $('#txtEvDate').datepicker({
          format: "mm/dd/yyyy",
          startDate: today
  
        }).datepicker('setDate', 'new Date()');
        $('#txtEEDate').datepicker({
          format: "mm/dd/yyyy",
          startDate: today
  
        });*/
  
        $('.form-section').append(renderOption);
        $('.appendOption').after(renderhtmlImageEvents);
        $('.form-imgsec').after(rendercrop);
        $('#canvasdisplay').after(rendertext)
        $('.text').after(renderRequiredDescription);
        $('.textarea').after(renderStartEndDate);
        $('#txtStartDate').datepicker({
          Format: 'yy/dd/mm',
          startDate: today
        }).datepicker('setDate', 'new Date()');
        $('#txtEndDate').datepicker({
          Format: 'yy/dd/mm',
          startDate: today
        });
  
  
  
      } else if (strLocalStorage == 'News') {
        $('.form-section').append(renderNews);
        $('.form-imgsec').after(rendercrop);
        $('#canvasdisplay').after(renderdate);
        $('.date').after(rendertext);
        $('.text').after(renderRequiredDescription);
  
        $('#txtExpires').datepicker({
          format: "mm/dd/yyyy",
          startDate: today
        }).datepicker('setDate', 'new Date()');
      } else if (strLocalStorage == 'Quick Links') {
  
        $('.form-section').append(rendertext);
        $('.text').after(renderHyperlink);
  
      } else if (strLocalStorage == 'Employee Corner') {
        $('.form-section').append(rendertext);
        $('.text').after(renderUploadfile);
      }
      // else if (strLocalStorage == 'AddNewImageGallery') {
      //   $('.form-section').append(renderNews);
      //   $('.form-imgsec').after(rendercrop);
      //   $('#canvasdisplay').after(renderFoldertext);
      //   this.bindgalleryImage();
  
      // }
      else if (strLocalStorage == 'Image Gallery') {
        $('.form-section').append(renderFoldertext);
        $('.text').after(renderUploadImagefile);
        this.bindgalleryImage();
  
      }
      else if (strLocalStorage == 'Video Gallery') {
        $('.form-section').append(renderOptionImageGallery+renderFoldertext+rendertext+renderUploadVideofile+renderHyperlink);
        $('.banner-upload').show();
        $('#divHyperLink').hide();  
      }
      else if (strLocalStorage == 'Organizational Policies') {
        $('.form-section').append(rendertext);
        $('.text').after(renderDepartmentddl + renderDepartment + renderUploadfile + renderDescription);
        //$('#txtDepartment').after(renderUploadfile);
        //$('.form-imgsec').after(renderDescription);
        this.bindorgDept()
      } else if (strLocalStorage == 'Banners') {
        $('.form-section').append(renderNews);
        $('.form-imgsec').after(rendercrop);
        $('#canvasdisplay').after(rendertext);
        $('.text').after(renderRequiredDescription);
        $('#rrdescription').after(renderHyperSitelink);
  
      } else if (strLocalStorage == 'Corporate Discounts') {
        $('.form-section').append(renderhtmlCorporateImage);
        $('.form-imgsec').after(rendercrop);
        $('#canvasdisplay').after(rendertext);
        $('.text').after(renderSiteLink);
        $('#siteLink').after(renderCorpUploadfile);
  
      } else if (strLocalStorage == 'Events') {
        $('.form-section').append(renderOption);
        $('.appendOption').after(renderhtmlImageEvents);
        $('.form-imgsec').after(rendercrop);
        $('#canvasdisplay').after(rendertext)
        $('.text').after(renderRequiredDescription);
        $('.textarea').after(renderStartEndDate);
        $('#txtStartDate').datepicker({
          Format: 'yy/dd/mm',
          startDate: today
        }).datepicker('setDate', 'new Date()');
        $('#txtEndDate').datepicker({
          Format: 'yy/dd/mm',
          startDate: today
        });
      } else if (strLocalStorage == 'Polls') {
        $('.form-section').append(renderQuestion);
        $('.textarea').after(renderAnswers);
      }
      $('.date-selector').on('changeDate', function (ev) {
        $(this).datepicker('hide');
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
        var siteImageURL = window.location.origin;;
        if ($('#cropped-img')[0].src == '../../../_catalogs/masterpage/BloomHomepage/images/prof-img.jpg') {
          alertify.set('notifier', 'position', 'top-right');
          alertify.error("Please Upload the Image File");
        }
        else if ($('#inputImage').length > 0) {
          $('#cropped-img').removeClass("crop-imagedisplay");
          $('.image-upload').css('width', '103px');
          $("#cropped-img").attr('src', '../../../_catalogs/masterpage/BloomHomepage/images/prof-img.jpg');
          $("#inputImage").val("");
        }
  
      });
      if ($('#inputImage').length > 0) {
  
        var canvas = $("#canvas"),
          context = canvas.get(0).getContext("2d"),
          $result = $('#cropped-img');
  
        $('#inputImage').change(function () {
  
          var iscropflag = true;
          var docname = $(this).val().split('.');
          docname = docname[docname.length - 1].toLowerCase();
          //$(this).attr("value", "");
          if ($.inArray(docname, ['jpg', 'jpeg', 'png', 'gif', 'bmp', 'tiff']) == -1) {
            //alertify.set('notifier', 'position', 'bottom-right');
            //alertify.error("Please Select Valid file Format");
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
                alertify.set('notifier', 'position', 'top-right');
                alertify.error("Invalid file type! Please select an image file.");
              }
            } else {
              alertify.set('notifier', 'position', 'top-right');
              alertify.error("No file(s) selected.");
            }
          }
        });
        $('#btnCrop').click(function () {
          // Get a string base 64 data url
          $result.empty();
          var croppedImageDataURL = canvas.cropper('getCroppedCanvas').toDataURL("image/png");
          $result.attr('class', 'crop-imagedisplay');
          //$('.image-upload').css('width', '42%');
          $result.attr('src', croppedImageDataURL);
          $('#canvasdisplay').css('display', 'none');
          canvas.cropper('reset');
          $result.empty();
          // $('#inputImage').val("");
        });
  
        $('#btnRestore').click(function () {
          canvas.cropper('reset');
          $result.empty();
          $result.attr('src', '../../../_catalogs/masterpage/BloomHomepage/images/prof-img.jpg');
          $('#canvasdisplay').css('display', 'none');
          $('#inputImage').val("");
        });
      }
  
      function InputChange() {
        var _this = $('#InputImage');
        var iscropflag = true;
        var docname = _this.val().split('.');
        docname = docname[docname.length - 1].toLowerCase();
        if ($.inArray(docname, ['jpg', 'jpeg', 'png', 'gif', 'bmp', 'tiff']) == -1) {
          //alertify.set('notifier', 'position', 'bottom-right');
          //alertify.error("Please Select Valid file Format");
          $("#inputImage").val("");
          iscropflag = false
        }
        if (iscropflag) {
          canvas.cropper('destroy');
          if (_this.files && _this.files[0]) {
            if (_this.files[0].type.match(/^image\//)) {
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
              alertify.set('notifier', 'position', 'top-right');
              alertify.error("Invalid file type! Please select an image file.");
            }
          } else {
            alertify.set('notifier', 'position', 'top-right');
            alertify.error("No file(s) selected.");
          }
        }
      }
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
