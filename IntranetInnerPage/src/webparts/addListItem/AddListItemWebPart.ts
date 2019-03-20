import {
  Version
} from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import 'jquery';
import * as Croppie from 'croppie';

//import styles from './AddListItemWebPart.module.scss';
import * as strings from 'AddListItemWebPartStrings';
require('bootstrap');
require('../../ExternalRef/js/jquery.richtext.js');
import pnp from 'sp-pnp-js';
import '../../ExternalRef/css/cropper.min.css'
import '../../ExternalRef/css/cropper.css';
import '../../ExternalRef/css/richtext.min.css';
import '../../ExternalRef/js/cropper-main.js';
import '../../ExternalRef/js/cropper.min.js';
import '../../ExternalRef/css/bootstrap-datepicker.min.css';
require('../../ExternalRef/js/bootstrap-datepicker.min.js');

import {
  readItems,
  addItems, checkUserinGroup,
  GetQueryStringParams,
  additemsimage,
  base64ToArrayBuffer
} from '../../commonService'
require('../../ExternalRef/js/alertify.min.js');
require('../../ExternalRef/js/jquery.richtext.js');
import '../../ExternalRef/css/bootstrap-datepicker.min.css';
import '../../ExternalRef/css/alertify.min.css';
import '../../ExternalRef/css/richtext.min.css';
require('../../ExternalRef/js/bootstrap-datepicker.min.js');
import '../../ExternalRef/css/cropper.min.css';
require('../../ExternalRef/js/cropper-main.js');
require('../../ExternalRef/js/cropper.min.js');

import {
  SPComponentLoader
} from '@microsoft/sp-loader';
export interface IAddListItemWebPartProps {
  description: string;
  Count: string;
}
declare var $;
declare var datepicker: any;
declare var alertify: any;
//var _that = this;
export default class AddListItemWebPart extends BaseClientSideWebPart<IAddListItemWebPartProps> {

  public render(): void {
    let IsAnonymous: boolean = false;
    SPComponentLoader.loadCss('https://cdnjs.cloudflare.com/ajax/libs/croppie/2.6.3/croppie.css');
    //Remove these once master page is applied
    // SPComponentLoader.loadScript("https://code.jquery.com/ui/1.12.1/jquery-ui.min.js");
    //  SPComponentLoader.loadCss("https://cdnjs.cloudflare.com/ajax/libs/jqueryui/1.12.1/jquery-ui.css");
    var _this = this;
    checkUserinGroup("Admin", this.context.pageContext.user.email, function (result) {

      if (result == 1) {
        _this.loadComponent(IsAnonymous);
      } else {
        alertify.alert('Access Denied', 'Sorry You dont have access to this page', function () {
          history.go(-1);
        }).set('closable', false);
      }
    });
  }
  public loadComponent(IsAnonymous) {

    var siteURL = this.context.pageContext.web.absoluteUrl;

    var strLocalStorage = GetQueryStringParams("CName");

    strLocalStorage = strLocalStorage.split("%20").join(' ');

    this.domElement.innerHTML =

      "<div class='breadcrumb'>" +
      "<ol>" +
      "<li><a href='" + siteURL + "/Pages/Home.aspx' title='Home'>Home</a></li>" +
      "<li><a class='pointer' id='breadTilte' title='" + strLocalStorage + "'>" + strLocalStorage + " List View</a></li>" +
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

    $('#my-image,#getcroppie,#cancel').hide();

    document.title = 'Add' + strLocalStorage;

    document.getElementById("ComponentName").innerHTML = GetQueryStringParams("CName").split('%20').join(" ");
    this.AddListItems();
    //For Radio Button
    $("input[name='selectionradio']").click(function () {
      var radioValue = $("input[name='selectionradio']:checked").val();
      if (radioValue == "Holiday") {
        $('.themelogo-upload').hide();
        $('#cropped-img').attr("src", siteURL + '/_catalogs/masterpage/BloomHomepage/images/prof-img.jpg');
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
    $("#anonymous").click(function () {
      IsAnonymous = $('#anonymous').prop('checked')//$('#test').is(':checked'));


    });


    if (strLocalStorage == "Holiday") {
      $("#radio-2").attr('checked', 'checked');
      $('.themelogo-upload').hide();
      $('#cropped-img').attr("src", siteURL + '/_catalogs/masterpage/BloomHomepage/images/prof-img.jpg');
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
    Addevent.addEventListener("click", (e: Event) => this.AddItem(siteURL, e, IsAnonymous));
    //}
    let breadTilte = document.getElementById('breadTilte');
    //for (let i = 0; i < Addevent.length; i++) {
    breadTilte.addEventListener("click", (e: Event) => this.pageBack());

    let Closeevent = document.getElementById('DelItem');
    // Addevent.addEventListener("click", (e: Event) => this.UpdateItem(siteURL, strComponentId));
    //for (let i = 0; i < Closeevent.length; i++) {
    Closeevent.addEventListener("click", (e: Event) => this.pageBack());
    // }

    this.datepickerkeyTypeBlocker();
    //$("#txtExpires").datepicker("setDate", '12/12/2018');

    let changeEvent = document.getElementById('uploadImageFile');
    if (changeEvent) {
      changeEvent.addEventListener("change", (e: Event) => this.validateFileType());
    }

    let videochangeEvent = document.getElementById('uploadVideoFile');
    if (videochangeEvent) {
      videochangeEvent.addEventListener("change", (e: Event) => this.validateVideoFileType());
    }
    $('#txtFolderName').keyup(function () {
      let array = [];
      $.ajax({
        url: siteURL + "/_api/Web/Lists/GetByTitle('" + strLocalStorage + "')/Items?$expand=ContentType&$select=LinkFilename,FileSystemObjectType",
        type: "GET",
        headers: {
          "accept": "application/json;odata=verbose",
        },
        success: function (data) {
          let ItemLength = data.d.results.length;
          for (let i = 0; i < ItemLength; i++) {
            if (data.d.results[i].FileSystemObjectType == 1) {
              array.push(data.d.results[i].LinkFilename)
            }
          }
          $('#txtFolderName').autocomplete({ source: array });
        },
        error: function (data) {
          console.log(data);
        },
      });
    });
    function readURL(input, width, height) {

      if (input.files && input.files[0] && input.files.length == 1) {

        $('.icon-camera').css("pointer-events", "none");
        $('#inputImage').css("pointer-events", "none");
        var reader = new FileReader();
        reader.onload = function (e: any) {

          $('#my-image').attr('src', e.target.result);
          var resize = new Croppie($('#my-image')[0], {

            //enableExif: true,
            viewport: { width: width, height: 265 },
            // boundary: { width: width + 200, height: height + 200 },
            boundary: {
              width: width + 100,
              height: 300
            },
            showZoomer: false,
            enableResize: true,
            enforceBoundary: false,
            enableOrientation: true

          });
          $('#getcroppie').fadeIn();
          $('#cancel').fadeIn();
          // $('#bannernote').fadeIn();
          $('#getcroppie').on('click', function () {
            resize.result({ type: 'base64' }).then(function (dataImg) {
              var data = [{ image: dataImg }];
              $('.cr-boundary').hide();
              $('#cropped-img').attr('src', dataImg);
              $('#getcroppie,#cancel').hide();
              $('.icon-camera').css("pointer-events", "");
              $('#inputImage').css("pointer-events", "");
            })
          })

          $('#cancel').on('click', function () {
            $('.croppie-container').hide();
            $('.cr-boundary').hide();
            $('#getcroppie,#cancel').hide();
            $('.icon-camera').css("pointer-events", "");
            $('#inputImage').css("pointer-events", "");
            $("#inputImage").val("");

          })
        }
        reader.readAsDataURL(input.files[0]);

      }
    }

    $("#inputImage").change(function (e) {
      if ($('.cr-boundary')[0]) {
        $('.cr-boundary')[0].remove();
      }
      if ($('#inputImage').length > 0) {
        $('.croppie-container').show();
        $('.custom-upload').css('pointer-events: none');
        let docname = $(this).val().split('.');
        docname = docname[docname.length - 1].toLowerCase();
        if ($.inArray(docname, ['jpg', 'jpeg', 'png', 'gif', 'bmp', 'tiff']) == -1) {
          $("#inputImage").val("");
          alertify.set('notifier', 'position', 'top-right');
          alertify.error("Invalid file type! Please select an image file.");
        } else {
          var _URL = window.URL;
          var file, img;
          if ((file = this.files[0])) {
            img = new Image();
            img.onload = function () {
              readURL($('#inputImage')[0], this.width, this.height);
            };
            img.src = _URL.createObjectURL(file);
          }
        }
      }
    });
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

  AddItem(siteURL, e, IsAnonymous) {
    var $body = $('body');
    if ($('.ajs-message').length > 0) {
      $('.ajs-message').remove();
    }
    e.preventDefault();
    $(this).prop('disabled', true);
    var dropOptionValue = $("input[name='selectionradio']:checked").val()
    var dropOptionImageValue = $("input[name='selectionradio']:checked").val()
    var $body = $("body");
    var strLocalStorage = GetQueryStringParams("CName");
    strLocalStorage = strLocalStorage.split("%20").join(' ');
    var isAllfield = true;
    var siteweburl = this.context.pageContext.web.serverRelativeUrl;

    //Add Announcements Part
    if (strLocalStorage == "Announcements") {
      var files = <HTMLInputElement>document.getElementById("inputImage");
      let file = files.files[0];
      if ($('#cropped-img')[0].src == siteURL + '/_catalogs/masterpage/BloomHomepage/images/prof-img.jpg') {
        alertify.set('notifier', 'position', 'top-right');
        alertify.error("Please Upload the Image File");
        isAllfield = false;
      }
      else if (!$('#txtExpires').val().trim()) {
        alertify.set('notifier', 'position', 'top-right');
        alertify.error("Please Select the Date");
        isAllfield = false;
      } else if (!$('#txtTitle').val().trim()) {
        alertify.set('notifier', 'position', 'top-right');
        alertify.error("Please Enter the Title");
        isAllfield = false;
      } else if (!$('.richText-editor').text().trim()) {
        alertify.set('notifier', 'position', 'top-right');
        alertify.error("Please Enter the Description");
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
            url: siteURL + "/_api/web/getfilebyserverrelativeurl('" + siteweburl + "/_catalogs/masterpage/Bloom/images/logo.png')/openbinarystream",
            type: "GET",
            success: function (data) {
              var name = Math.random().toString(36).substr(2, 9) + ".png";
              pnp.sp.web.getFolderByServerRelativeUrl("Announcements").files.add(name, data, true)
                .then(function (result) {
                  result.file.listItemAllFields.get().then((listItemAllFields) => {
                    pnp.sp.web.lists.getByTitle("Announcements").items.getById(listItemAllFields.Id).update({
                      Title: $("#txtTitle").val().trim(),
                      Expires: new Date($('#txtExpires').val()),
                      Explanation: $('.richText-editor').html(),
                      ExplanationText: $('.richText-editor').text(),
                      Image: {
                        "__metadata": {
                          "type": "SP.FieldUrlValue"
                        },
                        Url: siteURL + "/_catalogs/masterpage/Bloom/images/logo.png"
                      }
                    }).then(r => {
                      $body.removeClass("loading");
                      $('.addbutton').prop('disabled', true);
                      //this.pageBack();
                      window.history.back();
                      //window.location.href = "" + siteURL + "/Pages/ListView.aspx?CName=" + strLocalStorage + "";
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
          //  var file1 = $('#cropped-img').attr('src').split("base64,");   $('#croppieImge').attr('src', dataImg);
          var file1 = $('#cropped-img').attr('src').split("base64,");
          var blob = base64ToArrayBuffer(file1[1]);
          let myobjQl = {
            Title: $('#txtTitle').val().trim(),
            Expires: new Date($('#txtExpires').val()),
            Explanation: $('.richText-editor').html(),
            ExplanationText: $('.richText-editor').text(),
            Image: {
              "__metadata": {
                "type": "SP.FieldUrlValue"
              },
              Url: siteURL + "/" + strLocalStorage + "/" + uniquename
            }
          }

          additemsimage(strLocalStorage, uniquename, blob, myobjQl, function (e) {
            $body.addClass("loading");
            if (e.data) {
              $body.removeClass("loading");
              $('.addbutton').prop('disabled', true);
              window.history.back();
              //   window.location.href = "" + siteURL + "/Pages/ListView.aspx?CName=" + strLocalStorage + "";

            } else {
              $body.removeClass("loading");
              console.log(e);
            }
          });
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
        addItems("Holiday", myobjHol, function (e) {
          $body.addClass("loading");
          if (e.data) {
            $body.removeClass("loading");
            $('.addbutton').prop('disabled', true);
            // this.pageBack();
            window.history.back();
            // window.location.href = "" + siteURL + "/Pages/ListView.aspx?CName=" + strLocalStorage + "";

          } else {
            $body.removeClass("loading");
            console.log(e);
          }
        });

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
        addItems("Quick Links", myobjQl, function (e) {
          $body.addClass("loading");
          if (e.data) {
            $body.removeClass("loading");
            $('.addbutton').prop('disabled', true);
            // this.pageBack();
            window.history.back();
            // window.location.href = "" + siteURL + "/Pages/ListView.aspx?CName=" + strLocalStorage + "";

          } else {
            $body.removeClass("loading");
            console.log(e);
          }
        });
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
            Url: siteURL + "/" + strLocalStorage + "/" + uniquename
          }
        }
        additemsimage(strLocalStorage, uniquename, blob, myobjQl, function (e) {
          if (e.data) {
            $body.removeClass("loading");
            $('.addbutton').prop('disabled', true);
            // this.pageBack();
            window.history.back();
            //  window.location.href = "" + siteURL + "/Pages/ListView.aspx?CName=" + strLocalStorage + "";

          } else {
            console.log(e);
          }
        });
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
          Url: siteURL + "/" + strLocalStorage + "/" + file.name
        }
        //,Display: true
      }
      if (isAllfield) {

        $body.addClass("loading");
        additemsimage(strLocalStorage, file.name, file, myobjQl, function (e) {
          if (e.data) {
            $body.removeClass("loading");
            $('.addbutton').prop('disabled', true);
            // this.pageBack();
            window.history.back();
            // window.location.href = "" + siteURL + "/Pages/ListView.aspx?CName=" + strLocalStorage + "";
          } else {
            console.log(e);
          }
        });
      }

    }       ////Add Banners Part
    else if (strLocalStorage == "Banners") {
      var regexp = /^(https?|s?ftp):\/\/(((([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(%[\da-f]{2})|[!\$&'\(\)\*\+,;=]|:)*@)?(((\d|[1-9]\d|1\d\d|2[0-4]\d|25[0-5])\.(\d|[1-9]\d|1\d\d|2[0-4]\d|25[0-5])\.(\d|[1-9]\d|1\d\d|2[0-4]\d|25[0-5])\.(\d|[1-9]\d|1\d\d|2[0-4]\d|25[0-5]))|((([a-z]|\d|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(([a-z]|\d|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])*([a-z]|\d|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])))\.)+(([a-z]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(([a-z]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])*([a-z]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])))\.?)(:\d*)?)(\/((([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(%[\da-f]{2})|[!\$&'\(\)\*\+,;=]|:|@)+(\/(([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(%[\da-f]{2})|[!\$&'\(\)\*\+,;=]|:|@)*)*)?)?(\?((([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(%[\da-f]{2})|[!\$&'\(\)\*\+,;=]|:|@)|[\uE000-\uF8FF]|\/|\?)*)?(#((([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(%[\da-f]{2})|[!\$&'\(\)\*\+,;=]|:|@)|\/|\?)*)?$/i
      var files = <HTMLInputElement>document.getElementById("inputImage");
      let file = files.files[0];
      var siteImageURL = window.location.origin;
      if ($('#cropped-img')[0].src == siteURL + '/_catalogs/masterpage/BloomHomepage/images/prof-img.jpg') {
        alertify.set('notifier', 'position', 'top-right');
        alertify.error("Please Select the Image");
        isAllfield = false;
      } else if (!$('#txtTitle').val().trim()) {
        alertify.set('notifier', 'position', 'top-right');
        alertify.error("Please Enter the Title");
        isAllfield = false;
      }
      else if (!$('.richText-editor').text().trim()) {
        alertify.set('notifier', 'position', 'top-right');
        alertify.error("Please Enter the Description");
        isAllfield = false;
      }
      else if ($('#txtHyper').val() && !regexp.test($('#txtHyper').val().trim())) {
        alertify.set('notifier', 'position', 'top-right');
        alertify.error("Please give a valid Link URL");
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

      let myobjQl = {
        Title: $("#txtTitle").val().trim(),
        // BannerContent: $("#txtrequiredDescription").val(),
        BannerContent: $('.richText-editor').html(),
        Image: {
          "__metadata": {
            "type": "SP.FieldUrlValue"
          },
          Url: siteURL + "/" + strLocalStorage + "/" + uniquename
        },
        LinkURL: {
          "__metadata": {
            "type": "SP.FieldUrlValue"
          },
          Url: $('#txtHyper').val(),
        }
      }
      if (isAllfield) {
        $body.addClass("loading");
        additemsimage(strLocalStorage, uniquename, blob, myobjQl, function (e) {
          if (e.data) {
            $body.removeClass("loading");
            $('.addbutton').prop('disabled', true);
            //   this.pageBack();
            window.history.back();
          } else {
            console.log(e);
          }
        });
      }
    }     ////Add New ImageGallery Part

    else if (strLocalStorage == "Image Gallery") {

      var regexp = /^(https?|s?ftp):\/\/(((([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(%[\da-f]{2})|[!\$&'\(\)\*\+,;=]|:)*@)?(((\d|[1-9]\d|1\d\d|2[0-4]\d|25[0-5])\.(\d|[1-9]\d|1\d\d|2[0-4]\d|25[0-5])\.(\d|[1-9]\d|1\d\d|2[0-4]\d|25[0-5])\.(\d|[1-9]\d|1\d\d|2[0-4]\d|25[0-5]))|((([a-z]|\d|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(([a-z]|\d|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])*([a-z]|\d|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])))\.)+(([a-z]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(([a-z]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])*([a-z]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])))\.?)(:\d*)?)(\/((([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(%[\da-f]{2})|[!\$&'\(\)\*\+,;=]|:|@)+(\/(([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(%[\da-f]{2})|[!\$&'\(\)\*\+,;=]|:|@)*)*)?)?(\?((([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(%[\da-f]{2})|[!\$&'\(\)\*\+,;=]|:|@)|[\uE000-\uF8FF]|\/|\?)*)?(#((([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(%[\da-f]{2})|[!\$&'\(\)\*\+,;=]|:|@)|\/|\?)*)?$/i
      var files = <HTMLInputElement>document.getElementById("uploadImageFile");
      let file = files.files[0];
      if (!$('#txtFolderName').val().trim()) {
        alertify.set('notifier', 'position', 'top-right');
        alertify.error("Please Enter the Folder Name");
        isAllfield = false;
      }
      else if (!$('#uploadImageFile').val().trim()) {
        alertify.set('notifier', 'position', 'top-right');
        alertify.error("Please Select the Image");
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
              })

          })

      }
    }
    else if (strLocalStorage == "Video Gallery") {

      var regexp = /^(https?|s?ftp):\/\/(((([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(%[\da-f]{2})|[!\$&'\(\)\*\+,;=]|:)*@)?(((\d|[1-9]\d|1\d\d|2[0-4]\d|25[0-5])\.(\d|[1-9]\d|1\d\d|2[0-4]\d|25[0-5])\.(\d|[1-9]\d|1\d\d|2[0-4]\d|25[0-5])\.(\d|[1-9]\d|1\d\d|2[0-4]\d|25[0-5]))|((([a-z]|\d|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(([a-z]|\d|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])*([a-z]|\d|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])))\.)+(([a-z]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(([a-z]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])*([a-z]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])))\.?)(:\d*)?)(\/((([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(%[\da-f]{2})|[!\$&'\(\)\*\+,;=]|:|@)+(\/(([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(%[\da-f]{2})|[!\$&'\(\)\*\+,;=]|:|@)*)*)?)?(\?((([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(%[\da-f]{2})|[!\$&'\(\)\*\+,;=]|:|@)|[\uE000-\uF8FF]|\/|\?)*)?(#((([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(%[\da-f]{2})|[!\$&'\(\)\*\+,;=]|:|@)|\/|\?)*)?$/i
      var files = <HTMLInputElement>document.getElementById("uploadVideoFile");
      let file = files.files[0];
      var filename = $("#uploadVideoFile").val().split(String.fromCharCode(92));
      $("#browsedVideofile").val(filename[filename.length - 1]);
      var radioValue = $("input[name='selectionradioImage']:checked").val();
      if (radioValue == "Upload") {
        if (!$('#txtFolderName').val().trim()) {
          alertify.set('notifier', 'position', 'top-right');
          alertify.error("Please enter the Folder name");
          isAllfield = false;
        }
        else if (!$('#txtTitle').val().trim()) {
          alertify.set('notifier', 'position', 'top-right');
          alertify.error("Please enter the Title");
          isAllfield = false;
        }
        if (!$('#uploadVideoFile').val().trim()) {
          alertify.set('notifier', 'position', 'top-right');
          alertify.error("Please Select the video File");
          isAllfield = false;
        }
      }
      if (radioValue == "Stream") {
        if (!$('#txtFolderName').val().trim()) {
          alertify.set('notifier', 'position', 'top-right');
          alertify.error("Please enter the Folder name");
          isAllfield = false;
        }
        else if (!$('#txtTitle').val().trim()) {
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

            if (radioValue == "Upload") {
              var VideoTitle = { Title: $('#txtTitle').val().trim() };
              pnp.sp.web.getFolderByServerRelativeUrl("Video Gallery" + "/" + $('#txtFolderName').val()).files.add(file.name, file, true)
                .then(({ file }) => file.getItem())
                .then(item => item.update(VideoTitle))
                .then(function (result: any) {
                  $body.removeClass("loading");
                  window.history.back();
                });
            }
            else {
              var Videojson = {
                Title: $("#txtTitle").val(),
                LinkURL: {
                  "__metadata": { "type": "SP.FieldUrlValue" },
                  Url: $('#txtHyper').val().trim()
                }
              };

              if ($("#uploadImageFile").val() == undefined || $("#uploadImageFile").val() == null || $("#uploadImageFile").val() == '') {
                $.ajax({
                  url: this.context.pageContext.site.absoluteUrl + "/_api/web/getfilebyserverrelativeurl('/sites/spuat/_catalogs/masterpage/Bloom/images/logo.png')/openbinarystream",
                  type: "GET",
                  success: function (data) {
                    var name = $("#txtTitle").val().trim() + '.jpg';
                    pnp.sp.web.getFolderByServerRelativeUrl("Video Gallery" + "/" + $('#txtFolderName').val()).files.add(name, data, true)
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
        alertify.error("Please Enter the Title");
        isAllfield = false;
      } else if ($('#ddlDepartment').val() == "Select") {
        //$('#uploadFile').focus();
        alertify.set('notifier', 'position', 'top-right');
        alertify.error("Please Select the Department");
        isAllfield = false;
      } else if (!$('.richText-editor').text().trim()) {
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
        Explanation: $('.richText-editor').html(),
        DocumentFile: {
          "__metadata": {
            "type": "SP.FieldUrlValue"
          },
          Url: siteURL + "/" + strLocalStorage + "/" + file.name
        }
      }
      if (isAllfield) {

        $body.addClass("loading");
        var _thiss = this;
        additemsimage(strLocalStorage, file.name, file, myobjQl, function (e) {
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

          } else {
            console.log(e);
          }
        });
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
                  Url: siteURL + "/_catalogs/masterpage/Bloom/images/logo.png"
                },
                DocumentFile: {
                  "__metadata": {
                    "type": "SP.FieldUrlValue"
                  },
                  Url: fileURL + datefile.data.ServerRelativeUrl
                }
              }
              $body.addClass("loading");
              additemsimage(strLocalStorage, name, docfile, CorDocOnly, function (e) {
                if (e.data) {
                  $body.removeClass("loading");
                  $('.addbutton').prop('disabled', true);
                  window.history.back();
                } else {
                  console.log(e);
                }
              });
            });
        }
        else if (docfiles.files.length == 0 && files.files.length > 0) {
          var file1 = $('#cropped-img').attr('src').split("base64,");
          var blob = base64ToArrayBuffer(file1[1]);
          var uniquename = Math.random().toString(36).substr(2, 9) + ".png";
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
              Url: siteURL + "/" + strLocalStorage + "/" + uniquename
            },

          }
          $body.addClass("loading");
          additemsimage(strLocalStorage, uniquename, blob, myobjQl, function (e) {
            if (e.data) {
              $body.removeClass("loading");
              $('.addbutton').prop('disabled', true);

              window.history.back();


            } else {
              console.log(e);
            }
          });

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
              Url: siteURL + "/_catalogs/masterpage/Bloom/images/logo.png"
            },
          }
          $body.addClass("loading");
          additemsimage(strLocalStorage, uniquename, blob, myobjQl, function (e) {
            if (e.data) {
              $body.removeClass("loading");
              $('.addbutton').prop('disabled', true);
              window.history.back();
            } else {
              console.log(e);
            }
          });

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
      if ($('#inputImage').length > 0 && $('#cropped-img')[0].src == siteURL + "/_catalogs/masterpage/BloomHomepage/images/prof-img.jpg") {
        alertify.set('notifier', 'position', 'top-right');
        alertify.error("Please Crop the Image First");
        isAllfield = false;
      }
      if (!$('#txtTitle').val().trim()) {
        //$('#txtTitle').focus();
        alertify.set('notifier', 'position', 'top-right');
        alertify.error("Please Enter Title");
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
                  Explanation: $('.richText-editor').html(),
                  ExplanationText: $('.richText-editor').text(),
                  Image: {
                    "__metadata": {
                      "type": "SP.FieldUrlValue"
                    },
                    Url: siteURL + "/" + "Events" + "/" + name
                  }
                }).then(r => {
                  $body.removeClass("loading");
                  $('.addbutton').prop('disabled', true);
                  // this.pageBack();
                  window.history.back();
                  //   window.location.href = "" + siteURL + "/Pages/ListView.aspx?CName=" + strLocalStorage + "";
                  //console.log(name + " properties updated successfully!");
                });
              });
            })

        } else {
          if ($("#inputImage").val() == undefined || $("#inputImage").val() == null || $("#inputImage").val() == '') {
            $.ajax({
              url: siteURL + "/_api/web/getfilebyserverrelativeurl('" + siteweburl + "/_catalogs/masterpage/Bloom/images/logo.png')/openbinarystream",
              type: "GET",
              success: function (data) {
                var name = Math.random().toString(36).substr(2, 9) + ".png";

                pnp.sp.web.getFolderByServerRelativeUrl("Events").files.add(name, data, true)
                  .then(function (result) {
                    result.file.listItemAllFields.get().then((listItemAllFields) => {
                      pnp.sp.web.lists.getByTitle("Events").items.getById(listItemAllFields.Id).update({
                        Title: $("#txtTitle").val().trim(),
                        StartDate: new Date($('#txtStartDate').val()),
                        EndDate: new Date($('#txtEndDate').val()),
                        Explanation: $('.richText-editor').html(),
                        ExplanationText: $('.richText-editor').text(),
                        Image: {
                          "__metadata": {
                            "type": "SP.FieldUrlValue"
                          },
                          Url: siteURL + "/_catalogs/masterpage/Bloom/images/logo.png"
                        }
                      }).then(r => {
                        $body.removeClass("loading");
                        $('.addbutton').prop('disabled', true);
                        window.history.back();
                      });
                    });
                  })
              },
              error: function (data) {
                $body.removeClass("loading");
                console.log(data);
              },

            });
          }

        }
      }

    }    //Add Polls Part
    else if (strLocalStorage == "Polls") {
      var optionseperate = $('#txtOptions').val();
      var resultarray = optionseperate.split(";");
      var newArray = resultarray.filter(function (v) {
        return v !== '' && v !== ' '
      });

      if (!$('#txtQuestion').val().trim()) {
        //$('#txtQuestion').focus();
        alertify.set('notifier', 'position', 'top-right');
        alertify.error("Please Enter Question");
        isAllfield = false;
      } else if (!$('#txtOptions').val().trim() || newArray.length < 2) {
        // $('#txtOptions').focus();
        alertify.set('notifier', 'position', 'top-right');
        alertify.error("Minimum Two options needed");
        isAllfield = false;
      } else if (!$('#txtOptions').val().trim() || newArray.length > 4) {
        alertify.set('notifier', 'position', 'top-right');
        alertify.error("Maximum Four Answers only Allowed");
        return false;
      }
      let myobjPols = {
        Question: $("#txtQuestion").val().trim(),
        Options: $("#txtOptions").val().trim(),
        IsVisibles: IsAnonymous

      }
      if (isAllfield) {
        $body.addClass("loading");
        addItems("Polls", myobjPols, function (e) {
          $body.addClass("loading");
          if (e.data) {
            $body.removeClass("loading");
            $('.addbutton').prop('disabled', true);
            // this.pageBack();
            window.history.back();
            // window.location.href = "" + siteURL + "/Pages/ListView.aspx?CName=" + strLocalStorage + "";

          } else {
            $body.removeClass("loading");
            console.log(e);
          }
        });

      }
    }
  }
  async AddDepartments() {
    let myobjPols = {
      Departments: $("#txtDepartment").val().trim()
    }
    await addItems("Departments", myobjPols, function (e) {
      if (e.data) {
        window.history.back();
      } else {
        console.log(e);
      }

    });
  }

  //// Bind Organizational Policies Data to Dropdown
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


  //// Bind ImageGallery folder name to autocomplete
  public bindgalleryImage() {
    let isSearch: boolean = true;
    let siteURL = this.context.pageContext.web.absoluteUrl;
    let _that = this;
    $("#txtTitle").keyup(function (event) {
      var get = $('#txtTitle').val();
      if (!get) {
        alertify.set('notifier', 'position', 'top-right');
        alertify.error("Please Enter the Topic");
        isSearch = false;
      } if (isSearch) {
        _that.bindSearchTitle(siteURL, isSearch)
      }
    });
  }

  async bindSearchTitle(siteURL, isSearch) {
    if (isSearch) {
      let columnArray: any = ["ID", "FileLeafRef", "FileSystemObjectType", "FileDirRef"];
      let picItems = await readItems("Image Gallery", columnArray, 5000, "ID");
      picItems.then((items: any[]) => {
        let itemLength = picItems.length;
        for (var i = 0; i < itemLength; i++) {
          if (picItems[i].FileSystemObjectType == 1) {
          }
        }
      });
    }
  }

  public validateVideoFileType() {
    var fileName = $("#uploadVideoFile").val();
    var idxDot = fileName.lastIndexOf(".") + 1;
    var extFile = fileName.substr(idxDot, fileName.length).toLowerCase();
    if (extFile == "mp4" || extFile == "x-m4v" || extFile == "m4a" || extFile == "f4v" || extFile == "m4b" || extFile == "mov" || extFile == "f4b" || extFile == "flv") {
      var filename = $("#uploadVideoFile").val().split(String.fromCharCode(92));
      $("#browsedfileName").val(filename[filename.length - 1]);

    } else {
      $("#uploadVideoFile").val(null);
      $("#browsedfileName").val('');
    }
    var filename = $("#uploadVideoFile").val().split(String.fromCharCode(92));
    $("#browsedVideofile").val(filename[filename.length - 1]);

  }
  public validateFileType() {
    var fileName = $("#uploadImageFile").val();
    var idxDot = fileName.lastIndexOf(".") + 1;
    var extFile = fileName.substr(idxDot, fileName.length).toLowerCase();
    if (extFile == "jpg" || extFile == "jpeg" || extFile == "png" || extFile == "gif" || extFile == "bmp" || extFile == "tiff") {
      var filename = $("#uploadImageFile").val().split(String.fromCharCode(92));
      $("#browsedfileName").val(filename[filename.length - 1]);
      //TO DO
    } else {
      $("#uploadImageFile").val(null);
      $("#browsedfileName").val('');
    }
  }
  ////Add Items in to List
  public AddListItems() {
    var strLocalStorage = GetQueryStringParams("CName");
    strLocalStorage = strLocalStorage.split("%20").join(' ');
    var renderUploadImagefile = "";
    var renderUploadVideofile = "";
    var rendertext = "";
    var renderFoldertext = "";
    var renderdate = "";
    var renderHyperlink = "";
    var renderHyperSitelink = "";
    var renderUploadfile = "";
    var renderSiteLink = "";
    var renderStartEndDate = "";
    var renderhtmlImageBanners = "";
    var renderQuestion = "";
    var renderAnswers = "";
    var renderDepartment = "";
    var renderRequiredDescription = "";
    var requirednewrichTextEditor = "";
    var newrichTextEditor = "";
    var renderCorpUploadfile = "";
    var renderOptionImageGallery = "";
    var renderOption = "";
    var renderDepartmentddl = "";
    var renderhtmlCheckBox = "";
    var strLocalStorage = GetQueryStringParams("CName");
    strLocalStorage = strLocalStorage.split("%20").join(' ');

    var radioValue = $("input[name='selectionradio']:checked").val();
    if (radioValue == "Holiday") {
      strLocalStorage = "Holiday";

    } else if (radioValue == "Events") {
      strLocalStorage = "Events";
    }
    var siteURL = this.context.pageContext.web.absoluteUrl;

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


    renderhtmlImageBanners += "<div class='form-imgsec'>" +
      "<div class='themelogo-upload'>" +
      "<label id='imageLabel' class='control-label'>Image</label>" +
      "<img id='cropped-img' src=" + siteURL + "/_catalogs/masterpage/BloomHomepage/images/prof-img.jpg>" +
      "<div class='image-upload' id='imagerestrict'>" +
      "<div class='custom-upload'>" +
      "<input type='file' id='inputImage' name='file' accept='.jpg,.jpeg,.png,.gif,.bmp,.tiff' multiple='' class='file' />" +
      "<button class='btn btn-primary' type='button'><i class='icon-camera'></i></button>" +
      "<img id='result' src=''>" +
      "</div>" +
      "<a href='#' title='Delete' id='image-delete'>" +
      "<i class='icon-delete'></i></a>" +
      "</div>" +
      "<img id='my-image' src='#' />" +
      "</div>" +
      "<div class='crop-button col-md-12'>" +
      "<button id='getcroppie' type='button'>Crop Image</button>" +
      "<button id='cancel' type='button'>Cancel</button>" +
      "</div>" +
      "</div>";

    renderhtmlCheckBox += "<div class='check-box anonymous'>" +
      "<input id='anonymous' type='checkbox' name='' value=''>" +
      "<label>Is Anonymous</label>" +
      "</div>";

    rendertext += "<div class='input text'>" +
      "<label class='control-label'>Title</label>" +
      "<input class='form-control' type='text' value=''  id='txtTitle' /></div>";

    renderFoldertext += "<div id='foldername' class='input text'>" +
      "<label class='control-label'>Folder Name</label>" +
      "<input class='form-control' type='text' value='' id='txtFolderName' /></div>";

    renderdate += "<div class='input date'>" + "<i class='icon-calenter'></i>" +
      "<label class='control-label'>Date</label>" +
      "<input class='form-control date-selector' type='text' value='' id='txtExpires' /></div>";

    newrichTextEditor += "<div class='textarea input'>" +
      "<label>Description</label>" +
      "<textarea id='txtDescription' class='form-control content'></textarea>" +
      "</div>";

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
      "<input type='file' id='uploadImageFile' name='image' accept='.jpg,.jpeg,.png,.gif,.bmp,.tiff' multiple=''   class='file'>" +
      "<div class='input-group'>" +
      "<span class='input-group-btn input-group-sm'>" +
      "<button type='button' class='btn btn-fab btn-fab-mini'>Browse</button>" +
      "</span>" +
      "<input type='text' id='browsedfileName' readonly='' class='form-control' placeholder='Upload Files'>" +
      "</div>" +
      "</div>" +
      "</div>" +
      "</div>";

    renderUploadVideofile += "<div class='form-imgsec'>" +
      "<div class='themelogo-upload' style='display: block;'>" +
      "<div class='custom-upload banner-upload'>" +
      "<label class='control-label'>Upload Video File</label>" +
      "<input type='file' id='uploadVideoFile' name='file' accept='video/mp4,video/x-m4v,video/*' multiple='' class='file'>" +
      "<div class='input-group'>" +
      "<span class='input-group-btn input-group-sm'>" +
      "<button type='button' class='btn btn-fab btn-fab-mini'>Browse</button>" +
      "</span>" +
      "<input type='text' id='browsedVideofile' readonly='' class='form-control' placeholder='Upload Files'>" +
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
    renderAnswers += "<div id='answers' class='input text'>" +
      "<i class=''></i>" +
      "<label class='control-label'>Options</label>" +
      "<input class='form-control' type='text' value='' id='txtOptions'/>" +
      "<span>Please Enter more than one Options with semicolon ( ; ) Maximun four options</span>" +
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

    requirednewrichTextEditor += "<div class='textarea input'>" +
      "<label class='control-label'>Description</label>" +
      "<textarea id='txtrequiredDescription' class='form-control content'></textarea>" +
      "</div>";

    renderRequiredDescription += "<div id='rrdescription' class='input textarea'><label class='control-label'>Description</label><textarea class='form-control' id='txtrequiredDescription'></textarea></div>";

    var date = new Date();
    var today = new Date(date.getFullYear(), date.getMonth(), date.getDate());
    if (strLocalStorage == 'Announcements') {
      $('.form-section').append(renderhtmlImageBanners);
      $('#my-image,#getcroppie,#cancel').hide();
      $('.form-imgsec').after(renderdate);
      $('.date').after(rendertext);
      $('.text').after(requirednewrichTextEditor);
      $('#txtExpires').datepicker({
        format: "mm/dd/yyyy",
        startDate: today

      }).datepicker('setDate', 'new Date()');
    } else if (strLocalStorage == 'Holiday') {
      $('.form-section').append(renderOption);
      $('.appendOption').after(renderhtmlImageBanners);
      $('#my-image,#getcroppie,#cancel').hide();
      $('.form-imgsec').after(rendertext)
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



    } else if (strLocalStorage == 'Quick Links') {

      $('.form-section').append(rendertext);
      $('.text').after(renderHyperlink);

    } else if (strLocalStorage == 'Employee Corner') {
      $('.form-section').append(rendertext);
      $('.text').after(renderUploadfile);

    }

    else if (strLocalStorage == 'Image Gallery') {
      $('.form-section').append(renderFoldertext);
      $('.text').after(renderUploadImagefile);
      this.bindgalleryImage();

    }
    else if (strLocalStorage == 'Video Gallery') {
      $('.form-section').append(renderOptionImageGallery + renderFoldertext + rendertext + renderUploadVideofile + renderHyperlink);
      $('.banner-upload').show();
      $('#divHyperLink').hide();
    }
    else if (strLocalStorage == 'Organizational Policies') {
      $('.form-section').append(rendertext);
      $('.text').after(renderDepartmentddl + renderDepartment + renderUploadfile + newrichTextEditor);
      this.bindorgDept()
    } else if (strLocalStorage == 'Banners') {
      $('.form-section').append(renderhtmlImageBanners);
      $('#my-image,#getcroppie,#cancel').hide();
      $('.form-imgsec').after(rendertext);
      $('.text').after(requirednewrichTextEditor);
      $('.textarea').after(renderHyperSitelink);

    } else if (strLocalStorage == 'Corporate Discounts') {
      $('.form-section').append(renderhtmlImageBanners);
      $('#my-image,#getcroppie,#cancel').hide();
      $('.form-imgsec').after(rendertext);
      $('.text').after(renderSiteLink);
      $('#siteLink').after(renderCorpUploadfile);

    } else if (strLocalStorage == 'Events') {
      $('.form-section').append(renderOption);
      $('.appendOption').after(renderhtmlImageBanners);
      $('#my-image,#getcroppie,#cancel').hide();
      $('#imageLabel').removeClass('control-label');
      $('.form-imgsec').after(rendertext)
      $('.text').after(newrichTextEditor);
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
      $('#answers').after(renderhtmlCheckBox);
    }
    $('.content').richText();
    $('.date-selector').on('changeDate', function (ev) {
      $(this).datepicker('hide');
    });
    if ($('#uploadFile').length > 0) {
      $(document).on('change', '#uploadFile', function () {
        var docname = $(this).val().split('.');
        docname = docname[docname.length - 1].toLowerCase();
        if ($.inArray(docname, ['doc', 'docx', 'xls', 'csv', 'ppt', 'pdf']) == -1) {
          alertify.set('notifier', 'position', 'top-right');
          alertify.error("Please Select Valid File Format");
          $("#uploadFile").val("");
        } else {
          $(this).parent('.custom-upload').find('.form-control').val($(this).val().replace(/C:\\fakepath\\/i, ''));
        }
      });
    }
    $('#image-delete').click(function () {
      var siteImageURL = window.location.origin;;
      if ($('#cropped-img')[0].src == siteURL + '/_catalogs/masterpage/BloomHomepage/images/prof-img.jpg') {
        alertify.set('notifier', 'position', 'top-right');
        alertify.error("Please Upload the Image File");
      } else {
        $('#cropped-img')[0].src = siteURL + '/_catalogs/masterpage/BloomHomepage/images/prof-img.jpg';
        $('#inputImage').val("");
      }
    });

  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [{
        header: {
          description: strings.PropertyPaneDescription
        },
        groups: [{
          groupName: strings.BasicGroupName,
          groupFields: [
            PropertyPaneTextField('description', {
              label: strings.DescriptionFieldLabel
            })
          ]
        }]
      }]
    };
  }
}
