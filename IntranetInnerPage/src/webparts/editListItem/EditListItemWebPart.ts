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

import 'jquery';
import * as Croppie from 'croppie';
//import styles from './EditListItemWebPart.module.scss';
import * as strings from 'EditListItemWebPartStrings';
import {
    SPComponentLoader
} from '@microsoft/sp-loader';

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
    updateitems, checkUserinGroup,
    GetQueryStringParams,
    base64ToArrayBuffer
} from '../../commonService';

export interface IEditListItemWebPartProps {
    description: string;
}
declare var $;
declare var alertify: any;

export default class EditListItemWebPart extends BaseClientSideWebPart<IEditListItemWebPartProps> {
    strcropstorage = "";
    imageValue = 0;
    imgsrc;
    qustion_lenght;
    answer_lenght;
    siteURL = "";
    public render(): void {
        SPComponentLoader.loadCss('https://cdnjs.cloudflare.com/ajax/libs/croppie/2.6.3/croppie.css');
        var _this = this;
        checkUserinGroup("Admin", this.context.pageContext.user.email, function (result) {
            if (result == 1) {
                _this.loadEditComponent();
            } else {
                alertify.alert('Access Denied', 'Sorry You dont have access to this page', function () {
                    history.go(-1);
                }).set('closable', false);
            }
        });
    }
    public loadEditComponent() {
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
            "<h2 id='ComponentName'></h2>" +
            "</div>" +
            "<div class='form-section required'>" +
            "</div>" +
            "<div class='modal'><!-- Place at bottom of page --></div>";

        document.title = 'Edit' + strLocalStorage;
        document.getElementById("ComponentName").innerHTML = GetQueryStringParams("CName").split("%20").join(" ");

        var strComponentId = GetQueryStringParams("CID");
        this.renderhtml(strComponentId);
        let Addevent = document.getElementById('UpdateItem');

        Addevent.addEventListener("click", (e: Event) => this.UpdateItem(this.siteURL, strLocalStorage, strComponentId));

        let breadTilte = document.getElementById('breadTilte');

        breadTilte.addEventListener("click", (e: Event) => this.pageBack());

        let Closeevent = document.getElementById('DelItem');

        Closeevent.addEventListener("click", (e: Event) => this.pageBack());

        this.datepickerkeyTypeBlocker();
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
        
        $("#inputImage").change(function () {
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
    pageBack() {
        window.history.back();
    }
    datepickerkeyTypeBlocker() {
        $("#txtExpires,#txtStartDate,#txtEndDate,#txtEvDate,#txtEEDate").keypress(
            function (event) {
                event.preventDefault();
            });
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

        if ($('#cropped-img')[0].src == this.siteURL + '/_catalogs/masterpage/BloomHomepage/images/prof-img.jpg') {
            alertify.set('notifier', 'position', 'top-right');
            alertify.error("Please Upload the Image File");
            return false;
        }
        else if (!$('#txtExpires').val()) {
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
            alertify.error("Please Enter the Title");
            return false;
        }
        else if (!$('#txtStartDate').val()) {
            alertify.set('notifier', 'position', 'top-right');
            alertify.error("Please Select the Start Date");
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
        } else if (!$('#txtTitle').val().trim()) {
            alertify.set('notifier', 'position', 'top-right');
            alertify.error("Please Enter the Title");
            return false;

        }
        else if (!$('#txtrequiredDescription').val().trim()) {
            alertify.set('notifier', 'position', 'top-right');
            alertify.error("Please Enter the Description");
            return false;

        }
        return true;
    }
    pollsValidation() {
        var optionseperate = $('#txtOptions').val();
        var resultarray = optionseperate.split(";");
        var newArray = resultarray.filter(function (v) {
            return v !== ' ' && v !== ''
        });
        if (!$('#txtQuestion').val().trim()) {
            alertify.set('notifier', 'position', 'top-right');
            alertify.error("Please Enter the Question");
            return false;
        } else if (!$('#txtOptions').val().trim() || newArray.length < 2) {
            alertify.set('notifier', 'position', 'top-right');
            alertify.error("Please Enter the Answers Correctly");
            return false;
        }
        else if (!$('#txtOptions').val().trim() || newArray.length > 4) {
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

    UpdateItem(siteURL, strLocalStorage, strComponentId) {
        var $body = $('body');
        if ($('.ajs-message').length > 0) {
            $('.ajs-message').remove();
        }
        var that = this;
        let strcrop = localStorage.getItem("crop");
        var count;
        let objResults;
        var $body = $("body");
        var isAllfield = true;

        if (strLocalStorage == "Announcements") {

            var files = <HTMLInputElement>document.getElementById("inputImage");
            let file = files.files[0];

            if (strcrop == "1" && files.files.length == 0) {
                var saveData = {
                    Title: $("#txtTitle").val(),
                    Explanation: $('.richText-editor').html(),
                    ExplanationText: $('.richText-editor').text(),
                    Expires: new Date($('#txtExpires').val()),
                    Image: {
                        "__metadata": {
                            "type": "SP.FieldUrlValue"
                        },
                        Url: siteURL + "/_catalogs/masterpage/Bloom/images/logo.png"
                    }

                };
                isAllfield = this.announcementsValidtion();

                if (isAllfield) {
                    $body.addClass("loading");
                    updateitems("Announcements", strComponentId, saveData, function (e) {

                        if (e.data) {
                            $body.removeClass("loading");
                            that.pageBack();
                            localStorage.clear();
                        } else {
                            $body.removeClass("loading");
                            console.log(e);
                        }
                    });
                    strcrop = "0";
                }

            }
            else if (files.files.length == 0) {

                var saveDatas = {
                    Title: $("#txtTitle").val(),
                    Explanation: $('.richText-editor').html(),
                    ExplanationText: $('.richText-editor').text(),
                    Expires: new Date($('#txtExpires').val()),

                };
                isAllfield = this.announcementsValidtion();

                if (isAllfield) {
                    $body.addClass("loading");
                    updateitems("Announcements", strComponentId, saveDatas, function (e) {

                        if (e.data) {
                            $body.removeClass("loading");
                            that.pageBack();

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
                if (isAllfield) {
                    var fileURL = window.location.origin;
                    $body.addClass("loading");
                    var uniquename = Math.random().toString(36).substr(2, 9) + ".png";
                    var file1 = $('#cropped-img').attr('src').split("base64,");
                    var blob = base64ToArrayBuffer(file1[1]);
                    pnp.sp.web.getFolderByServerRelativeUrl("PublishingImages").files.add(file.name, blob, true)
                        .then(function (result) {
                            pnp.sp.web.lists.getByTitle("Announcements").items.getById(strComponentId).update({
                                ID: strComponentId,
                                Title: $('#txtTitle').val().trim(),
                                Expires: new Date($('#txtExpires').val()),
                                Explanation: $('.richText-editor').html(),
                                ExplanationText: $('.richText-editor').text(),
                                Image: {
                                    "__metadata": {
                                        "type": "SP.FieldUrlValue"
                                    },
                                    Url: fileURL + result.data.ServerRelativeUrl
                                }
                            }).then(r => {
                                $body.removeClass("loading");
                                that.pageBack();

                            });

                        });
                }
            }

        } else if (strLocalStorage == "Holiday") {
            isAllfield = this.holidaysValidtion();

            let myobjHol = {
                Title: $("#txtTitle").val(),
                EndEventDate: $("#txtEEDate").val(),
                EventDate: $("#txtEvDate").val()
            }
            let isDateChecker = this.DateChecker();

            if (isAllfield && isDateChecker) {
                $body.addClass("loading");
                updateitems("Holiday", strComponentId, myobjHol, function (e) {
                    if (e.data) {
                        $body.removeClass("loading");
                        that.pageBack();
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

            if (isAllfield) {
                $body.addClass("loading");
                isAllfield = this.quickLinksValidation();
                updateitems("Quick Links", strComponentId, myobjQl, function (e) {

                    if (e.data) {
                        $body.removeClass("loading");
                        that.pageBack();

                    } else {
                        $body.removeClass("loading");
                        console.log(e);
                    }
                });
            }
        } else if (strLocalStorage == "Employee Corner") {
            var files = <HTMLInputElement>document.getElementById("uploadFile");
            let file = files.files[0];
            if (files.files.length == 0) {

                if (!$('#txtTitle').val().trim()) {

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
                        } else {
                            $body.removeClass("loading");
                            console.log(e);
                        }
                    });

                }
            }
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

                            });
                        });
                }
            }
        } else if (strLocalStorage == "Events") {
            var files = <HTMLInputElement>document.getElementById("inputImage");
            let file = files.files[0];
            if (($('#cropped-img')[0].src == siteURL + '/_catalogs/masterpage/BloomHomepage/images/prof-img.jpg')) {
                isAllfield = this.EventDateChecker();
                let saveEvents = {
                    Title: $("#txtTitle").val(),
                    StartDate: $('#txtStartDate').val(),
                    EndDate: new Date($('#txtEndDate').val()),
                    Explanation: $('.richText-editor').html(),
                    ExplanationText: $('.richText-editor').text(),
                    Image: {
                        "__metadata": {
                            "type": "SP.FieldUrlValue"
                        },
                        Url: siteURL + "/_catalogs/masterpage/Bloom/images/logo.png"
                    }
                }

                if (isAllfield) {
                    $body.addClass("loading");
                    updateitems("Events", strComponentId, saveEvents, function (e) {
                        if (e.data) {
                            $body.removeClass("loading");
                            window.history.back();

                        } else {
                            $body.removeClass("loading");
                            console.log(e);
                        }
                    });
                }
                strcrop = "0";
            }
            else if (strcrop == "1" && files.files.length == 0) {
                isAllfield = this.EventDateChecker();
                let saveEvents = {
                    Title: $("#txtTitle").val(),
                    StartDate: $('#txtStartDate').val(),
                    EndDate: new Date($('#txtEndDate').val()),
                    Explanation: $('.richText-editor').html(),
                    ExplanationText: $('.richText-editor').text(),
                    Image: {
                        "__metadata": {
                            "type": "SP.FieldUrlValue"
                        },
                        Url: siteURL + "/_catalogs/masterpage/Bloom/images/logo.png"
                    }

                }

                if (isAllfield) {
                    $body.addClass("loading");
                    updateitems("Events", strComponentId, saveEvents, function (e) {
                        if (e.data) {
                            $body.removeClass("loading");
                            window.history.back();
                            localStorage.clear();
                        } else {
                            $body.removeClass("loading");

                        }
                    });

                }
                strcrop = "0";
            }

            else if (files.files.length == 0) {
                isAllfield = this.EventDateChecker();
                let saveEvents = {
                    Title: $("#txtTitle").val(),
                    StartDate: $('#txtStartDate').val(),
                    EndDate: new Date($('#txtEndDate').val()),
                    Explanation: $('.richText-editor').html(),
                    ExplanationText: $('.richText-editor').text(),

                }
                    ;
                if (isAllfield) {
                    $body.addClass("loading");
                    updateitems("Events", strComponentId, saveEvents, function (e) {

                        if (e.data) {
                            $body.removeClass("loading");
                            that.pageBack();
                        } else {
                            $body.removeClass("loading");
                            console.log(e);
                        }
                    });

                }
                strcrop = "0";
            } else {
                var fileURL = window.location.origin;
                isAllfield = this.EventDateChecker();
                if (isAllfield) {
                    $body.addClass("loading");
                    var file1 = $('#cropped-img').attr('src').split("base64,");
                    var blob = base64ToArrayBuffer(file1[1]);
                    pnp.sp.web.getFolderByServerRelativeUrl("PublishingImages").files.add(file.name, blob, true)
                        .then(function (result) {
                            pnp.sp.web.lists.getByTitle("Events").items.getById(strComponentId).update({
                                ID: strComponentId,
                                Title: $("#txtTitle").val().trim(),
                                StartDate: new Date($('#txtStartDate').val()),
                                EndDate: new Date($('#txtEndDate').val()),
                                Explanation: $('.richText-editor').html(),
                                ExplanationText: $('.richText-editor').text(),
                                Image: {
                                    "__metadata": {
                                        "type": "SP.FieldUrlValue"
                                    },
                                    Url: fileURL + result.data.ServerRelativeUrl
                                }
                            }).then(r => {
                                $body.removeClass("loading");
                                that.pageBack();

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
                    Departments: $('#ddlDepartment').val(),
                    Explanation: $('.richText-editor').html(),
                }
                if (isAllfield) {
                    $body.addClass("loading");
                    updateitems("Organizational Policies", strComponentId, saveData, function (e) {
                        if (e.data) {
                            $body.removeClass("loading");
                            that.pageBack();

                        } else {
                            $body.removeClass("loading");

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
                                Explanation: $('.richText-editor').html(),
                                DocumentFile: {
                                    "__metadata": {
                                        "type": "SP.FieldUrlValue"
                                    },
                                    Url: fileURL + result.data.ServerRelativeUrl
                                }

                            }).then(r => {
                                $body.removeClass("loading");
                                that.pageBack();

                            });

                        });
                }
            }

        } else if (strLocalStorage == "Banners") {
            var regexp = /^(https?|s?ftp):\/\/(((([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(%[\da-f]{2})|[!\$&'\(\)\*\+,;=]|:)*@)?(((\d|[1-9]\d|1\d\d|2[0-4]\d|25[0-5])\.(\d|[1-9]\d|1\d\d|2[0-4]\d|25[0-5])\.(\d|[1-9]\d|1\d\d|2[0-4]\d|25[0-5])\.(\d|[1-9]\d|1\d\d|2[0-4]\d|25[0-5]))|((([a-z]|\d|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(([a-z]|\d|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])*([a-z]|\d|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])))\.)+(([a-z]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(([a-z]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])*([a-z]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])))\.?)(:\d*)?)(\/((([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(%[\da-f]{2})|[!\$&'\(\)\*\+,;=]|:|@)+(\/(([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(%[\da-f]{2})|[!\$&'\(\)\*\+,;=]|:|@)*)*)?)?(\?((([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(%[\da-f]{2})|[!\$&'\(\)\*\+,;=]|:|@)|[\uE000-\uF8FF]|\/|\?)*)?(#((([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(%[\da-f]{2})|[!\$&'\(\)\*\+,;=]|:|@)|\/|\?)*)?$/i
            var files = <HTMLInputElement>document.getElementById("inputImage");
            let file = files.files[0];
            if ($('#cropped-img')[0].src && $('#cropped-img')[0].src == siteURL + '/_catalogs/masterpage/BloomHomepage/images/prof-img.jpg') {
                alertify.set('notifier', 'position', 'top-right');
                alertify.error("Please Select the Image");
                isAllfield = false;
            }
            else if (!$('#txtTitle').val().trim()) {
                alertify.set('notifier', 'position', 'top-right');
                alertify.error("Please Enter the Title");
                isAllfield = false;

            } else if ($('#txtHyper').val() && !regexp.test($('#txtHyper').val().trim())) {
                alertify.set('notifier', 'position', 'top-right');
                alertify.error("Please give a valid Link URL");
                isAllfield = false;
            }
            else if (!$('.richText-editor').text().trim()) {
                alertify.set('notifier', 'position', 'top-right');
                alertify.error("Please Enter the Description");
                isAllfield = false;
            }
            if (files.files.length > 0) {
                var uniquename = Math.random().toString(36).substr(2, 9) + ".png";
                var file1 = $('#cropped-img').attr('src').split("base64,");
                var blob = base64ToArrayBuffer(file1[1]);

            }
            if (isAllfield) {
                if (files.files.length == 0) {
                    let myobjBanners = {
                        Title: $("#txtTitle").val(),
                        BannerContent: $('.richText-editor').html(),
                        LinkURL: {
                            "__metadata": {
                                "type": "SP.FieldUrlValue"
                            },
                            Url: $('#txtHyper').val().trim(),
                        }
                    }
                    $body.addClass("loading");
                    updateitems("Banners", strComponentId, myobjBanners, function (e) {

                        if (e.data) {
                            $body.removeClass("loading");
                            that.pageBack();

                        } else {
                            $body.removeClass("loading");

                        }
                    });

                }
                else
                    var fileURL = window.location.origin;
                $body.addClass("loading");
                pnp.sp.web.getFolderByServerRelativeUrl("Images").files.add(uniquename, blob, true)
                    .then(function (result) {
                        pnp.sp.web.lists.getByTitle("Banners").items.getById(strComponentId).update({
                            ID: strComponentId,
                            Title: $("#txtTitle").val(),
                            BannerContent: $('.richText-editor').html(),
                            Image: {
                                "__metadata": {
                                    "type": "SP.FieldUrlValue"
                                },
                                Url: fileURL + result.data.ServerRelativeUrl
                            },
                            LinkURL: {
                                "__metadata": {
                                    "type": "SP.FieldUrlValue"
                                },
                                Url: $('#txtHyper').val().trim(),
                            }
                        }).then(r => {
                            $body.removeClass("loading");
                            window.history.back();

                        });

                    });
            }

        } else if (strLocalStorage == "Polls") {
            isAllfield = this.pollsValidation()
            var updatequs = $("#txtQuestion").val().trim().length;
            var updateans = $("#txtOptions").val().trim().length;
            if (this.qustion_lenght == updatequs && this.answer_lenght == updateans) {
                let myobjPols = {
                    Question: $("#txtQuestion").val(),
                    Options: $("#txtOptions").val(),
                    IsVisibles: $('#anonymous').prop('checked')
                }
                if (isAllfield) {
                    $body.addClass("loading");
                    updateitems("Polls", strComponentId, myobjPols, function (e) {

                        if (e.data) {
                            $body.removeClass("loading");
                            that.pageBack();
                        } else {
                            $body.removeClass("loading");
                            console.log(e);
                        }
                    });

                }
            }
            if (this.qustion_lenght != updatequs || this.answer_lenght != updateans) {
                pnp.sp.web.lists.getByTitle("PollsResults").items.filter("QuestionID eq '" + strComponentId + "'").get().then((items: any[]) => {
                    for (var i = 0; i < items.length; i++) {
                        pnp.sp.web.lists.getByTitle("PollsResults").items.getById(items[i].ID).delete().then((result: any) => {
                            let myobjPols = {
                                Question: $("#txtQuestion").val(),
                                Options: $("#txtOptions").val(),
                                IsVisibles: $('#anonymous').prop('checked')
                            }
                            if (isAllfield) {
                                $body.addClass("loading");
                                updateitems("Polls", strComponentId, myobjPols, function (e) {
                                    if (e.data) {
                                        $body.removeClass("loading");
                                        that.pageBack();
                                    } else {
                                        $body.removeClass("loading");
                                        console.log(e);
                                    }
                                });

                            }
                        });
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

                        } else {
                            $body.removeClass("loading");

                        }
                    });
                }
            } else if (files.files.length > 0 && docfiles.files.length > 0) {
                var fileURL = window.location.origin;
                var file1 = $('#cropped-img').attr('src').split("base64,");
                var blob = base64ToArrayBuffer(file1[1]);

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
        var siteURL = this.context.pageContext.web.absoluteUrl;
        var renderhtml = "<ul>";

        var rendertext = "";
        var renderdate = "";
        var renderhtmlCheckBox = "";
        var renderEventDate = "";
        var renderHyperlink = "";
        var renderHyperSitelink = "";
        var renderUploadfile = "";
        var renderCorpUploadfile = "";
        var newrichTextEditor = "";
        var requirednewrichTextEditor = "";
        var renderSiteLink = "";
        var renderStartEndDate = "";
        var renderhtmlImageBanners = "";
        var renderQuestion = "";
        var renderAnswers = "";
        var renderDepartment = "";
        var renderDepartmentddl = "";
        var strLocalStorage = GetQueryStringParams("CName");
        strLocalStorage = strLocalStorage.split('%20').join(' ');

        var strComponentMode = GetQueryStringParams("CMode");

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

        rendertext += "<div id='renderText' class='input text'>" +
            "<label class='control-label'>Title</label>" +
            "<input class='form-control' type='text' value='' id='txtTitle' /></div>";

        renderdate += "<div class='input date'><i class='icon-calenter'></i>" +
            "<label class='control-label'>Date</label>" +
            "<input class='form-control date-selector' type='text' value='' id='txtExpires' /></div>";

        newrichTextEditor += "<div class='textarea input'>" +
            "<label>Description</label>" +
            "<textarea id='txtDescription' class='form-control content'></textarea>" +
            "</div>";

        requirednewrichTextEditor += "<div class='textarea input'>" +
            "<label class='control-label'>Description</label>" +
            "<textarea id='txtrequiredDescription' class='form-control content'></textarea>" +
            "</div>";
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
            "<span>Please Enter more than one Options with Semicolon ( ; ) Maximum Four answers</span>" +
            "</div>";




        this.getListItems(strComponentId);
        $('.appendsec').append(renderhtml);
        console.log(strLocalStorage);
        var date = new Date();
        var today = new Date(date.getFullYear(), date.getMonth(), date.getDate());

        if (strLocalStorage == 'Announcements') {

            $('.form-section').append(renderhtmlImageBanners);
            $('#my-image,#getcroppie,#cancel').hide();
            //  $('.form-imgsec').after(rendercrop);
            $('.form-imgsec').after(renderdate);
            $('.date').after(rendertext);
            $('.text').after(requirednewrichTextEditor);
            // $('.text').after(newrichTextEditor);
            $('#txtExpires').datepicker({
                format: "mm/dd/yyyy",
                startDate: today
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

            });
            $('#txtEEDate').datepicker({
                format: "mm/dd/yyyy",

            });
            this.ViewMode(strComponentMode);
        } else if (strLocalStorage == 'Quick Links') {
            $('.form-section').append(rendertext);
            $('.text').after(renderHyperlink);
            this.ViewMode(strComponentMode);
        } else if (strLocalStorage == 'Employee Corner') {
            $('.form-section').append(rendertext);
            $('.text').after(renderUploadfile);

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
            $('.text').after(renderDepartmentddl + renderDepartment + renderUploadfile + newrichTextEditor);
            $("#DivDepartment").hide();
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
            $('#my-image,#getcroppie,#cancel').hide();
            $('.form-imgsec').after(rendertext);
            $('.text').after(requirednewrichTextEditor);
            $('.textarea').after(renderHyperSitelink);
            this.ViewMode(strComponentMode);
        } else if (strLocalStorage == 'Corporate Discounts') {
            $('.form-section').append(renderhtmlImageBanners);
            $('#my-image,#getcroppie,#cancel').hide();

            $('.form-imgsec').after(rendertext);
            $('.text').after(renderSiteLink);
            $('#siteLink').after(renderCorpUploadfile);

            this.ViewMode(strComponentMode);

        } else if (strLocalStorage == 'Events') {
            $('.form-section').append(renderhtmlImageBanners);
            $('#my-image,#getcroppie,#cancel').hide();
            $('#imageLabel').removeClass('control-label');
            $('.form-imgsec').after(rendertext);
            $('#renderText').after(newrichTextEditor);
            $('.textarea').after(renderStartEndDate);
            $('#txtStartDate').datepicker({
                format: "mm/dd/yyyy",
            });
            $('#txtEndDate').datepicker({
                format: "mm/dd/yyyy",
            });
            this.ViewMode(strComponentMode);

        } else if (strLocalStorage == 'Polls') {
            $('.form-section').append(renderQuestion);
            $('.textarea').after(renderAnswers);
            $('#answers').after(renderhtmlCheckBox);
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

        $('.content').richText();
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
            if ($('#cropped-img')[0].src == siteURL + '/_catalogs/masterpage/BloomHomepage/images/prof-img.jpg') {
                alertify.set('notifier', 'position', 'top-right');
                alertify.error("Please upload the Image File");
                $("#inputImage").val("");
            }
            else {
                $('#cropped-img')[0].src = siteURL + '/_catalogs/masterpage/BloomHomepage/images/prof-img.jpg';
                $('#inputImage').val("");
            }
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
                $('#cropped-img').attr("src", items[0].Image.Url);
                $('#txtTitle').val(items[0].Title);
                var eedate = "";
                if ((items[0].Expires) != null) {
                    eedate = this.nullDateValidate(items[0].Expires);
                }
                $('#txtExpires').datepicker('setDate', new Date(eedate));
                $('.richText-editor').html(items[0].Explanation);
            })
        } else if (strLocalStorage == "Holiday") {
            objResults = readItems("Holiday", ["Title", "Modified", "EventDate", "EndEventDate", "Display"], count, "Modified", "ID", strComponentId)
            objResults.then((items: any[]) => {
                $('#txtTitle').val(items[0].Title);
                var eedate = "";
                if ((items[0].EventDate) != null) {
                    eedate = this.nullDateValidate(items[0].EventDate);
                }
                $('#txtEvDate').datepicker('setDate', new Date(eedate));
                if ((items[0].EndEventDate) != null) {
                    eedate = this.nullDateValidate(items[0].EndEventDate);
                }
                $('#txtEEDate').datepicker('setDate', new Date(eedate));

            })
        } else if (strLocalStorage == "News") {
            objResults = readItems("News", ["Title", "Modified", "Date", "Image", "Explanation", "Display",], count, "Modified", "ID", strComponentId)
            objResults.then((items: any[]) => {
                $('.crapImagesevent').attr("src", items[0].Image.Url);
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
            this.bindorgDept()

            objResults = readItems("Organizational Policies", ["ID", "Title", "Departments", "Modified", "DocumentFile", "Explanation",], count, "Modified", "ID", strComponentId)
            objResults.then((items: any[]) => {
                var orgDept = items[0].Departments;
                $("#txtTitle").val(items[0].Title);
                $("#txtDepartment").val(items[0].Departments);
                $("#ddlDepartment").val(items[0].Departments);
                if (items[0].Departments == 'Others') {
                    $("#DivDepartment").hide();
                }


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

                $('.richText-editor').html(items[0].Explanation);

            })
        } else if (strLocalStorage == "Banners") {
            objResults = readItems("Banners", ["Title", "Modified", "BannerContent", "Display", "LinkURL", "Order", "Image"], count, "Modified", "ID", strComponentId)
            objResults.then((items: any[]) => {
                $('#cropped-img').attr("src", items[0].Image.Url);
                $('#txtTitle').val(items[0].Title);
                $('.richText-editor').html(items[0].BannerContent)
                if (items[0].LinkURL == null) {
                    $('#txtHyper').val('');
                } else {
                    $('#txtHyper').val(items[0].LinkURL.Url)
                }
            })
        } else if (strLocalStorage == "Corporate Discounts") {
            objResults = readItems("Corporate Discounts", ["Title", "Modified", "VendorLogo", "SiteLink", "DocumentFile"], count, "Modified", "ID", strComponentId)
            objResults.then((items: any[]) => {
                $('#cropped-img').attr("src", items[0].VendorLogo.Url);

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
                $('#cropped-img').attr("src", items[0].Image.Url);
                $('#txtTitle').val(items[0].Title);
                $('.richText-editor').html(items[0].Explanation);
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
            objResults = readItems("Polls", ["ID", "Title", "Modified", "Display", "Question", "Options", "IsVisibles"], count, "Modified", "ID", strComponentId)
            objResults.then((items: any[]) => {
                this.qustion_lenght = items[0].Question.length;
                this.answer_lenght = items[0].Options.length;
                $('#txtQuestion').val(items[0].Question);
                $('#txtOptions').val(items[0].Options);
                if (items[0].IsVisibles) {
                    $('#anonymous').prop('checked', true)
                } else {
                    $('#anonymous').prop('checked', false)
                }

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
