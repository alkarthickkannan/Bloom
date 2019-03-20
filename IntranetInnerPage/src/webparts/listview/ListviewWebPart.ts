import { Version } from '@microsoft/sp-core-library';
import {
    BaseClientSideWebPart,
    IPropertyPaneConfiguration,
    PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';
import 'jquery';
//import styles from './ListviewWebPart.module.scss';
import * as strings from 'ListviewWebPartStrings';
import pnp from 'sp-pnp-js';
import {
    readItems, updateItem, formatDate, checkUserinGroup, GetQueryStringParams, batchDelete
} from '../../commonJS';
require('jplist-core');
require('jplist-pagination');
require('../../ExternalRef/js/jplist-core.js');
require('../../ExternalRef/js/jplist-pagination.js');
declare var $;
declare var alertify: any;
export interface IListviewWebPartProps {
    description: string;
}

export default class ListviewWebPart extends BaseClientSideWebPart<IListviewWebPartProps> {
    userflag: boolean = false;
    public render(): void {
        var _this = this;
        var strLocalStorage = GetQueryStringParams("CName").replace("%20", " ");
        //Checking user details in group
        checkUserinGroup(strLocalStorage, this.context.pageContext.user.email, function (result) {
            //console.log(result);
            if (result == 1) {
                _this.userflag = true;
            }
            else {
                if (strLocalStorage == "Quick Links" || strLocalStorage == "Employee Corner") {
                    alertify.alert('Access Denied', 'Sorry You dont have access to this page', function () {
                        history.go(-1);
                    }).set('closable', false);
                }
            }
            _this.viewlistitemdesign();
        })
    }
    public viewlistitemdesign() {

        var siteURL = this.context.pageContext.web.absoluteUrl;
        var strLocalStorage = GetQueryStringParams("CName").replace("%20", " ");
        this.domElement.innerHTML =
            "<div class='breadcrumb'>" +
            "<ol>" +
            "<li><a href='" + siteURL + "/Pages/Home.aspx' title='Home'>Home</a></li>" +
            "<li><span id='breadcrumb-name'></span></li>" +
            "</ol>" +
            "</div>" +
            "<div class='title-section'>" +
            "<div class='button-field'>" +
            "<a class='add-class pointer'  title='Add New'><i class='icon-add add-class'></i>Add New</a>" +
            "<a class='delete-icon pointer' title='Delete' id='deleteitems'><img src='" + siteURL + "/_catalogs/masterpage/BloomHomepage/images/close-icon.png'>Delete</a>" +
            "</div>" +
            "<h2 id='ComponentName'></h2>" +
            "</div>" +
            "<div class='content-area'>" +

            "<div class='list-tabcont'>" +
            "<div class='list-tabcontsec'>" +
            "</div>" +
            "<div class='list-tabcontsec'>" +
            "</div>" +
            "</div>" +

            "</div>" +
            "<div class='modal'><!-- Place at bottom of page --></div>";

        // localStorage.getItem("ComponentName");
        document.title = strLocalStorage + '-View';
        this.ViewListItems(GetQueryStringParams("CName").replace("%20", " "));
        if (this.userflag == false) {
            $('.button-field').hide();
        }
        else {
            $('.button-field').show();
        }
    }
    public ViewListItems(strLocalStorage) {
        document.getElementById("ComponentName").innerHTML = strLocalStorage;
        document.getElementById("breadcrumb-name").innerHTML = strLocalStorage;
        var count = 50;
        //var strLocalStorage = GetQueryStringParams("CName").replace("%20", " ");

        //localStorage.getItem("ComponentName") || "";
        let objResults;
        if (this.userflag == false) {
            if (strLocalStorage == "Holiday") {
                objResults = readItems(strLocalStorage, ["ID", "Title", "Modified", "EventDate", "EndEventDate", "Display"], count, "Modified", "Display", 1);
                objResults.then((items: any[]) => {
                    this.renderhtml(items, strLocalStorage);
                });
            } else if (strLocalStorage == "Announcements") {
                objResults = readItems(strLocalStorage, ["ID", "Title", "Expires", "Modified", "Image", "Explanation", "Expires", "Display", "ViewedUsers"], count, "Modified", "Display", 1);
                objResults.then((items: any[]) => {
                    this.renderhtml(items, strLocalStorage);
                });

            } else if (strLocalStorage == "Banners") {
                objResults = readItems(strLocalStorage, ["ID", "Title", "Modified", "Image", "BannerContent", "Display"], count, "Modified", "Display", 1);
                objResults.then((items: any[]) => {
                    this.renderhtml(items, strLocalStorage);
                });
            }
            else if (strLocalStorage == "Quick Links") {
                objResults = readItems(strLocalStorage, ["ID", "Title", "Modified", "LinkURL", "Display"], count, "Modified", "Display", 1);
                objResults.then((items: any[]) => {
                    this.renderhtml(items, strLocalStorage);
                });
            }
            else if (strLocalStorage == "News") {
                objResults = readItems(strLocalStorage, ["ID", "Title", "Date", "Modified", "Image", "Display", "Explanation"], count, "Modified", "Display", 1);
                objResults.then((items: any[]) => {
                    this.renderhtml(items, strLocalStorage);
                });
            }
            else if (strLocalStorage == "Employee Corner") {
                objResults = readItems(strLocalStorage, ["ID", "Title", "Modified", "Display", "DocumentFile", "File_x0020_Type"], count, "Modified", "Display", 1);
                objResults.then((items: any[]) => {
                    this.renderhtml(items, strLocalStorage);
                });
            }
            else if (strLocalStorage == "Events") {
                objResults = readItems(strLocalStorage, ["ID", "Title", "Modified", "Display", "Explanation", "HyperLink", "StartDate", "EndDate"], count, "Modified", "Display", 1);
                objResults.then((items: any[]) => {
                    this.renderhtml(items, strLocalStorage);
                });
            }
            else if (strLocalStorage == "Polls") {
                //pnp.sp.web.lists.getByTitle("Polls").items.select("ID","Display").filter("Display eq 1").top(100).get()
                //objResults = readItems(strLocalStorage, ["ID", "Title", "Modified", "Display", "Question", "Options"], count, "Modified", "Display", 1);
                objResults = pnp.sp.web.lists.getByTitle("Polls").items.select("ID", "Title", "Modified", "Display", "Question", "Options", "IsDeleted").filter("IsDeleted eq 0").top(100).get();
                objResults.then((items: any[]) => {
                    this.renderhtml(items, strLocalStorage);
                });
            }
            else if (strLocalStorage == "Corporate Discount") {
                objResults = readItems(strLocalStorage, ["ID", "Title", "Modified", "Display", "SiteLink"], count, "Modified", "Display", 1);
                objResults.then((items: any[]) => {
                    this.renderhtml(items, strLocalStorage);
                });
            }
        } else {
            if (strLocalStorage == "Holiday") {
                objResults = readItems(strLocalStorage, ["ID", "Title", "Modified", "EventDate", "EndEventDate", "Display"], count, "Modified");
                objResults.then((items: any[]) => {
                    this.renderhtml(items, strLocalStorage);
                });
            } else if (strLocalStorage == "Announcements") {
                objResults = readItems(strLocalStorage, ["ID", "Title", "Expires", "Modified", "Image", "Explanation", "Expires", "Display"], count, "Modified");
                objResults.then((items: any[]) => {
                    this.renderhtml(items, strLocalStorage);
                });
            } else if (strLocalStorage == "Banners") {
                objResults = readItems(strLocalStorage, ["ID", "Title", "Modified", "Image", "BannerContent", "Display"], count, "Modified");
                objResults.then((items: any[]) => {
                    this.renderhtml(items, strLocalStorage);
                });
            } else if (strLocalStorage == "Quick Links") {
                objResults = readItems(strLocalStorage, ["ID", "Title", "Modified", "Display", "LinkURL"], count, "Modified");
                objResults.then((items: any[]) => {
                    this.renderhtml(items, strLocalStorage);
                });
            }
            else if (strLocalStorage == "News") {
                objResults = readItems(strLocalStorage, ["ID", "Title", "Date", "Modified", "Display", "Image", "Explanation"], count, "Modified");
                objResults.then((items: any[]) => {
                    this.renderhtml(items, strLocalStorage);
                });
            }
            else if (strLocalStorage == "Employee Corner") {
                objResults = readItems(strLocalStorage, ["ID", "Title", "Modified", "Display", "DocumentFile", "File_x0020_Type"], count, "Modified");
                objResults.then((items: any[]) => {
                    this.renderhtml(items, strLocalStorage);
                });
            }
            else if (strLocalStorage == "Events") {
                objResults = readItems(strLocalStorage, ["ID", "Title", "Modified", "Display", "Explanation", "HyperLink", "StartDate", "EndDate"], count, "Modified");
                objResults.then((items: any[]) => {
                    this.renderhtml(items, strLocalStorage);
                });
            }
            else if (strLocalStorage == "Polls") {
                objResults = pnp.sp.web.lists.getByTitle("Polls").items.select("ID", "Title", "Modified", "Display", "Question", "Options", "IsDeleted").filter("IsDeleted eq 0").top(count).get();
                // objResults = readItems(strLocalStorage, ["ID", "Title", "Modified", "Display", "Question", "Options"], count, "Modified");
                objResults.then((items: any[]) => {
                    this.renderhtml(items, strLocalStorage);
                });
                this.displaypollcheck();
            }
            else if (strLocalStorage == "Corporate Discount") {
                objResults = readItems(strLocalStorage, ["ID", "Title", "Modified", "Display", "SiteLink"], count, "Modified");
                objResults.then((items: any[]) => {
                    this.renderhtml(items, strLocalStorage);
                });
            }
        }

    }
    displaypollcheck() {
        pnp.sp.web.lists.getByTitle("Polls").items.select("ID", "Display").filter("Display eq 1").top(100).get().then((items: any) => {
            for (let j = 0; j < items.length; j++) {
                var displayitem = Number(items[j].Display);
                if (displayitem == 1) {
                    //$('.delete-item').hide();deleteitem111
                    $('#deleteitem' + items[j].ID).hide();
                }
            }
        })
    }
    renderhtml(objResults, strLocalStorage) {
        var renderhtml = "<div id='pagination-list' class='list-section jplist'><ul class='list'>";
        var checkboxstatus = "";
        var strcheckboxstatus = "Not Displayed";
        //var strLocalStorage = GetQueryStringParams("CName").replace("%20", " ");
        var siteURL = this.context.pageContext.web.absoluteUrl;
        if (objResults.length > 0) {
            objResults.sort(function (a, b) {
                return new Date(b.Modified).getTime() - new Date(a.Modified).getTime();
            });
            for (let i = 0; i < objResults.length; i++) {
                if (objResults[i].Display == "1") {
                    checkboxstatus = "checked";
                    strcheckboxstatus = "Displayed";
                }
                else {
                    checkboxstatus = "";
                    strcheckboxstatus = "Not Displayed";
                }
                renderhtml += "<li class='list-item'>" +
                    "<div class='list-imgcont img-bind" + [i] + "'>" +
                    "<span class='displaydate" + [i] + "'></span>" +
                    "<h3 class='item-title" + [i] + "'></h3>" +
                    "<p class='add-description" + [i] + "'></p>" +
                    "<div class='switch'>" +
                    "<input type='checkbox' id='switch" + objResults[i].ID + "' class='switch-input' " + checkboxstatus + "/>" +
                    "<label for='switch" + objResults[i].ID + "' class='switch-label'>" + strcheckboxstatus + "</label>" +
                    "</div>" +
                    "<div class='list-icons'>" +
                    "<div class='icon-list2 viewitem' style='cursor: pointer;' id='viewitem" + objResults[i].ID + "'>" +
                    "<a  title='View' class='viewitem' ><i class='icon-eye viewitem'></i></a>" +
                    "</div>" +
                    "<div class='icon-list2 edititemuser edititem' style='cursor: pointer;' id='edititem" + objResults[i].ID + "'>" +
                    "<a  title='Edit' class='edititem' ><i class='icon-edit edititem'></i></a>" +
                    "</div>" +
                    "<div class='icon-list2 deleteitemuser likecounts" + objResults[i].ID + "'>" +
                    "<div class='check-box'>" +
                    "<input type='checkbox' name='' value='' class='delete-item' id='deleteitem" + objResults[i].ID + "'/>" +
                    "<label>Checkbox</label>" +
                    "</div>" +
                    "</div>" +
                    "</div>" +
                    "</div>" +
                    "</li>";
            }
        }
        else {
            renderhtml += "<li class='list-item'>No items to display" +
                "</li>";
        }
        renderhtml += "</ul>";
        renderhtml += "<div class='jplist-panel box panel-top'>" +
            "<div class='jplist-pagination' data-control-type='pagination' data-control-name='paging' data-control-action='paging'></div>" +
            "<select class='jplist-select' data-control-type='items-per-page-select' data-control-name='paging' data-control-action='paging'>" +
            "<option data-number='5' data-default='true'> 5 </option>" +
            "<option data-number='10'> 10 </option>" +
            "<option data-number='15'> 15 </option>" +
            "</select>" +
            "</div>";
        $('.content-area').append(renderhtml);

        if (this.userflag == false) {
            $('.edititemuser').hide();
            $('.deleteitemuser').hide();
            $('.switch').hide();
        }
        else {
            $('.edititemuser').show();
            $('.deleteitemuser').show();
        }
        //console.log(strLocalStorage);

        if (strLocalStorage == 'Holiday') {
            for (let i = 0; i < objResults.length; i++) {
                $('.item-title' + i).append(objResults[i].Title);
                $('.displaydate' + i).append("<strong>" + formatDate(objResults[i].EventDate) + "</strong>");
                var eedate = "";
                if ((objResults[i].EndEventDate) != null) {
                    eedate = formatDate(objResults[i].EndEventDate);
                }
                var edate = "";
                if ((objResults[i].EventDate) != null) {
                    edate = formatDate(objResults[i].EventDate);
                }
                if (eedate == "" && edate == "") {
                    $('.add-description' + i).append("");
                }
                else if (eedate == "" && edate != "") {
                    $('.add-description' + i).append("Start date: <strong>" + edate + "</strong>");
                }
                else if (edate == "" && eedate != "") {
                    $('.add-description' + i).append("End date: <strong>" + edate + "</strong>");
                }
                else {
                    $('.add-description' + i).append("Start date: <strong>" + edate + "</strong> End date: <strong>" + eedate + "</strong>");
                }
            }
            $('.title-section').after("<div class='list-tab'><ul><li class='event-class'>Events</li><li class='active holiday-class'>Holidays</li></ul></div>");
        } else if (strLocalStorage == 'Announcements') {
            let defaultimage = siteURL + "/_catalogs/masterpage/BloomHomepage/images/logo.png";
            for (let i = 0; i < objResults.length; i++) {
                if (objResults[i].Image != null) {
                    $('.img-bind' + i).prepend("<div class='list-imgsec'><img src='" + objResults[i].Image.Url + "' alt='' title=''/></div>");
                } else {
                    $('.img-bind' + i).prepend("<div class='list-imgsec'><img src='" + defaultimage + "' alt='' title=''/></div>");
                }
                var eedate = "";
                if ((objResults[i].Expires) != null) {
                    eedate = formatDate(objResults[i].Expires);
                }
                if (objResults[i].Explanation != null && objResults[i].Explanation.length > 160) {
                    objResults[i].Explanation = objResults[i].Explanation.substring(0, 160) + "...";
                }
                // $('.add-description' + i).append(objResults[i].Explanation);
                $('.item-title' + i).append(objResults[i].Title);
                $('.displaydate' + i).append("<strong>" + eedate + "</strong>");
                if (this.userflag == false) {
                    $('.edititemuser').show();
                    $('.deleteitemuser').show();
                    $('#viewitem' + objResults[i].ID).empty();
                    $('#edititem' + objResults[i].ID).empty();
                    $('.likecounts' + objResults[i].ID).empty();
                    let ViewedUsers = 0;
                    if (objResults[i].ViewedUsers != null) {
                        if (objResults[i].ViewedUsers.split(',') != null)
                            ViewedUsers = objResults[i].ViewedUsers.split(',').length;
                        $('#viewitem' + objResults[i].ID).append("<a href='" + siteURL + "/Pages/Viewlistitem.aspx?CName=" + strLocalStorage + "&CID=" + objResults[i].ID + "' title='View' class='viewitem' ><i class='icon-eye viewitem'><span>" + ViewedUsers + "</span></i></a>");
                    }
                    else {
                        $('#viewitem' + objResults[i].ID).append("<a href='" + siteURL + "/Pages/Viewlistitem.aspx?CName=" + strLocalStorage + "&CID=" + objResults[i].ID + "' title='View' class='viewitem' ><i class='icon-eye viewitem'><span>" + ViewedUsers + "</span></i></a>");
                    }
                    var objResults1 = readItems("AnnouncementComments", ["AnnouncementID"], 1000, "Modified", "AnnouncementID", objResults[i].ID);
                    objResults1.then((itemsCount: any[]) => {
                        $('#edititem' + objResults[i].ID).append("<a href='" + siteURL + "/Pages/Viewlistitem.aspx?CName=" + strLocalStorage + "&CID=" + objResults[i].ID + "' title='View' class='viewitem' ><i class='icon-comments viewitem'><span>" + itemsCount.length + "</span></i></a>");

                    });
                    var objResults2 = readItems("AnnouncementsLikes", ["AnnouncementID", "Liked"], 1000, "Modified", "AnnouncementID", objResults[i].ID);
                    objResults2.then((itemsCount2: any[]) => {
                        let LikesCount = 0;
                        for (let j = 0; j < itemsCount2.length; j++) {
                            if (itemsCount2[j].Liked == true) {
                                LikesCount++;
                            }
                        }
                        $('.likecounts' + objResults[i].ID).append("<a href='" + siteURL + "/Pages/Viewlistitem.aspx?CName=" + strLocalStorage + "&CID=" + objResults[i].ID + "' title='Edit'><i class='icon-heart'><span>" + LikesCount + "</span></i></a>");
                    });
                }
            }
        }
        else if (strLocalStorage == 'Banners') {
            let defaultimage = siteURL + "/_catalogs/masterpage/BloomHomepage/images/logo.png";
            for (let i = 0; i < objResults.length; i++) {
                if (objResults[i].Image != null) {
                    $('.img-bind' + i).prepend("<div class='list-imgsec'><img src='" + objResults[i].Image.Url + "' alt='' title=''/></div>");
                } else {
                    $('.img-bind' + i).prepend("<div class='list-imgsec'><img src='" + defaultimage + "' alt='' title=''/></div>");
                }
                if (objResults[i].BannerContent != null && objResults[i].BannerContent.length > 160) {
                    objResults[i].BannerContent = objResults[i].BannerContent.substring(0, 160) + "...";
                }
                //  $('.add-description' + i).append(objResults[i].BannerContent);
                $('.item-title' + i).append(objResults[i].Title);
            }
        } else if (strLocalStorage == "Quick Links") {
            for (let i = 0; i < objResults.length; i++) {
                if (objResults[i].LinkURL != null) {
                    $('.item-title' + i).append("<a href='" + objResults[i].LinkURL.Url + "' target='_blank'>" + objResults[i].Title + "</a>");
                } else {
                    $('.item-title' + i).append("<a href='#' target='_blank'>" + objResults[i].Title + "</a>");
                }
            }
        }
        else if (strLocalStorage == "News") {
            let defaultimage = siteURL + "/_catalogs/masterpage/BloomHomepage/images/logo.png";
            for (let i = 0; i < objResults.length; i++) {
                if (objResults[i].Image != null) {
                    $('.img-bind' + i).prepend("<div class='list-imgsec'><img src='" + objResults[i].Image.Url + "' alt='' title=''/></div>");
                } else {
                    $('.img-bind' + i).prepend("<div class='list-imgsec'><img src='" + defaultimage + "' alt='' title=''/></div>");
                }
                if (objResults[i].Explanation != null && objResults[i].Explanation.length > 160) {
                    objResults[i].Explanation = objResults[i].Explanation.substring(0, 160) + "...";
                }
                // $('.add-description' + i).append(objResults[i].Explanation);
                $('.item-title' + i).append(objResults[i].Title);
                var eedate = "";
                if ((objResults[i].Date) != null) {
                    eedate = formatDate(objResults[i].Date);
                }
                $('.displaydate' + i).append("<strong>" + eedate + "</strong>");
            }
        }
        else if (strLocalStorage == "Employee Corner") {
            let defaultimage = siteURL + "/_catalogs/masterpage/BloomHomepage/images/logo.png";
            for (let i = 0; i < objResults.length; i++) {
                if (objResults[i].DocumentFile != null) {
                    var filetype = objResults[i].DocumentFile.Url.split('.').pop();
                    if (filetype == "pdf") {
                        defaultimage = siteURL + "/_catalogs/masterpage/BloomHomepage/images/pdf-view.png";
                    }
                    else if (filetype == "doc" || filetype == "docx") {
                        defaultimage = siteURL + "/_catalogs/masterpage/BloomHomepage/images/doc-view.png";
                    }
                    else if (filetype == "ppt") {
                        defaultimage = siteURL + "/_catalogs/masterpage/BloomHomepage/images/ppt-view.png";
                    } else if (filetype == "xls" || filetype == "csv") {
                        defaultimage = siteURL + "/_catalogs/masterpage/BloomHomepage/images/xls-view.png";
                    } else if (filetype == "jpg" || filetype == "png" || filetype == "jpeg") {
                        defaultimage = siteURL + "/_catalogs/masterpage/BloomHomepage/images/img-view.png";
                    }
                    $('.item-title' + i).append("<a href='" + objResults[i].DocumentFile.Url + "' target='_blank'>" + objResults[i].Title + "</a>");
                    $('.img-bind' + i).prepend("<div class='list-imgsec'><img src='" + defaultimage + "' alt='' title=''/></div>");
                }
            }
        }
        else if (strLocalStorage == 'Events') {
            for (let i = 0; i < objResults.length; i++) {
                $('.item-title' + i).append(objResults[i].Title);
                var eedate = "";
                if ((objResults[i].EndDate) != null) {
                    eedate = formatDate(objResults[i].EndDate);
                }
                var edate = "";
                if ((objResults[i].StartDate) != null) {
                    edate = formatDate(objResults[i].StartDate);
                }
                if (objResults[i].Explanation != null && objResults[i].Explanation.length > 160) {
                    objResults[i].Explanation = objResults[i].Explanation.substring(0, 160) + "...";
                }
                // if (eedate == "" && edate == "") {
                //     $('.add-description' + i).append(objResults[i].Explanation);
                // }
                else if (eedate == "" && edate != "") {
                    $('.add-description' + i).append("Start date: <strong>" + edate + "</strong></br>");
                }
                else if (edate == "" && eedate != "") {
                    $('.add-description' + i).append("End date: <strong>" + edate + "</strong></br>");
                }
                else {
                    $('.add-description' + i).append("Start date: <strong>" + edate + "</strong> End date: <strong>" + eedate + "</strong></br>");
                }
                //$('.add-description' + i).append("Start date: <strong>" + edate + "</strong> End date: <strong>" + eedate + "</strong></br>" + objResults[i].Explanation);
                $('.displaydate' + i).append("<strong>" + formatDate(objResults[i].Modified) + "</strong>");
            }

            $('.title-section').after("<div class='list-tab'><ul><li class='active event-class'>Events</li><li class='holiday-class'>Holidays</li></ul></div>");
        }
        else if (strLocalStorage == 'Polls') {
            for (let i = 0; i < objResults.length; i++) {
                $('.item-title' + i).append(objResults[i].Question);
                if (objResults[i].Options != null && objResults[i].Options.length > 35) {
                    objResults[i].Options = objResults[i].Options.substring(0, 160) + "...";
                }
                //$('.add-description' + i).append(objResults[i].Options);
                if (objResults[i].Options.split(';') != null) {
                    //let arrOption = [];
                    let renderOptions = "";
                    let arrOption = objResults[i].Options.split(';');
                    arrOption = arrOption.filter(function (v) {
                        return /\S/.test(v);
                    });
                    for (var j = 0; j < arrOption.length; j++) {
                        renderOptions += arrOption[j] + "</br>";
                    }
                    $('.add-description' + i).append(renderOptions);
                }
                else {
                    $('.add-description' + i).append(objResults[i].Options);
                }
            }
        }
        else if (strLocalStorage == 'Corporate Discount') {
            let defaultimage = siteURL + "/_catalogs/masterpage/BloomHomepage/images/logo.png";
            for (let i = 0; i < objResults.length; i++) {
                if (objResults[i].Image != null) {
                    $('.img-bind' + i).prepend("<div class='list-imgsec'><img src='" + objResults[i].SiteLink.Url + "' alt='' title=''/></div>");
                } else {
                    $('.img-bind' + i).prepend("<div class='list-imgsec'><img src='" + defaultimage + "' alt='' title=''/></div>");
                }

                $('.item-title' + i).append(objResults[i].Title);
            }
        }
        let Viewevent = document.getElementsByClassName('viewitem');
        for (let i = 0; i < Viewevent.length; i++) {
            Viewevent[i].addEventListener("click", (e: Event) => this.viewitem(strLocalStorage));
        }
        let Editevent = document.getElementsByClassName('edititem');
        for (let i = 0; i < Editevent.length; i++) {
            Editevent[i].addEventListener("click", (e: Event) => this.edititem(strLocalStorage));
        }
        let eventfunction = document.getElementsByClassName('event-class');
        for (let i = 0; i < eventfunction.length; i++) {
            eventfunction[i].addEventListener("click", (e: Event) => this.eventfunction());
        }
        let holidayfunction = document.getElementsByClassName('holiday-class');
        for (let i = 0; i < holidayfunction.length; i++) {
            holidayfunction[i].addEventListener("click", (e: Event) => this.holidayfunction());
        }
        let addevent = document.getElementsByClassName('add-class');
        for (let i = 0; i < addevent.length; i++) {
            addevent[i].addEventListener("click", (e: Event) => this.addevent(strLocalStorage));
        }
        //Adding event for delete button click 
        let deleteevent = document.getElementById("deleteitems");
        //for (let i = 0; i < addevent.length; i++) {
        deleteevent.addEventListener("click", (e: Event) => this.deleteitems(strLocalStorage));
        //}

        $(document).on('change', '.switch-input', function (e) {
            var id = $(this).attr('id').replace('switch', '');
            var _thisid = $(this);

            if (strLocalStorage == 'Polls') {
                var strconfirm = "Do you want to display the selected poll?";
                alertify.confirm('Confirmation', strconfirm, function () {
                    $('.switch-input:checked').each(function () {
                        //var changedID=

                        var changeid = $(this).attr('id').replace('switch', '');
                        let myobj = {
                            Display: false
                        };
                        let item = updateItem(strLocalStorage, changeid, myobj);
                        item.then((items: any[]) => {
                            // console.log("Success2");
                        });
                        $('.switch-input').next().text("Not Displayed");
                        $(".switch-input").prop('checked', false);

                        var myobj1 = {
                            Display: true
                        };
                        let item1 = updateItem(strLocalStorage, id, myobj1);
                        item1.then((items: any[]) => {
                            // console.log("Success1");
                        });
                        _thisid.next().text("Displayed");
                        _thisid.prop('checked', true);
                        $('#deleteitem' + changeid).hide();

                    });
                }, function () {
                    if (_thisid.next().text("Displayed")) {
                        _thisid.prop('checked', true);

                    }
                }).set('closable', false);


            } else if (strLocalStorage == 'Announcements') {

                objResults = readItems(strLocalStorage, ["ID", "Title", "Modified", "Display"], 5000, "Modified", "Display", 1);
                objResults.then((items: any[]) => {
                    if (items.length < 3 && _thisid.prop("checked")) {
                        let myobj = {
                            Display: true
                        };
                        _thisid.next().text("Displayed");
                        _thisid.prop('checked', true);
                        let item = updateItem(strLocalStorage, id, myobj);
                        item.then((items: any[]) => {
                            // console.log("Success3");
                        });
                    } else if (items.length > 3 && _thisid.prop("checked")) {
                        //console.log("More than the count");
                        var strconfirm = "Please select maximum number 3 to be visible";
                        alertify.confirm('Confirmation', strconfirm, function () {
                            if (_thisid.prop("checked")) {
                                _thisid.next().text("Not Displayed");
                                _thisid.removeAttr('checked');
                                _thisid.prop('checked', false);
                            }
                            else {
                                _thisid.next().text("Displayed");
                                _thisid.prop('checked', true);
                            }
                        }, function () {
                            if (_thisid.prop("checked")) {
                                _thisid.next().text("Not Displayed");
                                _thisid.removeAttr('checked');
                                _thisid.prop('checked', false);
                            }
                            else {
                                _thisid.next().text("Displayed");
                                _thisid.prop('checked', true);
                            }
                        }).set('closable', false);
                    }
                    else {
                        let myobj = {
                            Display: false
                        };
                        _thisid.next().text("Not Displayed");
                        _thisid.removeAttr('checked');
                        let item = updateItem(strLocalStorage, id, myobj);
                        item.then((items: any[]) => {
                            // console.log("Success4");
                        });
                    }
                });

            }
            else if (strLocalStorage == 'Banners') {

                objResults = readItems(strLocalStorage, ["ID", "Title", "Modified", "Display"], 5000, "Modified", "Display", 1);
                objResults.then((items: any[]) => {
                    if (items.length < 3 && _thisid.prop("checked")) {
                        let myobj = {
                            Display: true
                        };
                        _thisid.next().text("Displayed");
                        _thisid.prop('checked', true);
                        let item = updateItem(strLocalStorage, id, myobj);
                        item.then((items: any[]) => {
                            // console.log("Success3");
                        });
                    } else if (items.length > 3 && _thisid.prop("checked")) {
                        //console.log("More than the count");
                        var strconfirm = "Please select maximum number 3 to be visible";
                        alertify.confirm('Confirmation', strconfirm, function () {
                            if (_thisid.prop("checked")) {
                                _thisid.next().text("Not Displayed");
                                _thisid.removeAttr('checked');
                                _thisid.prop('checked', false);
                            }
                            else {
                                _thisid.next().text("Displayed");
                                _thisid.prop('checked', true);
                            }
                        }, function () {
                            if (_thisid.prop("checked")) {
                                _thisid.next().text("Not Displayed");
                                _thisid.removeAttr('checked');
                                _thisid.prop('checked', false);
                            }
                            else {
                                _thisid.next().text("Displayed");
                                _thisid.prop('checked', true);
                            }
                        }).set('closable', false);
                    }
                    else {
                        let myobj = {
                            Display: false
                        };
                        _thisid.next().text("Not Displayed");
                        _thisid.removeAttr('checked');
                        let item = updateItem(strLocalStorage, id, myobj);
                        item.then((items: any[]) => {
                            // console.log("Success4");
                        });
                    }
                });

            }
            else if (strLocalStorage == 'Employee Corner') {

                objResults = readItems(strLocalStorage, ["ID", "Title", "Modified", "Display"], 5000, "Modified", "Display", 1);
                objResults.then((items: any[]) => {
                    if (items.length < 9 && _thisid.prop("checked")) {
                        let myobj = {
                            Display: true
                        };
                        _thisid.next().text("Displayed");
                        _thisid.prop('checked', true);
                        let item = updateItem(strLocalStorage, id, myobj);
                        item.then((items: any[]) => {
                            // console.log("Success3");
                        });
                    } else if (items.length >= 9 && _thisid.prop("checked")) {
                        //console.log("More than the count");
                        var strconfirm = "Please select maximum number 9 to be visible";
                        alertify.confirm('Confirmation', strconfirm, function () {
                            if (_thisid.prop("checked")) {
                                _thisid.next().text("Not Displayed");
                                _thisid.removeAttr('checked');
                                _thisid.prop('checked', false);
                            }
                            else {
                                _thisid.next().text("Displayed");
                                _thisid.prop('checked', true);
                            }
                        }, function () {
                            if (_thisid.prop("checked")) {
                                _thisid.next().text("Not Displayed");
                                _thisid.removeAttr('checked');
                                _thisid.prop('checked', false);
                            }
                            else {
                                _thisid.next().text("Displayed");
                                _thisid.prop('checked', true);
                            }
                        }).set('closable', false);
                    }
                    else {
                        let myobj = {
                            Display: false
                        };
                        _thisid.next().text("Not Displayed");
                        _thisid.removeAttr('checked');
                        let item = updateItem(strLocalStorage, id, myobj);
                        item.then((items: any[]) => {
                            // console.log("Success4");
                        });
                    }
                });

            }
            else if (strLocalStorage == 'Quick Links') {
                objResults = readItems(strLocalStorage, ["ID", "Title", "Modified", "Display"], 5000, "Modified", "Display", 1);
                objResults.then((items: any[]) => {
                    if (items.length < 18 && _thisid.prop("checked")) {
                        let myobj = {
                            Display: true
                        };
                        _thisid.next().text("Displayed");
                        _thisid.prop('checked', true);
                        let item = updateItem(strLocalStorage, id, myobj);
                        item.then((items: any[]) => {
                            // console.log("Success3");
                        });
                    } else if (items.length >= 18 && _thisid.prop("checked")) {
                        //console.log("More than the count");
                        var strconfirm = "Please select maximum number 18 to be visible";
                        alertify.confirm('Confirmation', strconfirm, function () {
                            if (_thisid.prop("checked")) {
                                _thisid.next().text("Not Displayed");
                                _thisid.removeAttr('checked');
                                _thisid.prop('checked', false);
                            }
                            else {
                                _thisid.next().text("Displayed");
                                _thisid.prop('checked', true);
                            }
                        }, function () {
                            if (_thisid.prop("checked")) {
                                _thisid.next().text("Not Displayed");
                                _thisid.removeAttr('checked');
                                _thisid.prop('checked', false);
                            }
                            else {
                                _thisid.next().text("Displayed");
                                _thisid.prop('checked', true);
                            }
                        }).set('closable', false);
                    } else {
                        let myobj = {
                            Display: false
                        };
                        _thisid.next().text("Not Displayed");
                        _thisid.removeAttr('checked');
                        let item = updateItem(strLocalStorage, id, myobj);
                        item.then((items: any[]) => {
                            // console.log("Success4");
                        });
                    }
                });

            }
            else {
                if (_thisid.prop("checked")) {
                    let myobj = {
                        Display: true
                    };
                    _thisid.next().text("Displayed");
                    _thisid.prop('checked', true);
                    let item = updateItem(strLocalStorage, id, myobj);
                    item.then((items: any[]) => {
                        // console.log("Success3");
                    });
                }
                else {
                    let myobj = {
                        Display: false
                    };
                    _thisid.next().text("Not Displayed");
                    _thisid.removeAttr('checked');
                    let item = updateItem(strLocalStorage, id, myobj);
                    item.then((items: any[]) => {
                        // console.log("Success4");
                    });
                }
            }
        });

        $('#pagination-list').jplist({
            itemsBox: '.list'
            , itemPath: '.list-item'
            , panelPath: '.jplist-panel'
        });
    }
    viewitem(strLocalStorage) {
        var siteURL = this.context.pageContext.web.absoluteUrl;
        $('div.viewitem').click(function () {
            var id = $(this).attr('id').replace('viewitem', '');
            window.location.href = "" + siteURL + "/Pages/Viewlistitem.aspx?CName=" + strLocalStorage + "&CID=" + $(this).attr('id').replace('viewitem', '');
        });
    }
    edititem(strLocalStorage) {
        var siteURL = this.context.pageContext.web.absoluteUrl;
        $('div.edititem').click(function () {
            var id = $(this).attr('id').replace('edititem', '');
            window.location.href = "" + siteURL + "/Pages/EditListItem.aspx?CName=" + strLocalStorage + "&CID=" + $(this).attr('id').replace('edititem', '');
        });
    }
    deleteitems(strLocalStorage) {
        // var strLocalStorage = GetQueryStringParams("CName").replace("%20", " ");
        var $body = $('body');
        var deleteitemID = [];
        $('.delete-item:checked').each(function () {
            deleteitemID.push($(this).attr('id').replace('deleteitem', ''));
        });
        if (deleteitemID.length > 0 && strLocalStorage == "Polls") {
            var strconfirm = "Are you sure you want to delete selected item(s)?";
            var _that = this;
            alertify.confirm('Confirmation', strconfirm, function () {
                $body.addClass("loading");
                var selectedArray: number[] = deleteitemID;
                for (let k = 0; k < deleteitemID.length; k++) {
                    var delID = Number(deleteitemID[k]);
                    pnp.sp.web.lists.getByTitle("Polls").items.getById(delID).update({
                        IsDeleted: true
                    }).then(function () {
                        pnp.sp.web.lists.getByTitle("PollsResults").items.filter("QuestionID eq '" + delID + "'").get().then((items: any[]) => {
                            if (items.length > 0) {
                                for (var i = 0; i < items.length; i++) {
                                    pnp.sp.web.lists.getByTitle("PollsResults").items.getById(items[i].ID).update({
                                        IsDeleted: true
                                    })
                                }
                                // 
                                // $body.removeClass("loading");
                            }
                        });
                    })


                }
                setTimeout(function () {
                    location.reload();
                }, 1000);
                setTimeout(function () {
                    $body.removeClass("loading");

                }, 5000);
                //batchDelete(strLocalStorage, selectedArray, _that.context.pageContext.web.absoluteUrl);
            },
                function () {
                    $body.removeClass("loading");
                }).set('closable', false);

        }
        else if (deleteitemID.length > 0) {
            var strconfirm = "Are you sure you want to delete selected item(s)?";
            var _that = this;
            alertify.confirm('Confirmation', strconfirm, function () {

                $body.addClass("loading");
                var selectedArray: number[] = deleteitemID;
                //for (var i = 0; i < selectedArray.length; i++) {
                batchDelete(strLocalStorage, selectedArray, _that.context.pageContext.web.absoluteUrl);
            }
                , function () { $body.removeClass("loading"); }).set('closable', false);
        }
        else {
            alertify.set('notifier', 'position', 'top-right');
            alertify.error('Please select at least one item');
        }
    }

    eventfunction() {
        $(".content-area").empty();
        $(".list-tab").remove();
        this.ViewListItems("Events");
    }
    holidayfunction() {
        $(".content-area").empty();
        $(".list-tab").remove();
        this.ViewListItems("Holiday");
    }
    addevent(strLocalStorage) {

        var siteURL = this.context.pageContext.web.absoluteUrl;
        window.location.href = "" + siteURL + "/Pages/AddListItem.aspx?CName=" + strLocalStorage;
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
