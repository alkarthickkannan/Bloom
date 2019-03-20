import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';
import 'jquery';
//import styles from './CorporateDiscountWebPart.module.scss';
import * as strings from 'CorporateDiscountWebPartStrings';
require('jplist-core');
require('jplist-pagination');
require('../../ExternalRef/js/jplist-core.js');
require('../../ExternalRef/js/jplist-pagination.js');
import {
  readItems, updateItem, formatDate, checkUserinGroup, batchDelete
} from '../../commonJS';
import { Search } from 'sp-pnp-js';
declare var $;
import pnp from 'sp-pnp-js';
import { Web } from 'sp-pnp-js';
declare var alertify: any;
export interface ICorporateDiscountWebPartProps {
  description: string;
}

export default class CorporateDiscountWebPart extends BaseClientSideWebPart<ICorporateDiscountWebPartProps> {
  userflag: boolean = false;
  public render(): void {
    var siteURL = this.context.pageContext.web.absoluteUrl;
    var _thatt = this;
    checkUserinGroup("Corporate Discounts", this.context.pageContext.user.email, function (result) {

      if (result == 1) {
        _thatt.userflag = true;
        _thatt.loadcomponent();
      } else {
        _thatt.userflag = false;
        _thatt.loadcomponent();
      }

    });


  }
  public loadcomponent() {
    var siteURL = this.context.pageContext.web.absoluteUrl;
    this.domElement.innerHTML =
      "<div class='breadcrumb bread-pos'>" +
      "<ol>" +
      "<li><a href='" + siteURL + "/Pages/Home.aspx' title='Home' class='pointer'>Home</a></li>" +
      "<li><span>Corporate Discounts</span></li>" +
      "</ol>" +
      "<div class='input search'>" +
      "<input id='customSearch' class='CorporateDiscountsearch form-control' type='text' placeholder='Search..' name='search'>" +
      "<a id='corporateSearch' class='close-searchicon pointer' title='search'>" +
      "<i class='icon-search' style='float:right; margin: -20px 10px 10px 5px;'></i></a></div>" +
      "</div>" +
      "<div class='title-section'>" +
      "<div class='button-field'>" +
      "<a href='" + siteURL + "/Pages/AddListItem.aspx?CName=Corporate%20Discounts' title='Add New' class='pointer' id='AddingButtons'><i class='icon-add'></i>Add New</a>" +
      "<a class='delete-icon pointer' title='Delete' id='DeletingButtons'><img src='" + siteURL + "/_catalogs/masterpage/BloomHomepage/images/close-icon.png'>Delete</a>" +
      "</div>" +
      "<h2>Corporate Discounts</h2>" +
      "</div>" +
      "<div class='content-area'>" +
      "</div>" +
      "<div class='modal'><!-- Place at bottom of page --></div>";
    this.CorDis(null);

    let customsearchevent = document.getElementById('corporateSearch');
    //for (let i = 0; i < Addevent.length; i++) {
    var _globalthis = this;

    customsearchevent.addEventListener("click", (e: Event) => this.corporateSearch());

    document.title = "Corporate Discounts";
    $(document).keypress(function (event) {


      var keycode = event.which || event.keyCode || event.charCode;
      if (keycode == '13') {
        if ($('.ajs-message').length > 0) {
          $('.ajs-message').remove();
        }
        let isSearch: boolean = true;
        if (!$('#customSearch').val().trim()) {
          alertify.set('notifier', 'position', 'top-right');
          alertify.error("Please Enter the Value");
          isSearch = false;
          //isAllfield = false;
        } if (isSearch) {
          var searchvalue = $('#customSearch').val().trim();
          _globalthis.CorDis(searchvalue);
        }
      }

    });
    $(document).on('keypress', function () {

    }).on('keydown', function (e) {

      if (e.keyCode == 8 && !$('#customSearch').val().substring(1)) {
        location.reload()
      }

    });

  }


  corporateSearch() {
    if ($('.ajs-message').length > 0) {
      $('.ajs-message').remove();
    }
    let isSearch: boolean = true;
    if (!$('#customSearch').val().trim()) {
      alertify.set('notifier', 'position', 'top-right');
      alertify.error("Please Enter the Value");
      isSearch = false;

    } if (isSearch) {
      var searchvalue = $('#customSearch').val().trim();
      this.CorDis(searchvalue);
    }
  }


  public CorDis(searchText) {
    //var searchvalue = "Testing";
    var Corporate = "<div id='pagination-list' class='list-section jplist'><ul class='list'>";
    var count = 50;
    var checkboxstatus = "";
    var strcheckboxstatus = "Not Displayed";
    let objResults;

    if (searchText && this.userflag) {
      objResults = pnp.sp.web.lists.getByTitle("Corporate%20Discounts").items.select("ID", "Title", "Modified", "VendorLogo", "SiteLink", "Display").filter(("Title eq '" + searchText + "'") || "SiteLink eq '" + searchText + "'").top(100).get();
      //('Title eq '+searchText+' or SiteLink eq '+searchText+'')
    } else if (searchText && this.userflag == false) {
      // objResults = pnp.sp.web.lists.getByTitle("Corporate%20Discounts").items.select("ID", "Title", "Modified", "VendorLogo", "SiteLink", "Display").filter(("Title eq '" + searchText + "'") || "SiteLink eq '" + searchText + "'").top(100).get();

      objResults = pnp.sp.web.lists.getByTitle("Corporate%20Discounts").items.select("ID", "Title", "Modified", "VendorLogo", "SiteLink", "Display").filter("(Display eq '1' and Title eq '" + searchText + "' )" || "SiteLink eq '" + searchText + "'").top(100).get();
    }
    else {

      if (this.userflag == false) {
        objResults = readItems("Corporate%20Discounts", ["ID", "Title", "Modified", "VendorLogo", "SiteLink", "Display"], count, "Modified", "Display", 1);
      }
      else {
        objResults = readItems("Corporate%20Discounts", ["ID", "Title", "Modified", "VendorLogo", "SiteLink", "Display"], count, "Modified");
      }
    }

    objResults.then((items: any[]) => {

      $('.content-area').empty();
      if (items.length > 0) {

        for (let i = 0; i < items.length; i++) {
          if (items[i].Display == "1") {
            checkboxstatus = "checked";
            strcheckboxstatus = "Displayed";
          }
          else {
            checkboxstatus = "";
            strcheckboxstatus = "Not Displayed";
          }
          Corporate += "<li class='list-item'>" +
            "<div class='list-imgcont'>" +
            "<div class='list-imgsec list-imgsec" + i + "'>" +
            "</div>" +
            "<p class='Modified'><strong>" + formatDate(items[i].Modified) + "</strong></p>" +
            "<h3 class='CorTitle'>" + items[i].Title + "</h3>" +
            "<div class='switch'>" +
            "<input type='checkbox' id='switch" + items[i].ID + "' class='switch-input sndswitch' " + checkboxstatus + "/>" +
            "<label for='switch" + items[i].ID + "' class='switch-label sndswitch'>" + strcheckboxstatus + "</label>" +
            "<div class='list-icons'>" +
            "<div class='icon-list2 viewitem'>" +
            "<a  title='View' class='viewitem pointer'  id='viewitem" + items[i].ID + "'><i class='icon-eye viewitem'></i></a>" +
            "</div>" +
            "<div class='icon-list2 edititemuser edititem'>" +
            "<a  title='Edit' class='edititem pointer' id='edititem" + items[i].ID + "'><i class='icon-edit edititem' ></i></a>" +
            "</div>" +
            "<div class='icon-list2 deleteitemuser'>" +
            "<div class='check-box'>" +
            "<input type='checkbox' name='' value='' class='delete-item' id='deleteitem" + items[i].ID + "'/>" +
            "<label>Checkbox</label>" +
            "</div>" +
            "</div>" +
            "</div>" +
            "</div>" +
            "</div>" +
            "</li>";
        }
      }
      else {
        Corporate += "<li class='list-item'>No items to display" +
          "</li>";
      }
      Corporate += "</ul>";
      Corporate += "<div class='jplist-panel box panel-top'>" +
        "<div class='jplist-pagination' data-control-type='pagination' data-control-name='paging' data-control-action='paging'></div>" +
        "<select class='jplist-select' data-control-type='items-per-page-select' data-control-name='paging' data-control-action='paging'>" +
        "<option data-number='5' data-default='true'> 5 </option>" +
        "<option data-number='10'> 10 </option>" +
        "<option data-number='15'> 15 </option>" +
        "</select>" +
        "</div>";

      $('.content-area').append(Corporate);
      for (let i = 0; i < items.length; i++) {
        if (items[i].VendorLogo != null) {
          $('.list-imgsec' + i).append("<img src='" + items[i].VendorLogo.Url + "' alt='' title=''>");
        } else {
          var siteURL = this.context.pageContext.web.absoluteUrl;
          let defaultimage = siteURL + "/_catalogs/masterpage/BloomHomepage/images/logo.png";
          $('.list-imgsec' + i).append("<img src='" + defaultimage + "' alt='' title=''>");
        }
      }
      if (this.userflag == false) {
        $('.edititemuser').hide();
        $('.deleteitemuser').hide();
        $('.sndswitch').hide();
        $('.button-field').hide();
        $('.viewitem').show();
      }
      else {
        $('.edititemuser').show();
        $('.deleteitemuser').show();
      }
      let Viewevent = document.getElementsByClassName('viewitem');
      for (let i = 0; i < Viewevent.length; i++) {
        Viewevent[i].addEventListener("click", (e: Event) => this.viewitem());
      }
      let Editevent = document.getElementsByClassName('edititem');
      for (let i = 0; i < Editevent.length; i++) {
        Editevent[i].addEventListener("click", (e: Event) => this.edititem());
      }
      let deleteevent = document.getElementById("DeletingButtons");
      deleteevent.addEventListener("click", (e: Event) => this.deleteitems());

      $('#pagination-list').jplist({
        itemsBox: '.list',
        itemPath: '.list-item',
        panelPath: '.jplist-panel'
      });
      $(document).on('change', '.switch-input', function () {
        var id = $(this).attr('id').replace('switch', '');
        var _thisid = $(this);

        if (_thisid.prop("checked")) {
          let myobj = {
            Display: true
          };
          _thisid.next().text("Displayed");
          _thisid.attr("checked", "checked");
          let item = updateItem("Corporate%20Discounts", id, myobj);
          item.then((items: any[]) => {
            //console.log("Success update true");
          });
        }
        else {
          let myobj = {
            Display: false
          };
          _thisid.next().text("Not Displayed");
          _thisid.removeAttr('checked');
          let item = updateItem("Corporate%20Discounts", id, myobj);
          item.then((items: any[]) => {
            // console.log("Success update false");
          });
        }
      });
    });

  }
  public viewitem() {
    var siteURL = this.context.pageContext.web.absoluteUrl;
    $('a.viewitem').click(function () {
      var id = $(this).attr('id');
      window.location.href = "" + siteURL + "/Pages/Viewlistitem.aspx?CName=Corporate%20Discounts&CID=" + $(this).attr('id').replace('viewitem', '') + "&CMode=ViewMode";
    });
  }
  public edititem() {
    var siteURL = this.context.pageContext.web.absoluteUrl;
    $('a.edititem').click(function () {
      var id = $(this).attr('id');
      window.location.href = "" + siteURL + "/Pages/EditListItem.aspx?CName=Corporate%20Discounts&CID=" + $(this).attr('id').replace('edititem', '') + "&CMode=EditMode";
    });
  }
  public deleteitems() {
    var strLocalStorage = "Corporate Discounts";
    var deleteitemID = [];
    var $body = $('body');
    $('.delete-item:checked').each(function () {
      deleteitemID.push($(this).attr('id').replace('deleteitem', ''));
    });
    if (deleteitemID.length > 0) {
      var strconfirm = "Are you sure you want to delete selected item(s)?";
      var _that = this;
      alertify.confirm('Confirmation', strconfirm, function () {
        var selectedArray: number[] = deleteitemID;
        $body.addClass("loading");
        //for (var i = 0; i < selectedArray.length; i++) {
        batchDelete(strLocalStorage, selectedArray, _that.context.pageContext.web.absoluteUrl);
        //}
        //location.reload();
      }
        , function () { }).set('closable', false);
    }
    else {
      alertify.set('notifier', 'position', 'top-right');
      alertify.error('Please select at least one item');
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
