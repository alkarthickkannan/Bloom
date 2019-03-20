import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';
import 'jquery';
import pnp from 'sp-pnp-js';
//import styles from './OrganizationPolicyWebPart.module.scss';
import * as strings from 'OrganizationPolicyWebPartStrings';
require('jplist-core');
require('jplist-pagination');
require('../../ExternalRef/js/jplist-core.js');
require('../../ExternalRef/js/jplist-pagination.js');
require('../../ExternalRef/js/bootstrap-select.min.js');
import {
  readItems, updateItem, formatDate, checkUserinGroup, GetQueryStringParams, batchDelete
} from '../../commonJS';
declare var $;
declare var alertify: any;
export interface IOrganizationPolicyWebPartProps {
  description: string;
}

export default class OrganizationPolicyWebPart extends BaseClientSideWebPart<IOrganizationPolicyWebPartProps> {
  userflag: boolean = false;
  public render(): void {
    var _thatt = this;
    checkUserinGroup("Organizational Policies", this.context.pageContext.user.email, function (result) {
      if (result == 1) {
        _thatt.userflag = true;
        _thatt.loadcomponent();
      }else{
        _thatt.userflag = false;
        _thatt.loadcomponent();
      }
    });
document.title = "Organizational Policies";
  }

  public loadcomponent(){
    var siteURL = this.context.pageContext.web.absoluteUrl;
    var strLocalStorage = GetQueryStringParams("CName").replace("%20", " ");
    this.domElement.innerHTML =
    "<div class='breadcrumb'>" +
    "<ol>" +
    "<li><a href='" + siteURL + "/Pages/Home.aspx' title='Home' class='pointer'>Home</a></li>" +
    "<li><span>Organizational Policies</span></li>" +
    "</ol>" +
    "</div>" +
    "<div class='title-section'>" +
    "<div class='button-field'>" +
    "<a href='" + siteURL + "/Pages/AddListItem.aspx?CName=Organizational%20Policies' title='Add New' class='pointer' id='AddingButtons'><i class='icon-add'></i>Add New</a>" +
    "<a class='delete-icon DeletingButtons pointer' title='Delete' id='DeletingButtons'><img src='" + siteURL + "/_catalogs/masterpage/BloomHomepage/images/close-icon.png'>Delete</a>" +
    "</div>" +
    "<h2>Organizational Policies</h2>" +
    "</div>" +
  
    "</div></div>" +
    "<div class='content-area'>" +


    "</div>" +
    "<div class='modal'><!-- Place at bottom of page --></div>";
    
    this.OrgPage();

  }
  async OrgPage() {
    var Organ = "<div id='pagination-list' class='list-section jplist'><ul class='list'>";
    var count = 50;
    var checkboxstatus = "";
    var strcheckboxstatus = "Not Displayed";
    if (this.userflag == false) {      
      $('.button-field').hide();      
    }
    var dept = "Admin";
    dept = GetQueryStringParams("CName").replace("%20", " ");
    //if ($("#OptionsforSel").val() != null) {
    // dept = "" + $("#OptionsforSel").val() + "";
    // }
    var strLocalStorage = "";
    if (strLocalStorage == "") {
      strLocalStorage = "Organizational%20Policies";
    }
    var items;
    //objResults= readItems(strLocalStorage, ["ID", "Title", "Modified", "DocumentFile", "Explanation", "Display", "Departments"], count, "Modified", "Departments", dept)
    //pnp.sp.web.lists.getByTitle("OrganizationPolicy").items.filter("Departments eq "+dept+" and Display eq 1").get().then((items: any[]) => {
    if (this.userflag == false) {
      items = await pnp.sp.web.lists.getByTitle(strLocalStorage).items.filter("Departments eq '" + dept + "' and Display eq 1").top(count).orderBy("Modified").get();
    }
    else {
      items = await pnp.sp.web.lists.getByTitle(strLocalStorage).items.filter("Departments eq '" + dept + "'").top(count).orderBy("Modified").get();
    }
    if (items.length > 0) {
      //objResults.then((items: any[]) => {
      for (let i = 0; i < items.length; i++) {
        if (items[i].Display == "1") {
          checkboxstatus = "checked";
          strcheckboxstatus = "Displayed";
        }
        else {
          checkboxstatus = "";
          strcheckboxstatus = "Not Displayed";
        }
        if (items[i].Explanation != null && items[i].Explanation > 160) {
          items[i].Explanation = items[i].Explanation.substring(0, 160) + "...";
        }
        else if (items[i].Explanation == null) {
          items[i].Explanation = "";
        }
        Organ += "<li class='list-item'>" +
          "<div class='list-imgcont' >" +
          "<p class='Modified'><strong>" + formatDate(items[i].Modified) + "</strong></p>" +
          "<a href='" + items[i].DocumentFile.Url + "' target='_blank'><h3 class='OrgTitle'>" + items[i].Title + "</h3></a>" +
         // "<p class='OrgDescrip'>" + items[i].Explanation + "</p>" +
          "<p class='OrgDepart'>" + items[i].Departments + "</p>" +
          "<div class='switch'>" +
          "<input type='checkbox' id='switch" + items[i].ID + "' class='switch-input sndswitch' " + checkboxstatus + "/>" +
          "<label for='switch" + items[i].ID + "' class='switch-label sndswitch'>" + strcheckboxstatus + "</label>" +
          "<div class='list-icons'>" +
          "<div class='icon-list2 viewitem'>" +
          "<a  title='View' class='viewitem pointer' id='viewitem" + items[i].ID + "'><i class='icon-eye viewitem'></i></a>" +
          "</div>" +
          "<div class='icon-list2 edititemuser edititem'>" +
          "<a  title='Edit' class='edititem pointer' id='edititem" + items[i].ID + "'><i class='icon-edit edititem'></i></a>" +
          "</div>" +
          "<div class='icon-list2 deleteitemuser'>" +
          "<div class='check-box'>" +
          "<input type='checkbox'  name='' value='' class='delete-item' id='deleteitem" + items[i].ID + "'/>" +
          "<label>Checkbox</label>" +
          "</div>" +
          "</div>" +
          "</div>" +
          "</div>" +
          "</li>";
      }
    }
    else {
      Organ += "<li class='list-item'>No items to display" +
        "</li>";
    }
    Organ += "</ul>";
    Organ += "<div class='jplist-panel box panel-top'>" +
      "<div class='jplist-pagination' data-control-type='pagination' data-control-name='paging' data-control-action='paging'></div>" +
      "<select class='jplist-select' data-control-type='items-per-page-select' data-control-name='paging' data-control-action='paging'>" +
      "<option data-number='5' data-default='true'> 5 </option>" +
      "<option data-number='10'> 10 </option>" +
      "<option data-number='15'> 15 </option>" +
      "</select>" +
      "</div>";


    $('.content-area').append(Organ);
    let Viewevent = document.getElementsByClassName('viewitem');
    for (let i = 0; i < Viewevent.length; i++) {
      Viewevent[i].addEventListener("click", (e: Event) => this.viewitem());
    }
    let Editevent = document.getElementsByClassName('edititem');
    for (let i = 0; i < Editevent.length; i++) {
      Editevent[i].addEventListener("click", (e: Event) => this.edititem());
    }
    //Adding event for delete button click 
    let deleteevent = document.getElementById("DeletingButtons");

    deleteevent.addEventListener("click", (e: Event) => this.deleteitems());
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
        let item = updateItem("Organizational%20Policies", id, myobj);
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
        let item = updateItem("Organizational%20Policies", id, myobj);
        item.then((items: any[]) => {
          //console.log("Success update false");
        });
      }
    });


  }
  public Depart() {
    var strLocalStorage = GetQueryStringParams("CName");
    if (strLocalStorage === undefined) {
      strLocalStorage = "Organizational%20Policies";
    }
    let DepName = "";
    pnp.sp.web.lists.getByTitle(strLocalStorage).items.get()
      .then((items: any) => {
        if (items.length > 0) {
          var flags = [], output = [], l = items.length, i;
          for (i = 0; i < l; i++) {
            if (flags[items[i].Departments]) continue;
            flags[items[i].Departments] = true;
            output.push(items[i].Departments);
          }

          for (var k = 0; k < output.length; k++) {
            DepName += "<option value='" + output[k] + "'>" + output[k] + "</option>";
          }

          $('.selectpicker').append(DepName);
          $('.selectpicker').selectpicker();
        }
      }).then(r => {
        this.OrgPage();
      });
  }
  public viewitem() {
    var siteURL = this.context.pageContext.web.absoluteUrl;
    $('a.viewitem').click(function () {
      var id = $(this).attr('id');
      window.location.href = "" + siteURL + "/Pages/Viewlistitem.aspx?CName=Organizational%20Policies&CID=" + $(this).attr('id').replace('viewitem', '') + "&CMode=ViewMode";
    });
  }
  public edititem() {
    var siteURL = this.context.pageContext.web.absoluteUrl;
    $('a.edititem').click(function () {
      var id = $(this).attr('id');

      window.location.href = "" + siteURL + "/Pages/EditListItem.aspx?CName=Organizational%20Policies&CID=" + $(this).attr('id').replace('edititem', '') + "&CMode=EditMode";
    });
  }
  public deleteitems() {
    var strLocalStorage = "Organizational%20Policies";
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
