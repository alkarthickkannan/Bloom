import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './UserProfileWebPart.module.scss';
import * as strings from 'UserProfileWebPartStrings';
import pnp from 'sp-pnp-js';
import 'jquery';
export interface IUserProfileWebPartProps {
  description: string;
}
declare var $;
declare var alertify: any;

export interface IUserProfileWebPartProps {
  description: string;
}

var arrDocument = [];
export default class UserProfileWebPart extends BaseClientSideWebPart<IUserProfileWebPartProps> {

  public render(): void {
    this.domElement.innerHTML = `
    <div class="modal fade" id="AddUploadmodal" tabindex="-1" role="dialog" aria-labelledby="basicModal" aria-hidden="true">
    <div class="modal-dialog modal-md">
      <div class="modal-content">
        <div class="modal-header">
          <h4 class="modal-title" id="myDocModalLabel">Document Upload</h4>
          <button type="button" class="close" data-dismiss="modal" aria-label="Close"> <span class="icon-remove"></span> </button>
        </div>
        <div class="modal-body">
          <div class="col-xs-12 form-element">
            <label class="required">Title</label>
            <input type="text" id="txtDocTitle" placeholder="Title of the Document" class="form-control">
          </div>
          <div class="col-xs-12 form-element">
          <label class="required">Library</label>
          <select id="ddlDocLibrary" class='form-control'>
          
          </select>
          </div>
          <div class="col-xs-12 form-element" id="divUploadDoc">
          <div class="custom-upload banner-upload">
          <label class='control-label required'>Document File</label>
          <input type='file' id='uploadDocFile' title="" name='file' accept='.doc,.docx,.xls,.ppt,.pdf' multiple='' class='file'>
          <div class='input-group'>
          <span class='input-group-btn input-group-sm'>
          <button type='button' class='btn btn-fab btn-fab-mini'>Browse</button>
          </span>
          <input type='text' readonly='' id='Uploadedtxt' class='form-control' placeholder='Upload Files'>
          </div>
          </div>
          </div>
        </div>
          <div class="modal-footer">
            <div class="col-xs-12 form-element"> 
            <a id="btnDocAddSubmit" href="#" class="s-button">Submit</a><label id="lblwait" style="display:none;float:left;">Please Wait...</label>
            </div>
          </div>
        </div>
      </div>
    </div>
    <section class="user-profile">
      <div class="user-view user-detail">
      </div>
      <div class="user-view user-view1">
      <a href="" data-toggle="modal" data-target="#AddUploadmodal"><i class="icon-add"></i>Upload</a>
      <p class="center-p">(Upload your files)</p>
      </div>
    </section>
    <div class='modal-loader-cls'><!-- Place at bottom of page --></div>`;

    this.getuserdetails();
    this.getDocuments();


    let Submitevent = $('#btnDocAddSubmit');
    Submitevent.on("click", (e: Event) => this.addDocuments());

    $("#btnDocAddSubmit").hover(function() {
      $(this).css("background-color","#E42313")
    });

    $("input[type=file]").change(function(){
      $('#Uploadedtxt').val($('#uploadDocFile').val().replace(/C:\\fakepath\\/i, ''));
    });

    var _thiss = this;
    $("#ddlDocLibrary").change(function(){
      _thiss.getColumns($(this).val());
    });

  }

  async getDocuments(){

    pnp.sp.site.getDocumentLibraries(this.context.pageContext.web.absoluteUrl).then(function(data) {
      var strTitle = "";
      strTitle += '<option id="select" disabled selected>select</option>';
      for (var i = 0; i < data.length; i++) {
        // arrDocument.push({
        //   Title:data[i].Title,
        //   URL:data[i].ServerRelativeUrl
        // })
        var stringlen = data[i].AbsoluteUrl.split('/');   
        var Lib = stringlen[stringlen.length - 1]; 
        //console.log(stringlen);

        if(Lib != "SharedDocuments")
        {
          strTitle += "<option id='"+Lib+"'>"+Lib+"</option>";  
        }
      }
      $('#ddlDocLibrary').html(strTitle);
      
     }).catch(function(err) {
      alert(err);
     });

  }

   addDocuments(){
    var $body = $('body');
    if ($('.ajs-message').length > 0) {
      $('.ajs-message').remove();
    }
    var isAllfield = true;
    var dynamicvalidation = true;

    if (!$('#txtDocTitle').val().trim()) {
      alertify.set('notifier', 'position', 'top-right');
      alertify.error("Please enter the Title");
      isAllfield = false;
    } 
    else if ($('#ddlDocLibrary').find(":selected").text() == "select") {
      alertify.set('notifier', 'position', 'top-right');
      alertify.error("Please select the Document Library");
      isAllfield = false;
    } 
    else if (!$('#uploadDocFile').val().trim()) {
      alertify.set('notifier', 'position', 'top-right');
      alertify.error("Please Choose the File");
      isAllfield = false;
    }
    else{
      for(var i=0; i<arrDocument.length; i++)
      {
        if(!$('#txt'+arrDocument[i]["Title"]).val().trim()){
          alertify.set('notifier', 'position', 'top-right');
          alertify.error("Please Enter the "+ arrDocument[i]["Title"]);
          isAllfield = false;
          return false;
        }
      }
    }

    if(isAllfield){
    $('#btnDocAddSubmit').hide();
    $('#lblwait').show();
    // $('#btnDocAddSubmit').css('disabled','disabled');
    var files = <HTMLInputElement>document.getElementById("uploadDocFile");
    let file = files.files[0];

      var VideoTitle = { User: $('#txtDocTitle').val().trim() };

      var Json = [];
      var item = {};
      for(var i=0; i<arrDocument.length; i++)
      {
        if(arrDocument[i]["TypeDisplayName"] == "Single line of text")
        {
          item[arrDocument[i]["Title"]]=$('#txt'+arrDocument[i]["Title"]).val().trim();
        }
        else if(arrDocument[i]["TypeDisplayName"] == "Hyperlink or Picture"){
          var link = {
                      "__metadata": {
                        "type": "SP.FieldUrlValue"
                      },
                      Url: $('#txt'+arrDocument[i]["Title"]).val().trim(),
                    }
          item[arrDocument[i]["Title"]] = link;
        }
      }
      Json.push(item);
      var Vdotile = Json[0];
      $('body').addClass("loading");
      pnp.sp.web.getFolderByServerRelativeUrl($('#ddlDocLibrary').find(":selected").text()).files.add(file.name, file, true)
        .then(function(result) {

            console.log(file.name + " upload successfully!");
            result.file.listItemAllFields.get().then((listItemAllFields) => {
              pnp.sp.web.lists.getByTitle($('#ddlDocLibrary').find(":selected").text()).items
              .getById(listItemAllFields.Id).update(Vdotile).then(r=>{
                    alertify.set('notifier', 'position', 'top-right');
                    alertify.success("Document Uploaded Successfully");
                    window.location.reload();
              });
            });

      });
      $('body').removeClass("loading");
    }
  }


  async getColumns(DocumentLib){
    $('#divUploadDoc').nextAll().remove();
    arrDocument = [];
    $.ajax({
      url: this.context.pageContext.web.absoluteUrl + "/_api/web/lists/getbytitle('"+DocumentLib+"')/fields?select=Title,Name&$filter=Hidden eq false and ReadOnlyField eq false and Required eq true",
      type: "GET",
      headers: { Accept: "application/json;odata=verbose" },  
      success: function (Columndata) {
        var DymHtml = "";
        for(var i=0; i<Columndata.d.results.length; i++)
        {
          if(Columndata.d.results[i].Title != "Name")
          {
            DymHtml += '<div class="col-xs-12 form-element">'+
                        '<label class="required">'+Columndata.d.results[i].Title+'</label>'+
                        '<input type="text" id="txt'+Columndata.d.results[i].Title+'" placeholder="'+Columndata.d.results[i].Title+'" class="form-control">'+
                      '</div>';
            arrDocument.push({
              "Title": Columndata.d.results[i].Title,
              "TypeDisplayName": Columndata.d.results[i].TypeDisplayName
            });
          }
        }
        $('#divUploadDoc').after(DymHtml);
        
      },
      error: function (data) {
      console.log(data);
      },
      });
  }


  public getuserdetails() {
    pnp.sp.profiles.myProperties.get().then(result => {
      var props = result.UserProfileProperties;
      var propValue = {};
      props.forEach(function (prop) {
        if (typeof prop.Value === undefined || prop.Value == "" || prop.Value == "undefined") {
          propValue[prop.Key] = "Not Available";
        }
        else {
          propValue[prop.Key] = prop.Value;
        }
      });
      this.renderhtml(propValue);
      // console.log(propValue);
    });
  }


  public renderhtml(objResults) {
    var url = objResults["PictureURL"];
    var Email = objResults["WorkEmail"].length;
    if(objResults["WorkEmail"].length > 24)
    {
      Email = objResults["WorkEmail"].substring(0,24)+"...";
    }
    var pathname = new URL(url).origin + "/person.aspx";
    //console.log(pathname);
    var renderhtml = "";
    renderhtml += "<div align='center'>" +
      "<img src='" + objResults["PictureURL"] + "'>" +
      "</div>" +
      "<h3>" + objResults["FirstName"] + " " + objResults["LastName"] + "</h3>" +
      "<p class='pad-left0'>" + objResults["Department"] + "</p>" +
      "<p class='p-space' title='"+objResults["WorkEmail"]+"'><i class='icon-mail'></i>" + Email + "</p>" +
      "<p><i class='icon-phone'></i>" + objResults["WorkPhone"] + "</p>" +
      "<a href='" + pathname + "'>View Profile</a>";
    $('.user-detail').append(renderhtml);
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
