import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import * as $ from "jquery";
import styles from './EdmsDepartmentDocUploadWebPart.module.scss';
import * as strings from 'EdmsDepartmentDocUploadWebPartStrings';
import { SPComponentLoader } from '@microsoft/sp-loader';
import pnp from 'sp-pnp-js';

//declare var $;

export interface IEdmsDepartmentDocUploadWebPartProps {
  description: string;
}
var HTML = "";

export default class EdmsDepartmentDocUploadWebPart extends BaseClientSideWebPart<IEdmsDepartmentDocUploadWebPartProps> {

  public render(): void {
    this.domElement.innerHTML = `
      <div class="doc-upload">  <h3 class="tt-head"> Documents Uploaded<a href="#">More</a></h3>
        <ul id="UlDocumentUpload">
        </ul>
      </div>`;
    
    this.GetDocuments();
  }
  async GetDocuments(){    
   
    let read = await pnp.sp.web.lists.getByTitle("DocumentUpload").items.select("DocumentLink","Title","Modified").filter("User eq '" + this.context.pageContext.user.displayName + "'").top(3).orderBy("Modified",false).get()
    if(read){
      for(var i=0; i<3; i++)
      {
        if(read[i] != undefined && read[i] != null)
        {
          HTML += '<li><a href="'+read[i].DocumentLink.Url+'"><img src="images/pdf.png">'+read[i].Title+'</a></li>';
        }
      }
      $('#UlDocumentUpload').html(HTML);
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
