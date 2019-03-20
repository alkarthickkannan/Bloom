import { Version } from '@microsoft/sp-core-library';

import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './EdmsProjectHomeWebPart.module.scss';
import * as strings from 'EdmsProjectHomeWebPartStrings';

import { addItems, readItems, deleteItem, updateItem,additemsattachment,batchDelete  } from '../../commonJS';

export interface IEdmsProjectHomeWebPartProps {
  description: string;
}

export default class EdmsProjectHomeWebPart extends BaseClientSideWebPart<IEdmsProjectHomeWebPartProps> {

  public render(): void {
    this.domElement.innerHTML = `
    <section class="gallery-section">
    </section>`;
    this.FetchItems();
  }

  async FetchItems(){
    var listName = "Tiles";
    let columnArray = ["Name","LinkURL"];
    let GetListItems = await readItems(listName, columnArray, 1, "Modified");
    var HTML = "";
    if(GetListItems.length > 0)
    {
      for(var i=0; i<GetListItems.length; i++)
      {
        if(i==0)
        {
          HTML += '<div class="col-sm-3 col-xs-12  pad-left0">'+
                  '<div class="img-gallery"> <img src="images/announce-listimg1.jpg"> <a href="'+GetListItems[i].LinkURL.Url+'" data-toggle="modal" data-target="#Addmodal">'+GetListItems[i].Name+'<i class="icon-more"></i></a> </div>'+
                '</div>';
        }
        else if(i==GetListItems.length-1)
        {
          HTML += '<div class="col-sm-3 col-xs-12  over-right">'+
                  '<div class="img-gallery"> <img src="images/announce-listimg1.jpg"> <a href="'+GetListItems[i].LinkURL.Url+'" data-toggle="modal" data-target="#Addmodal">'+GetListItems[i].Name+'<i class="icon-more"></i></a> </div>'+
                '</div>';
        }
        else{
          HTML += '<div class="col-sm-3 col-xs-12">'+
                    '<div class="img-gallery"> <img src="images/announce-listimg1.jpg"> <a href="'+GetListItems[i].LinkURL.Url+'" data-toggle="modal" data-target="#Addmodal">'+GetListItems[i].Name+'<i class="icon-more"></i></a> </div>'+
                  '</div>';
        }
      }
      $('#gallery-section').append(HTML);
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
