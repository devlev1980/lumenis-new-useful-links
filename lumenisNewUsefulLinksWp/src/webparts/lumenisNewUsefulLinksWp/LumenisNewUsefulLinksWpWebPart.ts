import styles from './LumenisNewUsefulLinksWpWebPart.module.scss';
import * as strings from 'LumenisNewUsefulLinksWpWebPartStrings';

import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import { SPComponentLoader } from '@microsoft/sp-loader';

import * as pnp from 'sp-pnp-js';
import { Web } from "sp-pnp-js";
require('jquery');
import * as $ from 'jquery';

require('./LumenisNewUsefulLinksWebpart.scss')

export interface ILumenisNewUsefulLinksWpWebPartProps {
  description: string;
  wpTitle: string;
  listName: string;
  wpWebUrl: string;
}

export default class LumenisNewUsefulLinksWpWebPart extends BaseClientSideWebPart <ILumenisNewUsefulLinksWpWebPartProps> {

  public render(): void {
    // SPComponentLoader.loadCss("/Style Library/IMF.O365.Lumenis/css/usefullinks.css");
    this.domElement.innerHTML = `
      <p id="LumenisUsefulLinksWpWebPartID" class="anchorinpage"></p>
      <div class="container" id="usefulLinksWP">
          <h3 id="usefulLinksWPTitle">${this.properties.wpTitle}</h3>
          <div class="useful_list">
          </div>
      </div>`;
        this.renderLinks();
}

private renderLinks() {
  let webUrl = this.properties.wpWebUrl;
  let usefulLinksList = this.properties.listName;
  let absUrl = this.context.pageContext.site.absoluteUrl;
  let web = pnp.sp.web;
  if (webUrl != "") {
      web = new Web(absUrl + webUrl);
  }
  else {
      web = new Web(absUrl);
  }

  let resultContainer: Element = this.domElement.querySelector(`.useful_list`);
  resultContainer.innerHTML = "";

  const xml = `<View>
                    <Query>
                        <Where>
                            <Eq>
                                <FieldRef Name="ShowInWP" />
                                <Value Type="Boolean">1</Value>
                            </Eq>
                        </Where>
                        <OrderBy>
                            <FieldRef Name="Index" Ascending="TRUE" />
                        </OrderBy>
                    </Query>
                    <RowLimit>100</RowLimit>
               </View>`;

  const q: any = {
      ViewXml: xml,
  };

  web.get().then(w => {
      web.lists.getByTitle(usefulLinksList).getItemsByCAMLQuery(q).then((r: any[]) => {
          let html = "";
        //   let idx = 0;
              //r.forEach((result) => {
              for(let idx = 0; idx <r.length; idx++){

                  let result = r[idx];
                  if(idx % 3 == 0){
                      html += "";
                      }
                  html += `<div class='item'>
<div><img src='${result.Image.Url}'/></div>
<a  target='_blank'  href='${result.Link}'>
${result.Title}</a>
</div>`;
                  if(idx % 3 == 2){
                      html += "</div>";
                      }
              }
              //});
          resultContainer.insertAdjacentHTML("beforeend", html);
          })
          .catch(console.log);
  });
}

// protected get dataVersion(): Version {
// return Version.parse('1.0');
// }

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
              }),
            PropertyPaneTextField('wpTitle', {
                label: "Title"
            }),
            PropertyPaneTextField('listName', {
                label: "List name"
            }),
            PropertyPaneTextField('wpWebUrl', {
                label: "Web URL"
            })
          ]
        }
      ]
    }
  ]
};
}
}
