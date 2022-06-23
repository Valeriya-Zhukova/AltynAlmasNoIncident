import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './NoIncidentWebPart.module.scss';
import * as strings from 'NoIncidentWebPartStrings';

export interface INoIncidentWebPartProps {
  description: string;
}

import {
  SPHttpClient, SPHttpClientResponse
} from '@microsoft/sp-http';


export interface ISPLists {
  value: ISPList[];
}
export interface ISPList {
  Title: string;
  iDate: Date;
}

export default class NoIncidentWebPart extends BaseClientSideWebPart<INoIncidentWebPartProps> {


  private _getListData(): Promise<ISPLists> {
    return this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + `/_api/web/lists/GetByTitle('NoIncident')/Items`, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => { debugger; return response.json(); });
  }

  private _renderListAsync(): void {
    this._getListData()
      .then((response) => {
        this._renderList(response.value);
      });
  }

  private _renderList(items: ISPList[]): void {

    var akbdate =0;
    var pstdate = 0;
    var one_day = 1000 * 60 * 60 * 24;
    const newLocal = (new Date().getTime()) + (6 * 60 * 60 * 1000);

    items.forEach((item: ISPList) => {

      if (item.Title == "AKB") {
        var d = newLocal - new Date(item.iDate.toString()).getTime();
        akbdate = ((Math.round(Math.abs(d / one_day))) - 1);
      }
      else if (item.Title == "PST") {
        var d = newLocal - new Date(item.iDate.toString()).getTime();
        pstdate = ((Math.round(Math.abs(d / one_day))) - 1);
      }
    });

    let html: string = `<div class="${styles.project} ${styles.back}">
    <div class="${styles.subTitle}">
    <div>Проект</div>
    <div class="${styles.projectTitle}">«Акбакай»</div>
    </div>
    <div>
    <div class="${styles.number}"><span class="${styles.opas}">${akbdate-1}</span></div>
    <div class="${styles.number} ${styles.numwhite}">${akbdate}</div>
    <div class="${styles.number}"><span class="${styles.opas}">${akbdate+1}</span></div>    
    </div>
    <div class="${styles.days}">дн.</div>
</div>
<div class="${styles.project} ${styles.back}">
  <div class="${styles.subTitle}">
    <div>Проект</div>
      <div class="${styles.projectTitle}">«Актогай»</div>
      </div>
      <div>
      <div class="${styles.number}"><span class="${styles.opas}">${pstdate-1}</span> </div>
      <div class="${styles.number} ${styles.numwhite}">${pstdate}</div>
      <div class="${styles.number}"><span class="${styles.opas}">${pstdate+1}</span></div>
  </div>
  <div class="${styles.days}">дн.</div>`;

    const listContainer: Element = this.domElement.querySelector('#spListContainer');
    listContainer.innerHTML = html;
  }

  public render(): void {
    var test = new Date().toLocaleDateString();
    this.domElement.innerHTML = `
    <div class="${styles.noIncident}">
    <div class="${styles.container}">
    <div class="${styles.color}">
    <div class="${styles.square} ${styles.title}">
    <div class="${styles.back} ${styles.title1}"">
    Дни без происшествий
    <br />с потерей рабочих дней
</div>

<div  id="spListContainer"> </div>    

        </div>
       </div>
    </div>
    </div>`;
    this._renderListAsync();
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
