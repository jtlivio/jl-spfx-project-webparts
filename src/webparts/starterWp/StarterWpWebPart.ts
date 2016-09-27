import {
  BaseClientSideWebPart,
  IPropertyPaneSettings,
  IWebPartContext,
  PropertyPaneSlider
} from '@microsoft/sp-client-preview';

import styles from './StarterWp.module.scss';
import * as strings from 'starterWpStrings';
import { IStarterWpWebPartPropspropertiesStarter } from './IStarterWpWebPartProps';

import { EnvironmentType } from '@microsoft/sp-client-base';
import MockItemsStarter from './Mocks/MockItemsStarter';

export interface ISPListItem {
  Id: string;
  Title: string;
  MyMultiText: string;
}

export interface ISPListItems {
  value: ISPListItem[];
}

export default class StarterWpWebPart extends BaseClientSideWebPart<IStarterWpWebPartPropspropertiesStarter> {

  public constructor(context: IWebPartContext) {
    super(context);
  }

  public render(): void {
    this.domElement.innerHTML = `
      <div class="${styles.starterWp}">
        <div class="${styles.container}">
          <div class="ms-Grid-row ms-bgColor-themePrimary ms-fontColor-white ${styles.row}">
            <p class="ms-font-l ms-fontColor-white"><strong>JOAO LIVIO SPFx DEMOS</strong></p>
            <p class="ms-font-m ms-fontColor-white"><strong>Description:</strong> This WP will fetch the number of items
             that you choose in the propreties. It's non reactive, you have to apply for assuming you value</p>
            <p class="ms-font-l ms-fontColor-white"><strong>Items to fetch:</strong> ${this.properties.maxitems}</p>
            <div class='ms-font-m ms-fontColor-white'><strong>Loaded from</strong> ${this.context.pageContext.web.title}</div>
          </div>
        </div>
        <div id="spListItemsContainer" /></div>
      </div>`;

      this._renderListAsync();
  }


  private _getMockListData(): Promise<ISPListItems> {
    return MockItemsStarter.get(this.context.pageContext.web.absoluteUrl)
        .then((data: ISPListItem[]) => {
             var listData: ISPListItems = { value: data };
             return listData;
         }) as Promise<ISPListItems>;
  }

  private _getListData(): Promise<ISPListItems> {
  var lib="myCustomList";

    return this.context.httpClient.get(this.context.pageContext.web.absoluteUrl + `/_api/web/lists/getbytitle('${lib}')/items?$top=` + this.properties.maxitems)
    .then((response: Response) => {
          return response.json();
     });
  }

   private _renderList(items: ISPListItem[]): void {
      let html: string = "";
      items.forEach((item: ISPListItem) => {
          html += `
          <ul class="${styles.starterWp}2>
              <p class="${styles.listItem}">
                  <span class="ms-fontSize-l">ID: ${item.Id} - Title: ${item.Title}</span>
                  <br/><br/>
                  <span class="ms-fontSize-l">Multi Text</span>
                  <br/><br/>
                  <span class="ms-font-l">${item.MyMultiText}</span>
              </p>
          </ul>
          <br/>`;
      });

      const listContainer: Element = this.domElement.querySelector('#spListItemsContainer');
      listContainer.innerHTML = html;

  }

  private _renderListAsync(): void {
      // Local environment
      if (this.context.environment.type === EnvironmentType.Local) {
          this._getMockListData().then((response) => {
          this._renderList(response.value);
          }); }
      else {
          this._getListData()
          .then((response) => {
              this._renderList(response.value);
          });
      }
  }



  protected get propertyPaneSettings(): IPropertyPaneSettings {
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
                 PropertyPaneSlider('maxitems', {
                    label: strings.MaxItemsFieldLabel,
                    max: 5000,
                    min: 100,
                    step: 100,
                    showValue: true
                })
              ]
            }
          ]
        }
      ]
    };
  }
   protected get disableReactivePropertyChanges(): boolean {
		return true;
	}
}
