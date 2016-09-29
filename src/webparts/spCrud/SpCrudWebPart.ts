import {
  BaseClientSideWebPart,
  IPropertyPaneSettings,
  IWebPartContext,
  PropertyPaneDropdown,
  PropertyPaneCheckbox,
  IPropertyPaneDropdownOption
} from '@microsoft/sp-client-preview';

import styles from './SpCrud.module.scss';
import * as strings from 'spCrudStrings';
import { ISpCrudWebPartProps } from './ISpCrudWebPartProps';

import * as pnp from 'sp-pnp-js';
import { Item, ItemAddResult, ItemUpdateResult } from '../../../node_modules/sp-pnp-js/lib/sharepoint/rest/items';

export interface ISPViews { value: ISPView[]; } export interface ISPView { Title: string; Id: string; ListId: string; }

export interface ISPLists {
    value: ISPList[];
}

export interface ISPList {
    Title: string;
    Id: string;
}

interface IListItem {
  Title?: string;
  Id: number;
}

export default class SpCrudWebPart extends BaseClientSideWebPart<ISpCrudWebPartProps> {
  public constructor(context: IWebPartContext) {
    super(context);
    pnp.setup({
      headers: {
        'Accept': 'application/json;odata=nometadata'
      }
    });
  }

  public render(): void {
    this.domElement.innerHTML = `
      <div class="${styles.spCrud}">
        <div class="${styles.container}">
          <div class="ms-Grid-row ms-bgColor-themePrimary ms-fontColor-white ${styles.row}">
            <p class="ms-font-l ms-fontColor-white"><strong>JOAO LIVIO SPFx DEMOS</strong></p>
            <p class="ms-font-m ms-fontColor-white"><strong>Description:</strong> This WP will help you.
              with CRUD operations using the new <a href="https://github.com/OfficeDev/PnP-JS-Core"><strong>PnP JS Core</strong></a>.
               Thanks Guys!, Go to WP Properties (Configure Element)</p>
            <p class="ms-font-l ms-fontColor-white"><strong>Selected List: </strong> ${this.properties.lists}</p>
            <div class='ms-font-m ms-fontColor-white'><strong>Loaded from</strong> ${this.context.pageContext.web.title}</div>
            <p>More samples go to Git from <a href="https://github.com/SharePoint/sp-dev-fx-webparts"><strong>Vesa Juvonen</strong></a></p>
          </div>
          <button class="ms-Button create-Button">
            <span class="ms-Button-label">Create</span>
          </button>
          <button class="ms-Button read-Button">
            <span class="ms-Button-label">Read 1st</span>
          </button>
          <button class="ms-Button readall-Button">
            <span class="ms-Button-label">Read all</span>
          </button>
          <button class="ms-Button update-Button">
            <span class="ms-Button-label">Update 1st</span>
          </button>
          <button class="ms-Button delete-Button">
            <span class="ms-Button-label">Delete 1 by 1st</span>
          </button>
        </div>
        <div id="spListItemsContainer" />
         <div class="${styles.container}; "ms-u-slideRightIn10"">
          <div class="ms-Grid-row ms-fontColor-neutralPrimary ${styles.row}">
            <div class="status"></div>
              <ul class="items"><ul>
            </div>
          </div>
        </div>
      </div>`;

       this._reloadStatus(this._noList() ? '<strong>No list defined</strong>' : 'LIST: ' + this.properties.lists);
       this.setButtonsEventHandlers();
  }

   private setButtonsEventHandlers(): void {
    const webPart: SpCrudWebPart = this;
    this.domElement.querySelector('button.create-Button').addEventListener('click', () => { webPart._createItem(); });
    this.domElement.querySelector('button.read-Button').addEventListener('click', () => { webPart._readItem(); });
    this.domElement.querySelector('button.readall-Button').addEventListener('click', () => { webPart._readItems(); });
    this.domElement.querySelector('button.update-Button').addEventListener('click', () => { webPart._updateItem(); });
    this.domElement.querySelector('button.delete-Button').addEventListener('click', () => { webPart._deleteItem(); });
  }

  private _noList(): boolean {
    return this.properties.lists === undefined ||
      this.properties.lists === null ||
      this.properties.lists.length === 0;
  }

  private _reloadStatus(status: string, items: IListItem[] = []): void {
    this.domElement.querySelector('.status').innerHTML = status;
    this._updateHtml(items);
  }

  private _updateHtml(items: IListItem[]): void {
    const itemsHtml: string[] = [];
    for (let i: number = 0; i < items.length; i++) {
      itemsHtml.push(`<ol>${items[i].Title} (${items[i].Id})</ol>`);
    }

    this.domElement.querySelector('.items').innerHTML = itemsHtml.join('');
  }

  private _createItem(): void {
   this._reloadStatus('Creating item...');
      pnp.sp.web.lists.getByTitle(this.properties.lists).items.add({
        'Title': `Item ${'CREATE AT ' + new Date().getDate()}`
      }).then((result: ItemAddResult): void => {
        const item: IListItem = result.data as IListItem;
        this._reloadStatus(`Item '${item.Title}' (ID: ${item.Id}) successfully created`);
      }, (error: any): void => {
        this._reloadStatus('Error while creating the item: ' + error);
      });
  }

  private _readItem(): void {
    this._reloadStatus('Loading latest items...');
    this._getLatestItemId()
      .then((itemId: number): Promise<IListItem> => {
        if (itemId === -1) {
          throw new Error('No items found in the list');
        }

        this._reloadStatus(`Loading information about item ID: ${itemId}...`);
        return pnp.sp.web.lists.getByTitle(this.properties.lists)
          .items.getById(itemId).select('Title', 'Id').get();
      })
      .then((item: IListItem): void => {
        this._reloadStatus(`Item ID: ${item.Id}, Title: ${item.Title}`);
      }, (error: any): void => {
        this._reloadStatus('Loading latest item failed with error: ' + error);
      });
  }

  private _getLatestItemId(): Promise<number> {
    return new Promise<number>((resolve: (itemId: number) => void, reject: (error: any) => void): void => {
      pnp.sp.web.lists.getByTitle(this.properties.lists)
        .items.orderBy('Id', false).top(1).select('Id').get()
        .then((items: { Id: number }[]): void => {
          if (items.length === 0) {
            resolve(-1);
          }
          else {
            resolve(items[0].Id);
          }
        }, (error: any): void => {
          reject(error);
        });
    });
  }

  private _readItems(): void {
    this._reloadStatus('Loading all items...');
    pnp.sp.web.lists.getByTitle(this.properties.lists)
      .items.select('Title', 'Id').get()
      .then((items: IListItem[]): void => {
        this._reloadStatus(`Successfully loaded ${items.length} items`, items);
      }, (error: any): void => {
        this._reloadStatus('Loading all items failed with error: ' + error);
      });
  }

  private _updateItem(): void {
    this._reloadStatus('Loading latest items...');
    let latestItemId: number = undefined;
    let etag: string = undefined;

    this._getLatestItemId()
      .then((itemId: number): Promise<Item> => {
        if (itemId === -1) {
          throw new Error('No items found in the list');
        }

        latestItemId = itemId;
        this._reloadStatus(`Loading information about item ID: ${itemId}...`);
        return pnp.sp.web.lists.getByTitle(this.properties.lists)
          .items.getById(itemId).get(undefined, {
            headers: {
              'Accept': 'application/json;odata=minimalmetadata'
            }
          });
      })
      .then((item: Item): Promise<IListItem> => {
        etag = item["odata.etag"];
        return Promise.resolve((item as any) as IListItem);
      })
      .then((item: IListItem): Promise<ItemUpdateResult> => {
        return pnp.sp.web.lists.getByTitle(this.properties.lists)
          .items.getById(item.Id).update({
            'Title': `Item ${'UPDATED AT: ' + new Date().getDate()}`
          }, etag);
      })
      .then((result: ItemUpdateResult): void => {
        this._reloadStatus(`Item with ID: ${latestItemId} successfully updated`);
      }, (error: any): void => {
        this._reloadStatus('Loading latest item failed with error: ' + error);
      });
  }

  private _deleteItem(): void {
    if (!window.confirm('Are you sure you want to delete the latest item?')) {
      return;
    }

    this._reloadStatus('Loading latest items...');
    let latestItemId: number = undefined;
    let etag: string = undefined;
    this._getLatestItemId()
      .then((itemId: number): Promise<Item> => {
        if (itemId === -1) {
          throw new Error('No items found in the list');
        }

        latestItemId = itemId;
        this._reloadStatus(`Loading information about item ID: ${latestItemId}...`);
        return pnp.sp.web.lists.getByTitle(this.properties.lists)
          .items.getById(latestItemId).select('Id').get(undefined, {
            headers: {
              'Accept': 'application/json;odata=minimalmetadata'
            }
          });
      })
      .then((item: Item): Promise<IListItem> => {
        etag = item["odata.etag"];
        return Promise.resolve((item as any) as IListItem);
      })
      .then((item: IListItem): Promise<void> => {
        this._reloadStatus(`Deleting item with ID: ${latestItemId}...`);
        return pnp.sp.web.lists.getByTitle(this.properties.lists)
          .items.getById(item.Id).delete(etag);
      })
      .then((): void => {
        this._reloadStatus(`Item with ID: ${latestItemId} successfully deleted`);
      }, (error: any): void => {
        this._reloadStatus(`Error deleting item: ${error}`);
      });
  }


  // Populate Dropdown in propreties
  private _lists: IPropertyPaneDropdownOption[] = [];

  public onInit<T>(): Promise<T> {
    this._getOptions().then((data) => {
        this._lists = data;
    });

    return Promise.resolve();
  }

  private _getLists(url: string) : Promise<ISPLists> {
    return this.context.httpClient.get(url).then((response: Response) => {
        if (response.ok) {
          return response.json();
        } else {
          return null;
        }
      });
  }

  private _getOptions(): Promise<IPropertyPaneDropdownOption[]> {
    var url = this.context.pageContext.web.absoluteUrl + `/_api/web/lists?$filter=Hidden eq false`;

    return this._getLists(url).then((response) => {
        var options: Array<IPropertyPaneDropdownOption> = new Array<IPropertyPaneDropdownOption>();
        var lists: ISPList[] = response.value;
        lists.forEach((list: ISPList) => {
            options.push( { key: list.Title, text: list.Title });
        });
        return options;
    });
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
                PropertyPaneDropdown('lists', {
                  label: strings.ListsFieldLabel,
                  options: this._lists
                }),
                 PropertyPaneCheckbox('allowcreate', {
                   text: "Allow create items",
                     isChecked: false,
                     isEnabled: true
                }),
                  PropertyPaneCheckbox('allowupdate', {
                     text: "Allow update items"
                }),
                  PropertyPaneCheckbox('allowdelete', {
                     text: "Allow delete items"
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
