import {
  BaseClientSideWebPart,
  IPropertyPaneSettings,
  IWebPartContext,
  PropertyPaneTextField
} from '@microsoft/sp-client-preview';

import styles from './PropertiesValidators.module.scss';
import * as strings from 'propertiesValidatorsStrings';
import { IPropertiesValidatorsWebPartProps } from './IPropertiesValidatorsWebPartProps';

export default class PropertiesValidatorsWebPart extends BaseClientSideWebPart<IPropertiesValidatorsWebPartProps> {

  public constructor(context: IWebPartContext) {
    super(context);
  }

  public render(): void {
    this.domElement.innerHTML = `
      <div class="${styles.propertiesValidators}">
        <div class="${styles.container}">
          <div class="ms-Grid-row ms-bgColor-themePrimary ms-fontColor-white ${styles.row}">
            <p class="ms-font-l ms-fontColor-white"><strong>JOAO LIVIO SPFx DEMOS</strong></p>
            <p class="ms-font-m ms-fontColor-white"><strong>Description:</strong> This WP will validate a
             PropertyPaneTextField in order to have 10 characters +. Go to WP Properties (Configure Element)</p>
            <p class="ms-font-l ms-fontColor-white"><strong>Text to validate:</strong> ${this.properties.description}</p>
            <div class='ms-font-m ms-fontColor-white'><strong>Loaded from</strong> ${this.context.pageContext.web.title}</div>
          </div>
        </div>
        <div id="spListItemsContainer" /></div>
      </div>`;
  }

  private _validateDescription(value: string): string {
    if (value.length < 10) {
      return "At least 10 characters are required to this property";
    }
    else {
      return "";
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
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel,
                  multiline: true,
                  resizable: true,
                  onGetErrorMessage: this._validateDescription
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
