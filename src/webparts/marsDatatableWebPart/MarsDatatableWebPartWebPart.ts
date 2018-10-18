//#region Import Section
import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneSlider,
} from '@microsoft/sp-webpart-base';
import { PropertyFieldListPicker, PropertyFieldListPickerOrderBy } from '@pnp/spfx-property-controls/lib/PropertyFieldListPicker';
import { PropertyFieldMultiSelect } from '@pnp/spfx-property-controls/lib/PropertyFieldMultiSelect';
import * as strings from 'MarsDatatableWebPartWebPartStrings';
import MarsDatatableWebPart from './components/MarsDatatableWebPart';
import { IListColumnOptions } from './Data/IDataService';
import { IMarsDatatableWebPartProps } from './components/IMarsDatatableWebPartProps';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
//#region IE fixes for ES6 JS Polyfills
import 'core-js/es6/number';
import 'core-js/es6/array';
//#endregion
//#endregion

/**
 * WebPart Props Interface 
 */
export interface IMarsDatatableWebPartWebPartProps {
  description: string;
  list: string;
  columns: string[];
  title: string;
  itemsToBePulled : number;
}

export default class MarsDatatableWebPartWebPart extends BaseClientSideWebPart<IMarsDatatableWebPartWebPartProps> {

  private _columnsMultiSelectDisabled: boolean = true;
  private _listColumns: IListColumnOptions[];
  private _listColumnDetails: any[];

  public render(): void {
    const element: React.ReactElement<IMarsDatatableWebPartProps> = React.createElement(
      MarsDatatableWebPart,
      {
        listId: this.properties.list,
        columnsSelected: this.properties.columns,
        webURL: this.context.pageContext.web.absoluteUrl,
        columnDetailsRetrieved: this._listColumnDetails,
        fPropertyPaneOpen: this.context.propertyPane.open,
        sphttpClient: this.context.spHttpClient,
        displayMode: this.displayMode,
        fUpdateProperty: (value: string) => {
          this.properties.title = value;
        },
        title: this.properties.title,
        itemsToBePulled : this.properties.itemsToBePulled
      }
    );

    ReactDom.render(element, this.domElement);
  }

  /**
   * Disable Reactive Property of WebPart - Adds Apply button to the WebPart Property Pane
  */
  protected get disableReactivePropertyChanges(): boolean {
    return true;
  }

  
  /**
   * SPFx WebPart Internal Method - Triggered when any Property Changed
   * @param propertyPath 
   * @param oldValue 
   * @param newValue 
   */
  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void {
    if (propertyPath == "list" && newValue) {
      
      super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
      // get previously selected item
      const previousColumn: string[] = this.properties.columns;
      this.properties.columns = [];
      this.onPropertyPaneFieldChanged('columns', previousColumn, this.properties.columns);
      this._columnsMultiSelectDisabled = true;
      // refresh the item selector control by repainting the property pane
      this.context.propertyPane.refresh();
      // communicate loading items
      this.context.statusRenderer.displayLoadingIndicator(this.domElement, 'columns');

      this.getColumnsForPropertyPane().then((columns: any[]): void => {
        var columnsRequired: IListColumnOptions[] = [];
        this._listColumnDetails = columns;
        columns.forEach((element: any) => {
          columnsRequired.push({
            key: element.InternalName,
            text: element.Title,

          });
        });
        this._listColumns = columnsRequired;
        this._columnsMultiSelectDisabled = false;
        this.context.statusRenderer.clearLoadingIndicator(this.domElement);
        this.render();
        this.context.propertyPane.refresh();
      });
    }
    else {
      super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);

    }
  }

  /**
   * SPFx WebPart Internal Method
   */
  protected onPropertyPaneConfigurationStart(): void {

    this._columnsMultiSelectDisabled = !this.properties.list || !this._listColumns;
    if (!this.properties.list) {
      return;
    }

    this.context.statusRenderer.displayLoadingIndicator(this.domElement, 'columns');

    this.getColumnsForPropertyPane()
      .then((columns: any[]): void => {
        var columnsRequired: IListColumnOptions[] = [];
        this._listColumnDetails = columns;
        columns.forEach((element: any) => {
          columnsRequired.push({
            key: element.InternalName,
            text: element.Title
          });
        });
        this._listColumns = columnsRequired;
        this._columnsMultiSelectDisabled = !this.properties.list;
        this.context.propertyPane.refresh();
        this.context.statusRenderer.clearLoadingIndicator(this.domElement);
        this.render();
      });
  }

  /**
   * Method to Retrieve Columns for the List Selected from the Property Pane Control
   */
  protected getColumnsForPropertyPane = (): Promise<any[]> => {

    if (!this.properties.list) {
      // resolve to empty options since no list has been selected
      return Promise.resolve();
    }

    return new Promise<any[]>((resolve: (columns: any[]) => void, reject: (error: any) => void) => {
      this.context.spHttpClient.get(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbyid('${this.properties.list}')/fields/?$filter=((Hidden eq false) and (ReadOnlyField eq false) and (FieldTypeKind ne 12) and (FieldTypeKind ne 19) and (FieldTypeKind ne 0))`, SPHttpClient.configurations.v1, {
        headers: {
          'Accept': 'application/json;odata=nometadata',
          'odata-version': ''
        }
      }).then((response: SPHttpClientResponse) => {
        return response.json();
      }).then((data: any) => {
        let _tempData: any[] = [];
        (data.value).forEach((element: any) => {
          _tempData.push(element);
        });
        resolve(_tempData);
      }).catch((error: Error) => {
        reject(error);
      });
    });
  }

  /**
   * SPFx WebPart Internal Method - Call To Dispose WebPart
   */
  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  /**
   * SPFx WebPart Internal Method
   */
  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  /**
   * SPFx WebPart Internal Methods - Define and Construct WebPart Property Pane
   */
  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription,
          },
          groups: [
            {
              groupFields: [
                PropertyFieldListPicker('list', {
                  label: strings.ListSelectionLabel,
                  selectedList: this.properties.list,
                  includeHidden: false,
                  context: this.context,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  deferredValidationTime: 0,
                  baseTemplate: 100,
                  orderBy: PropertyFieldListPickerOrderBy.Title,
                  key: 'listPickerFieldId'
                }),
                PropertyFieldMultiSelect('columns', {
                  label: strings.ColumnSelectionLabel,
                  key: 'columns',
                  selectedKeys: this.properties.columns,
                  options: this._listColumns,
                  disabled: this._columnsMultiSelectDisabled
                }),
                PropertyPaneSlider('itemsToBePulled', {
                  max : 5000,
                  min: 50,
                  showValue : true,
                  label: "Max # of list items"
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
