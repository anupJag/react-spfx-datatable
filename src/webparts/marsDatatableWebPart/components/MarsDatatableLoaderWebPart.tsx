//#region Import Section
import * as React from 'react';
import styles from './MarsDataTableLoaderWebPart.module.scss';
import { IMarsDatatableWebPartLoaderProps } from './IMarsDatatableWebPartProps';
import * as strings from 'MarsDatatableWebPartWebPartStrings';
import { Spinner, SpinnerSize } from 'office-ui-fabric-react/lib/components/Spinner';
import { uniq, escape } from '@microsoft/sp-lodash-subset';
import { FieldType, FieldTypeNames, QueryStructure, IDataCache } from '../Data/IDataService';
import pnp, { Web } from "sp-pnp-js";
import { Link } from 'office-ui-fabric-react/lib/Link';
import { IColumn, ColumnActionsMode } from 'office-ui-fabric-react/lib/DetailsList';
import DetailsListViewer from './DetailList/DetailList';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import ErrorHandler from './ErrorHandler/ErrorHandler';
import SearchText from './SearchTextField/SearchTextBox';
import {
  IContextualMenuProps,
  DirectionalHint
} from 'office-ui-fabric-react/lib/ContextualMenu';
import ItemDataInfo, { IItemDataInfo } from './ItemDataInfo/ItemDataInfo';
//#endregion

export interface IMarsDatatableWebPartLoaderState {
  loading: boolean;
  error: boolean;
  errorMessage: string;
  _listID: string;
  _listData: any[];
  _items: any[];
  _listColumnsDetails: any[];
  _columnDataStructure: {};
  _columnSelected: string[];
  _query: QueryStructure;
  _columns: IColumn[];
  searchQuery: string;
  ITEMS_COUNT: number;
  _isFetchingItems: boolean;
  contextualMenuProps: IContextualMenuProps;
  isSorting: boolean;
  pagedData: string;
  TOTAL_LIST_ITEM_COUNT: number;
  PAGED_DATA_QUERY_TOP: number;
}

const PAGING_SIZE = 30;
const CACHING_DURATION = 10;
//const PAGED_DATA_QUERY_TOP = 50;
const PAGING_DELAY = 3000;

export default class MarsDatatableLoaderWebPart extends React.Component<IMarsDatatableWebPartLoaderProps, IMarsDatatableWebPartLoaderState>{

  constructor(props: IMarsDatatableWebPartLoaderProps) {
    super(props);
    this.state = {
      loading: true,
      error: false,
      errorMessage: undefined,
      _listID: undefined,
      _listData: [],
      _items: [],
      _listColumnsDetails: [],
      _columnDataStructure: {},
      _columnSelected: [],
      _query: undefined,
      _columns: [],
      searchQuery: undefined,
      ITEMS_COUNT: 0,
      _isFetchingItems: false,
      contextualMenuProps: undefined,
      isSorting: false,
      pagedData: undefined,
      TOTAL_LIST_ITEM_COUNT: 0,
      PAGED_DATA_QUERY_TOP: props.itemsToBePulled
    };

    this.cleanLocalStorage(0);
  }

  /**
   * React Lifecycle Hook - invoked when props are received from Parent Controller
   * Updates the Selected Columns to be displayed 
   * @param nextProps 
   */
  componentWillReceiveProps(nextProps: IMarsDatatableWebPartLoaderProps) {
    let _columnsSelectedTemp = [...this.state._columnSelected];
    if (nextProps.columnsSelected != this.props.columnsSelected) {
      _columnsSelectedTemp = nextProps.columnsSelected;
    }
    this.setState({
      _columnSelected: _columnsSelectedTemp
    }, this.buildDetailListColumns);


    if (nextProps.itemsToBePulled != this.props.itemsToBePulled) {
      let newCount: number = nextProps.itemsToBePulled;
      this.setState({
        PAGED_DATA_QUERY_TOP: newCount
      },
        () => {
          this.cleanLocalStorage(1);
          this.recreateData();
        });
    }
  }

  protected recreateData = (): void => {
    this.setState((prevState: IMarsDatatableWebPartLoaderState) => {
      return {
        loading: !prevState.loading
      };
    });
    debugger;
    this.getListItemCollection().then(async (data: any[]) => {
      await this.buildDetailListColumns();

      await this.setState({
        // _listData: _items,
        _items: data,
        ITEMS_COUNT: data ? data.length : 0,
        loading: false
      });
      await this.lazyLoadData();
    });
  }

  /**
   * Method to Create the Data Structure for the FieldCollection.
   * Create the FieldRefs and QueryParamater($expand)
   * @updates state 
   */
  protected fieldCollectionDataStructure = async () => {
    let _columnCollection = {};
    let queryParameters: string[] = [];
    let expand: string[] = [];
    let query: QueryStructure;
    let _ListColumns = [...this.state._listColumnsDetails];

    _ListColumns.forEach((column: any) => {
      var _columnInternalName = column.InternalName;
      switch (column.FieldTypeKind) {

        case FieldType.SingleLineText:
          _columnCollection[_columnInternalName] = FieldTypeNames.SingleLineText;
          queryParameters.push(_columnInternalName);
          break;

        case FieldType.MultiLineText:
          _columnCollection[_columnInternalName] = FieldTypeNames.MultiLineText;
          queryParameters.push(_columnInternalName);
          break;

        case FieldType.Number:
          _columnCollection[_columnInternalName] = FieldTypeNames.Number;
          queryParameters.push(_columnInternalName);
          break;

        case FieldType.Boolean:
          _columnCollection[_columnInternalName] = FieldTypeNames.Boolean;
          queryParameters.push(_columnInternalName);
          break;

        case FieldType.Choice:
          _columnCollection[_columnInternalName] = FieldTypeNames.Choice;
          queryParameters.push(_columnInternalName);
          break;

        case FieldType.Currency:
          _columnCollection[_columnInternalName] = FieldTypeNames.Currency;
          queryParameters.push(_columnInternalName);
          break;

        case FieldType.DateTime:
          _columnCollection[_columnInternalName] = FieldTypeNames.DateTime;
          queryParameters.push(_columnInternalName);
          break;

        //Handle Lookup query here
        case FieldType.LookUp:
          _columnCollection[_columnInternalName] = FieldTypeNames.LookUp;
          var lookUpProperty: string = column.LookupField;
          queryParameters.push(_columnInternalName + "/" + lookUpProperty);
          expand.push(_columnInternalName);
          break;

        case FieldType.MultiChoice:
          _columnCollection[_columnInternalName] = FieldTypeNames.MultiChoice;
          queryParameters.push(_columnInternalName);
          break;

        case FieldType.People:
          _columnCollection[_columnInternalName] = FieldTypeNames.People;
          queryParameters.push(_columnInternalName + "/Title");
          expand.push(_columnInternalName);
          break;

        case FieldType.URL:
          _columnCollection[_columnInternalName] = FieldTypeNames.URL;
          queryParameters.push(_columnInternalName);
          break;

        case FieldType.Integer:
          _columnCollection[_columnInternalName] = FieldTypeNames.Integer;
          queryParameters.push(_columnInternalName);
          break;

        default:
          _columnCollection[_columnInternalName] = FieldTypeNames.SingleLineText;
          queryParameters.push(_columnInternalName);
          break;
      }
    });

    query = {
      queryParameter: queryParameters,
      expandParameter: expand
    };

    console.log(_columnCollection);
    await this.setState({
      _columnDataStructure: _columnCollection,
      _query: query
    });
  }

  /**
   * Method to keep Browser Local Storage Updated/Clean
   */
  protected cleanLocalStorage = (invokedFrom: number) => {
    debugger;
    const _GlobalLocalStorage = window.localStorage;

    if (invokedFrom === 0) {
      Object.keys(_GlobalLocalStorage)
        .filter((el: any) => el.indexOf("-List") >= 0 || el.indexOf("-Fields") >= 0 || el.indexOf("-ItemCount") >= 0)
        .map((key: any) => {
          let _tempLocalStorage: any = localStorage.getItem(key);
          _tempLocalStorage = JSON.parse(_tempLocalStorage);
          if (_tempLocalStorage && _tempLocalStorage.expiration) {
            if (new Date() > new Date(_tempLocalStorage.expiration)) {
              localStorage.removeItem(key);
            }
          }
        });
    }

    //Props changed for Item Count - cleaning local storage
    if (invokedFrom === 1) {
      Object.keys(_GlobalLocalStorage)
        .filter((el: any) => el.indexOf("-List") >= 0)
        .map((key: any) => {
          localStorage.removeItem(key);
        });
    }
  }

  /**
   * 
   */
  protected getTotalItemCount = async () => {
    let w = new Web(this.props.webURL);
    const totalItemCount = await w.lists.getById(this.props.listId).configure({
      headers: {
        'Accept': 'application/json;odata=nometadata',
        'odata-version': ''
      }
    }).usingCaching({
      expiration: pnp.util.dateAdd(new Date, "minute", CACHING_DURATION),
      key: this.props.listId + "-ItemCount",
      storeName: "local",
    }).get().then(p => p.ItemCount);

    await this.setState({
      TOTAL_LIST_ITEM_COUNT: totalItemCount
    });

  }

  /**
   * Method to retrieve List Items from the List selected from the Property Pane Control
   * @returns ListItemCollection
   */
  protected getListItemCollection = async () => {
    let w = new Web(this.props.webURL);
    let listItemCollection: any[] = [];

    let getData;
    try {
      getData = await w.lists.getById(this.props.listId).items.configure({
        headers: {
          'Accept': 'application/json;odata=nometadata',
          'odata-version': ''
        }
      }).usingCaching({
        expiration: pnp.util.dateAdd(new Date, "minute", CACHING_DURATION),
        key: this.props.listId + "-List",
        storeName: "local",
      }).select(...this.state._query.queryParameter).expand(...this.state._query.expandParameter).top(this.state.PAGED_DATA_QUERY_TOP).getPaged()
        .then(p => {
          listItemCollection = listItemCollection.concat(this.createDataCollection(p.results)); return p;
        });

      if (getData.hasNext || getData.nextUrl) {
        await this.setState({
          pagedData: getData.nextUrl
        });
        // listItemCollection = listItemCollection.concat(await this.HandlePagedQuery(getData).then(results => { return results; }));
      }

      //#region Manual Caching Commented
      //Implement Caching For Data
      // localStorage.removeItem(this.props.listId + "-List"); //Remove any exisiting Items in local storage
      // let localStorageCacheKey = this.props.listId + "-List";
      // let expirationDate = new Date();
      // expirationDate.setMinutes(expirationDate.getMinutes() + 20);
      // let cacheObject: IDataCache = {
      //   expirationTime: expirationDate.toISOString(),
      //   results: getData
      // }
      // try {
      //   localStorage.setItem(localStorageCacheKey, JSON.stringify(cacheObject));
      //   console.log("Data Pulled from SharePoint, Added to Local Storage");
      // }
      // catch (error) {
      //   console.log("Attempt to add Data to local storage was unsuccessfull, Normal Data Load will proceed");
      //   console.log(error.message);
      // }
      //#endregion
      return listItemCollection;
    }
    catch (error) {
      await this.setState({
        error: true,
        loading: false,
        errorMessage: strings.ErrorOnItemsFetch
      });
      console.log(error);
    }

  }

  /**
   * Recursive Method Handles Paged Queries
   * Entry Point to the Method is handled by GetListItemCollection Method
   * @param pagedData
   * @returns pagedSiteDirectoryData
   */
  protected handlePagedQuery = async (pagedData: any) => {
    let pagedSiteDirectoryData: any[] = [];
    try {
      pagedSiteDirectoryData = await this.props.sphttpClient.get(`${pagedData}`, SPHttpClient.configurations.v1, {
        headers: {
          'Accept': 'application/json;odata=nometadata',
          'odata-version': ''
        }
      }).then((response: SPHttpClientResponse) => {
        return response.json();
      }).then(async (data: any) => {
        if (data) {
          if (data["odata.nextLink"]) {
            await this.setState({
              pagedData: data["odata.nextLink"]
            });
          }
          else {
            await this.setState({
              pagedData: undefined
            });
          }

          pagedSiteDirectoryData = pagedSiteDirectoryData.concat(this.createDataCollection(data.value));
          return pagedSiteDirectoryData;
        }
        else {
          await this.setState({
            pagedData: undefined
          });
        }
      });

      return pagedSiteDirectoryData;
    }
    catch (error) {
      await this.setState({
        error: true,
        loading: false,
        errorMessage: strings.ErrorOnItemsFetch
      });
      console.log(error);
    }
  }

  /**
   * Method to Create resultant Item Collection from the Item Collection returned by PNP
   * @param collection
   * @returns collectionToBeReturned
   */
  protected createDataCollection = (collection: any[]): any[] => {
    let collectionToBeReturned: any[] = [];
    let _columnDetails: any[] = [...this.state._listColumnsDetails];
    let _columnStructure = this.state._columnDataStructure;
    collection.forEach((element, index: number) => {
      var resultset = {};
      for (var key in element) {
        var structure: string = _columnStructure[key];
        if (structure === FieldTypeNames.LookUp) {
          if (element[key] != null) {
            var tempLookupColumnData = _columnDetails.filter(el => el.InternalName === key);
            var lookupProjectedField: string = tempLookupColumnData[0].LookupField;
            var lookupFieldType: string = tempLookupColumnData[0].TypeAsString;
            if (lookupFieldType.toLowerCase() !== "LookupMulti".toLowerCase()) {
              resultset[key] = element[key][lookupProjectedField];
            }
            else {
              resultset[key] = element[key].map(el => el[lookupProjectedField]).join(';');
            }
          }
          else {
            resultset[key] = "";
          }
        }
        else if (structure === FieldTypeNames.People) {
          if (element[key] != null) {
            var tempColumnData = _columnDetails.filter(el => el.InternalName === key);
            var peopleFieldType: string = tempColumnData[0].TypeAsString;
            if (peopleFieldType.toLowerCase() !== "UserMulti".toLowerCase()) {
              resultset[key] = element[key].Title;
            }
            else {
              resultset[key] = element[key].map(el => el.Title).join(';');
            }
          }
          else {
            resultset[key] = "";
          }
        }
        else if (structure === FieldTypeNames.Boolean) {
          resultset[key] = element[key] ? "YES" : "NO";
        }
        else if (structure === FieldTypeNames.DateTime) {
          if (element[key]) {
            var dateTime = new Date(element[key]);
            resultset[key] = dateTime.toLocaleString();
          }
          else {
            resultset[key] = "";
          }
        }
        else if (structure === FieldTypeNames.MultiChoice) {
          if (element[key]) {
            resultset[key] = element[key].join(',');
          }
          else {
            resultset[key] = "";
          }
        }
        else if (structure === FieldTypeNames.URL) {
          if (element[key]) {
            resultset[key] = { URL: element[key].Url, Description: element[key].Description ? element[key].Description : element[key].Url };
          }
          else {
            resultset[key] = "";
          }
        }
        else if (structure === FieldTypeNames.MultiLineText) {
          if (element[key]) {
            var regex = new RegExp(/(<([^>]+)>)/ig);
            resultset[key] = element[key].toString().replace(regex, "");
          }
          else {
            resultset[key] = "";
          }
        }
        else {
          resultset[key] = element[key];
        }
      }
      resultset["data_custom_index"] = index;
      collectionToBeReturned.push(resultset);
    });
    return collectionToBeReturned;
  }

  /**
   * Get List Columns for the selected list
   * @returns FieldCollection
   */
  protected getListFields = (): Promise<any[]> => {
    let w = new Web(this.props.webURL);
    return new Promise<any[]>((resolve: (columns: any[]) => void, reject: (error: any) => void) => {
      w.lists.getById(this.props.listId).fields.filter("((Hidden eq false) and (ReadOnlyField eq false) and (FieldTypeKind ne 12) and (FieldTypeKind ne 19) and (FieldTypeKind ne 0))").configure({
        headers: {
          'Accept': 'application/json;odata=nometadata',
          'odata-version': ''
        }
      }).select("Title", "InternalName", "FieldTypeKind", "TypeAsString", "LookupField")
        .usingCaching({
          expiration: pnp.util.dateAdd(new Date, "minute", CACHING_DURATION),
          key: this.props.listId + "-Fields",
          storeName: "local"
        }).get().then(p => {
          resolve(p);
        }).catch((error: Error) => {
          reject(error);
        });
    });
  }


  /**
   * Method To Construct the Detail List Columns that are configured to be displayed
   * Bug Fix 1 : Updated Link column condition
   * Bug Fix 2 : Added Multiline condition for FieldType = MultiLineText
   */
  protected buildDetailListColumns = async () => {
    const _columns: IColumn[] = [];
    let _columnDateRetrieved: any[] = [...this.state._listColumnsDetails];
    let _columnsToBeViewed: string[] = [...this.props.columnsSelected];

    _columnsToBeViewed.forEach((col: string, index: number) => {
      var _columnProps = _columnDateRetrieved.filter(el => el.InternalName === col);
      if (_columnProps && _columnProps.length > 0) {
        _columns.push({
          key: (_columnProps[0].FieldTypeKind === FieldType.URL) ? "Link" + index : "Column" + index,
          fieldName: _columnProps[0].InternalName,
          name: _columnProps[0].Title,
          minWidth: 100,
          maxWidth: 180,
          isResizable: true,
          onColumnClick: this.columnContextualMenu.bind(this),
          headerClassName: styles.headerClassName,
          columnActionsMode: ColumnActionsMode.hasDropdown,
          isMultiline: (_columnProps[0].FieldTypeKind === FieldType.MultiLineText) ? true : false,
        });
      }
    });
    console.log(_columns);
    await this.setState({
      _columns: _columns
    });

  }

  /**
   * This Method filters out the data based on the keyword(s) searched using the Search Component
   * Bug Fix 1 : Handled condition for Numeric/Number filtering.
   * Bug Fix 2 : Handled condition for Picture/Hyperlink filtering. 
   * Bug Fix 3 : Handled condition for NULL values
   * Bug Fix 4 : Removed duplicate entries from result set
   * Bug Fix 5 : Handled Search Keyword whitespace
   * @param keyword
   * @returns resultset
   */
  protected keywordSearchHandler = async () => {
    const lowerCaseQuery: string = this.state.searchQuery.toString().toLowerCase().trim();
    let resultSet: any[] = [];
    let _columnCollection: any[] = [...this.state._listColumnsDetails];

    _columnCollection.map(col => {
      var columnTitle = col.InternalName;
      resultSet = resultSet.concat(this.state._items.filter(x => {
        if (x[columnTitle]) {
          if (typeof x[columnTitle] === "object") {
            return x[columnTitle].URL.toString().toLowerCase().indexOf(lowerCaseQuery) >= 0;
          }
          return x[columnTitle].toString().toLowerCase().indexOf(lowerCaseQuery) >= 0;
        }
      }));
    });
    resultSet = uniq(resultSet);
    if ((lowerCaseQuery === "" || lowerCaseQuery === undefined || lowerCaseQuery === null) && (this.state.pagedData != undefined)) {
      this.lazyLoadData();
    }
    else {
      await this.setState({
        _listData: resultSet
      });
    }
  }

  /**
   * React Lifecycle Hook
   * Source for populating the WebPart with Data
   */
  componentDidMount() {
    if (this.props.listId) {
      this.getListFields().then(async (data: any[]) => {
        await this.setState({
          _listColumnsDetails: data
        });
        await this.fieldCollectionDataStructure();
      }).then(async () => {
        await this.getTotalItemCount();
      }).then(() => {
        this.getListItemCollection().then(async (data: any[]) => {
          await this.buildDetailListColumns();

          await this.setState({
            // _listData: _items,
            _items: data,
            ITEMS_COUNT: data ? data.length : 0,
            loading: false
          });
          await this.lazyLoadData();
        });
      }).catch((error: Error) => {
        var regex = new RegExp(/slice|undefined/g);
        this.setState({
          error: true,
          loading: false,
          errorMessage: (error.name === "TypeError" && regex.test(error.message)) ? strings.ErrorOnPermissions : error.message
        });
        console.log(error);
      });
    }
  }

  /**
   * Details List Column Sort Handler
   * For Sort Algorithm refer https://github.com/OfficeDev/office-ui-fabric-react/blob/1a104a1cf2f9a43864ec7ba2ea5c52c99a3a6343/packages/office-ui-fabric-react/src/components/DetailsList/examples/DetailsList.CustomColumns.Example.tsx#L53 
   * @param event
   * @param column
   */
  protected columnSortHandler = (column: IColumn, isSortedDescending: boolean) => {
    const { _columns } = this.state;
    let key = column.fieldName;
    const _items = this.state.searchQuery ? this.state._listData : this.state._items;
    //let { _listData } = this.state;
    const sortedItems = _items
      .slice(0)
      .sort((a: any, b: any) => ((isSortedDescending ? a[key] < b[key] : a[key] > b[key]) ? 1 : -1));

    this.setState({
      _listData: sortedItems,
      isSorting: true,
      _columns: _columns!.map(col => {
        col.isSorted = col.key === column.key;

        if (col.isSorted) {
          col.isSortedDescending = isSortedDescending;
        }

        return col;
      }),

    });

  }

  /**
   * Method Updates the State with the Keyword(s) typed in the Search Text Box
   * @param value
   */
  protected searchQueryChangeHandler = (value: any): void => {
    this.setState({
      searchQuery: escape(value)
    },
      this.keywordSearchHandler
    );
  }

  /**
   * Method handles how data is shown in Detail List
   * Method is passed to Details List as a prop/reference
   * @param item
   * @param index
   * @param column   
   */
  protected renderItemColumnHandler = (item: any, index: number, column: IColumn): JSX.Element => {
    const fieldContent = item[column.fieldName || ''];
    if (column.key.toString().indexOf('Link') >= 0) {
      return <Link href={fieldContent.URL} style={{ color: "#006cc8" }}>{fieldContent.Description}</Link>;
    }
    else {
      return <span>{fieldContent}</span>;
    }
  }

  /**
   * Method used to rebuilt data to be displayed for Lazy Loaded data
   * @param index 
   */
  private onDataMiss(index: number): void {
    if (this.state.searchQuery || this.state.isSorting) {
      return;
    }

    const { ITEMS_COUNT } = this.state;
    index = Math.floor(index / PAGING_SIZE) * PAGING_SIZE;

    if (!this.state._isFetchingItems) {

      this.setState((prevState, prevProps) => {
        return {
          _isFetchingItems: !prevState._isFetchingItems
        };
      });

      let itemsCopy = ([] as any[]).concat(this.state._listData);
      itemsCopy.pop();
      itemsCopy.splice.apply(itemsCopy, [index, PAGING_SIZE].concat(this.state._items.slice(index, index + PAGING_SIZE)));

      if (itemsCopy.length < ITEMS_COUNT) {
        itemsCopy = itemsCopy.concat(new Array(1));
      }
      else {
        if (this.state.pagedData) {
          this.constructPagedData().then(() => {
            itemsCopy = itemsCopy.concat(new Array(1));
            console.log("Paginated Data Added, Item Count ", this.state.ITEMS_COUNT);
          });
        }
      }

      setTimeout(() => {
        this.setState((prevState, prevProps) => {
          return {
            _listData: itemsCopy,
            _isFetchingItems: !prevState._isFetchingItems
          };
        });
      }, PAGING_DELAY);
    }
  }

  /**
   * Async Method to Update Dataset when Next Set Of Paged Data is required 
   */
  protected constructPagedData = async () => {
    let _itemsToBeLoadedTo = [...this.state._items];
    _itemsToBeLoadedTo = _itemsToBeLoadedTo.concat(await this.handlePagedQuery(this.state.pagedData).then(results => { return results; }));
    await this.setState({
      ITEMS_COUNT: _itemsToBeLoadedTo.length,
      _items: _itemsToBeLoadedTo
    });
    return true;
  }

  /**
   * Method checks whether the List requires Lazy Loading
   */
  private lazyLoadData = async () => {
    const { ITEMS_COUNT } = this.state;
    const isLazyLoaded: boolean = ITEMS_COUNT > PAGING_SIZE ? true : false;
    await this.setState({
      _listData: isLazyLoaded ? this.state._items.slice(0, PAGING_SIZE).concat(new Array(1)) : this.state._items
    });
  }

  /**
   * Method called when Detail List renders an Empty array. This is used to populate the next batch of data
   */
  private onRenderMissingItem = (index: number): null => {
    this.onDataMiss(index as number);
    return null;
  }

  private getContextualMenuProps = (ev: React.MouseEvent<HTMLElement>, column: IColumn): IContextualMenuProps => {
    const items = [
      {
        key: 'aToZ',
        name: 'A to Z',
        iconProps: { iconName: 'SortUp' },
        canCheck: true,
        checked: column.isSorted && !column.isSortedDescending,
        onClick: () => this.columnSortHandler(column, false)
      },
      {
        key: 'zToA',
        name: 'Z to A',
        iconProps: { iconName: 'SortDown' },
        canCheck: true,
        checked: column.isSorted && column.isSortedDescending,
        onClick: () => this.columnSortHandler(column, true)
      }
    ];

    return {
      items: items,
      target: ev.currentTarget as HTMLElement,
      directionalHint: DirectionalHint.bottomLeftEdge,
      gapSpace: 10,
      isBeakVisible: true,
      directionalHintForRTL: DirectionalHint.bottomLeftEdge,
      onDismiss: this._onContextualMenuDismissed
    };
  }

  private columnContextualMenu = (ev: React.MouseEvent<HTMLElement>, column: IColumn): void => {
    if (column.columnActionsMode !== ColumnActionsMode.disabled) {
      this.setState({
        contextualMenuProps: this.getContextualMenuProps(ev, column)
      });
    }
  }

  private _onContextualMenuDismissed = (): void => {
    this.setState({
      contextualMenuProps: undefined
    });
  }

  /**
   * React Lifecycle Method to create Virtual DOM and Render the virtual DOM to the REAL DOM
   */
  public render(): React.ReactElement<IMarsDatatableWebPartLoaderProps> {

    const _errorHandler: JSX.Element = this.state.error ?
      <div>
        <ErrorHandler
          ErrorMessage={this.state.errorMessage}
          fPropertyPaneOpen={this.props.reConfigurePane}
        />
      </div>
      :
      <div />;

    const _IsLoading: JSX.Element = this.state.loading ?
      <div>
        <Spinner
          size={SpinnerSize.large}
          label={'Please wait loading data...'}
        />
      </div>
      :
      <div className={styles.container}>
        <div>
          <div className={styles.menu}>
            <ItemDataInfo
              CurrentCount={this.state._listData && this.state._listData.length > 0 ? this.state._listData.length : 0}
              InitialCount={this.state._items && this.state._items.length > 0 ? this.state._items.length : 0}
              TotalCount={this.state.TOTAL_LIST_ITEM_COUNT}
              IsFiltered={this.state.searchQuery && this.state.searchQuery.trim().length > 0 ? true : false}
            />
            <SearchText onSearchChanged={this.searchQueryChangeHandler.bind(this)} />
          </div>
          <div className={styles.DetailListContainer} role="region" data-is-scrollable="true">
            <div className={styles.root}>
              <div className={styles.layoutPositioningContainer}>
                <div className={styles.layoutSwitcherContainer}>
                  <DetailsListViewer
                    _columns={this.state._columns}
                    _items={this.state._listData}
                    ColumnSortClicked={this.columnSortHandler.bind(this)}
                    RenderItemColumn={this.renderItemColumnHandler.bind(this)}
                    RenderMissingItem={this.onRenderMissingItem.bind(this)}
                    ShowItemLoader={this.state._isFetchingItems}
                    ContextualMenuProps={this.state.contextualMenuProps}
                  />
                </div>
              </div>
            </div>
          </div>
        </div>
      </div >;

    return (
      <div className={styles.marsDataTableLoader}>
        {(this.state.error) ? _errorHandler : _IsLoading}
      </div>
    );
  }
}