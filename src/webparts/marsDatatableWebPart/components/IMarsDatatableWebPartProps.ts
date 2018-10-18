import { SPHttpClient } from '@microsoft/sp-http';
import { DisplayMode } from '@microsoft/sp-core-library';


export interface IMarsDatatableWebPartProps{
  listId: string; 
  columnsSelected: string[];
  webURL: string;
  columnDetailsRetrieved: any[];
  sphttpClient: SPHttpClient;
  fPropertyPaneOpen: () => void;
  displayMode: DisplayMode; 
  fUpdateProperty: (value: string) => void;
  title : string;
  itemsToBePulled : number;
}

export interface IMarsDatatableWebPartLoaderProps {
  listId: string;
  columnsSelected: string[];
  webURL: string;
  sphttpClient: SPHttpClient;
  columnDetailsRetrieved: any[];
  reConfigurePane : () => void;
  itemsToBePulled : number;
}
