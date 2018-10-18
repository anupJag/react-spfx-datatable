//#region Import Section
import * as React from 'react';
import styles from './MarsDatatableWebPart.module.scss';
import { IMarsDatatableWebPartProps } from './IMarsDatatableWebPartProps';
import DataTableLoader from './MarsDatatableLoaderWebPart';
import { Placeholder } from "@pnp/spfx-controls-react/lib/Placeholder";
import * as strings from 'MarsDatatableWebPartWebPartStrings';
import { WebPartTitle } from '@pnp/spfx-controls-react/lib/WebPartTitle';
import {
  Environment,
  EnvironmentType
 } from '@microsoft/sp-core-library';
//#endregion

/**
 * Component Handles rendering of Datatable Component and Placeholder component when the webpart is being configured.
 * This is the main entry point for the WebPart.
 */
export default class MarsDatatableWebPart extends React.Component<IMarsDatatableWebPartProps, {}> {

  public render(): React.ReactElement<IMarsDatatableWebPartProps> {
    const renderContent: JSX.Element = (this.props.columnsSelected && this.props.columnsSelected.length > 0) ? <DataTableLoader listId={this.props.listId} webURL={this.props.webURL} columnDetailsRetrieved={this.props.columnDetailsRetrieved} columnsSelected={this.props.columnsSelected} sphttpClient={this.props.sphttpClient}
      reConfigurePane={this.props.fPropertyPaneOpen}
      itemsToBePulled={this.props.itemsToBePulled} 
    /> :
      <Placeholder
        iconName='Edit'
        iconText={strings.noTilesIconText}
        description={Environment.type === EnvironmentType.SharePoint || Environment.type === EnvironmentType.Local ? strings.noTilesConfigured : strings.noTilesConfiguredClassicSP}
        buttonLabel={Environment.type === EnvironmentType.SharePoint || Environment.type === EnvironmentType.Local ? strings.noTilesBtn : null}
        onConfigure={Environment.type === EnvironmentType.SharePoint || Environment.type === EnvironmentType.Local ? this.props.fPropertyPaneOpen : null} />;
    return (
      <div className={styles.marsDatatableWebPart}>
        <div className={styles.mainHolder}>
          <WebPartTitle displayMode={this.props.displayMode}
            title={this.props.title}
            updateProperty={this.props.fUpdateProperty} />
          {renderContent}
        </div>
      </div>
    );
  }
}
