import * as React from 'react';
import {
    DetailsList, CheckboxVisibility, ConstrainMode
} from 'office-ui-fabric-react/lib/DetailsList';
import styles from './DetailList.module.scss';
import { Spinner, SpinnerSize } from 'office-ui-fabric-react/lib/components/Spinner';
import {
    ContextualMenu
} from 'office-ui-fabric-react/lib/ContextualMenu';


const DetailsListViewer = (props) => {

    const noDataLabel: JSX.Element = !(props._items && props._items.length > 0) ?
        <div className={styles.NoDataLabel}><label>No Data Available</label></div> : <div style={{ display: "none" }} />;

    const showItemLoader: JSX.Element = props.ShowItemLoader ?
        <div>
            <Spinner
                size={SpinnerSize.large}
                label={'Please wait loading data...'}
            />
        </div>
        :
        <div />;
    return (
        <div>
            <div className={styles.root101}>
                <div>
                    <DetailsList
                        columns={props._columns}
                        items={props._items}
                        checkboxVisibility={CheckboxVisibility.hidden}
                        className={styles.DetailListCustomRoot}
                        compact={true}
                        constrainMode={ConstrainMode.unconstrained}
                        onRenderItemColumn={props.RenderItemColumn}
                        onRenderMissingItem={props.RenderMissingItem}
                        usePageCache={true}
                    />
                    {props.ContextualMenuProps && <ContextualMenu {...props.ContextualMenuProps} />}
                </div>
                {showItemLoader}
                {noDataLabel}
            </div>
        </div>
    );
};

export default DetailsListViewer;