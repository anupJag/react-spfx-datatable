import * as React from 'react';
import styles from './ItemDataInfo.module.scss';

export interface IItemDataInfo {
    InitialCount: number;
    TotalCount: number;
    IsFiltered: boolean;
    CurrentCount: number;
}

const ItemDataInfo = (props: IItemDataInfo) => {
    const ifFiltered: JSX.Element = props.IsFiltered ?
        <p>Showing {props.CurrentCount} of {props.InitialCount} entries. (Total Items : {props.TotalCount})</p>
        :
        <p>Showing {props.CurrentCount == 0 ? props.CurrentCount : (props.CurrentCount === props.TotalCount) ? props.TotalCount : props.CurrentCount as number - 1} of {props.TotalCount} entries</p>;

    return (
        <div
            className={styles.ItemDataInfo}
        >
        {ifFiltered}
        </div>
    );
};

export default ItemDataInfo;
