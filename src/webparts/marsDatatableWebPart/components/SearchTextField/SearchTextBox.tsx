import * as React from 'react';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import * as strings from 'MarsDatatableWebPartWebPartStrings';
import styles from './SearchTextBox.module.scss';

const SearchTextBox = (props) => {
    return (
        <div className={styles.SearchTextBox}>
            <TextField
                onChanged={props.onSearchChanged}
                placeholder={strings.SearchTextBoxLabel}
                iconProps={{ iconName : "Search"}}                
            />
        </div>
    );
};

export default SearchTextBox;