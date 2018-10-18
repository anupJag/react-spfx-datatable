import * as React from 'react';
import { Placeholder } from "@pnp/spfx-controls-react/lib/Placeholder";
import * as strings from 'MarsDatatableWebPartWebPartStrings';

export interface IErrorHandlerProps {
    ErrorMessage: string;
    fPropertyPaneOpen: () => void;
}

const ErrorHandler = (props: IErrorHandlerProps) => {
    const errorHandling: JSX.Element = (props.ErrorMessage && props.ErrorMessage.indexOf("Access Denied") >= 0) ?
    <Placeholder
        iconName='Error'
        iconText={strings.errorOccured}
        description={props.ErrorMessage}
        onConfigure={props.fPropertyPaneOpen}
    />
    :
    <Placeholder
        iconName='Error'
        iconText={strings.errorOccured}
        description={props.ErrorMessage}
        buttonLabel={"Re-" + strings.noTilesBtn}
        onConfigure={props.fPropertyPaneOpen}
    />;

    return (
        <div>
            {errorHandling}
        </div>
    );
};

export default ErrorHandler;