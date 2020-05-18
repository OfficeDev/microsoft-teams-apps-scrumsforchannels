/*
    <copyright file="manage-experts.tsx" company="Microsoft">
    Copyright (c) Microsoft. All rights reserved.
    </copyright>
*/

import * as React from "react";
import { Text, Flex, Button } from "@fluentui/react-northstar";
import { AddIcon } from '@fluentui/react-icons-northstar';
import "../styles/style.css";

interface IFooterProps {
    resourceStrings: any,
    errorMessage: string,
    isSaveScrumSettingsLoading: boolean,
    addNewScrum: (event: any) => void,
    saveScrumSettings: (event: any) => void
}

/** Component for displaying settings header layout on UI. */
const Footer: React.FunctionComponent<IFooterProps> = props => {
    return (
        <div className="footer">
            <div>
                <Flex gap="gap.smaller">
                    <Button icon={<AddIcon />} text primary content={props.resourceStrings.addNewScrumButtonText} onClick={props.addNewScrum} />
                </Flex>
            </div>
            <div className="error">
                <Flex gap="gap.small">
                    {props.errorMessage !== null && <Text className="small-margin-left" content={props.errorMessage} error />}
                </Flex>
            </div>
            <div className="button">
                <Flex gap="gap.smaller">
                    <Button fluid content={props.resourceStrings.saveButtonText} primary loading={props.isSaveScrumSettingsLoading} disabled={props.isSaveScrumSettingsLoading} onClick={props.saveScrumSettings} />
                </Flex>
            </div>
        </div>
    );
}

export default Footer;