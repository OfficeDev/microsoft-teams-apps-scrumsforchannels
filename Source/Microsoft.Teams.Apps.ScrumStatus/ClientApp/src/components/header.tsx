/*
    <copyright file="manage-experts.tsx" company="Microsoft">
    Copyright (c) Microsoft. All rights reserved.
    </copyright>
*/

import * as React from "react";
import { Table } from "@fluentui/react-northstar";
import "../styles/style.css";

interface IHeaderProps {
    resourceStrings: any,
}

/** Component for displaying settings header layout on UI. */
const Header: React.FunctionComponent<IHeaderProps> = props => {
    return (
            <>
                <Table.Row header className="table-row-header">
                    <div className="table-header1"><Table.Cell content={props.resourceStrings.teamNameTitle} key="TeamName" /></div>
                    <div className="table-header2"><Table.Cell content={props.resourceStrings.scrumEnableButtonTitle} key="Status" /></div>
                    <div className="table-header3"><Table.Cell content={props.resourceStrings.membersListTitle} key="memberslist" /></div>
                    <div className="table-header4"><Table.Cell content={props.resourceStrings.startEveryDayTitle} key="StartTime" /></div>
                    <div className="table-header5"><Table.Cell content={props.resourceStrings.timeZoneTitle} key="Timezone" /></div>
                    <div className="table-header6"><Table.Cell content={props.resourceStrings.addToChannelTitle} key="channelname" /></div>
                    <div><Table.Cell content="" /></div>
                </Table.Row>
            </>
    );
}

export default Header;