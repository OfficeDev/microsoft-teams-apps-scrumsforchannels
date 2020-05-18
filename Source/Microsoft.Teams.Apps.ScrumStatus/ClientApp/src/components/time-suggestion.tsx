import * as React from "react";
import { Dropdown } from "@fluentui/react-northstar";
import moment from 'moment';
import _ from 'lodash';
import { getUserLocale } from '../localization/translate';
import Constants from "../constants";
import "../styles/style.css";

moment.locale(getUserLocale());

interface ITimeSuggestionProps {
    placeholder: string,
    onTimeChange: (event: any) => void,
    selectedValue: string,
    id: string
}

/** Component for displaying time suggestion drop down on UI. */
const TimeSuggestion: React.FunctionComponent<ITimeSuggestionProps> = props => {
    const timeSuggestions = _.range(0, 1440, 30).map(minutes => {    
        const timeTag = moment()
            .startOf('day')
            .minutes(minutes)
            .format(Constants.timePickerFormat);

        return timeTag;
    });

    return (
        <>
            <Dropdown
                fluid
                className="table-cell4"
                items={timeSuggestions}
                placeholder={props.placeholder}
                onChange={props.onTimeChange}
                value={props.selectedValue}
                id={props.id}
                title={props.selectedValue}
            />
        </>
    );
}

export default TimeSuggestion;
