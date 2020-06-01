/*
    <copyright file="manage-experts.tsx" company="Microsoft">
    Copyright (c) Microsoft. All rights reserved.
    </copyright>
*/

import * as React from "react";
import { Dropdown, Button, Loader, Flex, Checkbox, Table, Input } from "@fluentui/react-northstar";
import { TrashCanIcon } from '@fluentui/react-icons-northstar';
import { createBrowserHistory } from "history";
import { ApplicationInsights, SeverityLevel } from "@microsoft/applicationinsights-web";
import { ReactPlugin, withAITracking } from "@microsoft/applicationinsights-react-js";
import * as microsoftTeams from "@microsoft/teams-js";
import moment from 'moment';
import { getResourceStrings, getScrumConfigurationDetailsbyAADGroupID, getTeamDetails, deleteScrumConfigurationDetails, saveScrumConfigurationDetails, handleError, getTimeZoneInfo } from "../api/scrum-status-api";
import Header from "./header";
import Footer from './footer'
import TimeSuggestion from "./time-suggestion";
import { IScrumProps, IUserDetails } from "../models/type";
import "../styles/style.css";
import { getUserLocale } from '../localization/translate';
import Constants from "../constants";

moment.locale(getUserLocale());

const browserHistory = createBrowserHistory({ basename: "" });
let reactPlugin = new ReactPlugin();
let teamId = "";
let groupID = "";
let channels: any[] = [];
let timeZones: any[] = [];

interface IScrumState {
    scrums: IScrumProps[],
    showError: boolean,
    errorMessage: string,
    loading: boolean,
    resourceStrings: any;
    resourceStringsLoaded: boolean;
    teamMembers: IUserDetails[];
    isSaveScrumSettingsLoading: boolean;
    deletedScrums: IScrumProps[]
}

/** Component for displaying scrum configuration settings. */
class Settings extends React.Component<IScrumProps, IScrumState>
{
    customAPIAuthenticationToken?: string | null = null;
    locale?: string | null;
    serviceUrl: string | null = null;
    telemetry?: any = null;
    appInsights: ApplicationInsights;
    userEmail?: string | null = null;
    userObjectId?: string | null = null;

    constructor(props: any) {
        super(props);

        this.state = {
            scrums: [],
            showError: false,
            errorMessage: "",
            loading: false,
            resourceStrings: {},
            resourceStringsLoaded: false,
            teamMembers: [],
            isSaveScrumSettingsLoading: false,
            deletedScrums: [],
        };

        let search = window.location.search;
        let params = new URLSearchParams(search);
        this.telemetry = params.get("telemetry");
        this.customAPIAuthenticationToken = params.get("token");
        this.locale = params.get("locale");
        this.serviceUrl = params.get("serviceurl");

        // Initialize application insights for logging events and errors.
        try {
            this.appInsights = new ApplicationInsights({
                config: {
                    instrumentationKey: this.telemetry,
                    extensions: [reactPlugin],
                    extensionConfig: {
                        [reactPlugin.identifier]: { history: browserHistory }
                    }
                }
            });
            this.appInsights.loadAppInsights();
        }
        catch (exception) {
            this.appInsights = new ApplicationInsights({
                config: {
                    instrumentationKey: undefined,
                    extensions: [reactPlugin],
                    extensionConfig: {
                        [reactPlugin.identifier]: { history: browserHistory }
                    }
                }
            });
            console.log(exception);
        }
    }

    /** Called once component is mounted. */
    async componentDidMount() {
        microsoftTeams.initialize();
        microsoftTeams.getContext((context) => {
            this.userObjectId = context.userObjectId;
            this.userEmail = context.upn;
            this.locale = context.locale;

            teamId = context.teamId ? context.teamId : "";
            groupID = context.groupId ? context.groupId : "";
            
            this.getResourceStrings();
            this.getTimeZoneDetails();
            this.getTeamDetails(teamId)
                .then(() => {
                    this.getScrumConfigurationDetails(groupID)
                })
            
        });
    }

    /** 
    *  Get resource strings according to user locale.
    * */
    getResourceStrings = async () => {
        this.appInsights.trackTrace({ message: `'getResourceStrings' - Request initiated`, severityLevel: SeverityLevel.Information, properties: { User: this.userObjectId } });
        const resourceStringsResponse = await getResourceStrings(this.customAPIAuthenticationToken!, this.locale);
        if (resourceStringsResponse) {
            this.setState({ resourceStringsLoaded: true });

            if (resourceStringsResponse.status === 200) {
                this.setState({ resourceStrings: resourceStringsResponse.data });
            }
            else {
                handleError(resourceStringsResponse, this.customAPIAuthenticationToken);
            }
        }
    }

    /** 
    *  Get all team members.
    * */
    getTeamDetails = async (teamId: string) => {
        this.appInsights.trackTrace({ message: `'getTeamDetails' - Request initiated`, severityLevel: SeverityLevel.Information });
        this.setState({ loading: true });
        const teamDetailsResponse = await getTeamDetails(this.customAPIAuthenticationToken!, teamId);
        if (teamDetailsResponse) {
            if (teamDetailsResponse.status === 200) {
                let teamDetails: any = teamDetailsResponse;
                this.setState({ teamMembers: teamDetails.data.teamMembers });
                channels = teamDetails.data.channels;
            }
            else {
                handleError(teamDetailsResponse, this.customAPIAuthenticationToken);
            }
        }
        this.setState({ loading: false });
    }

    /**
    *  Get scrum configuration details from storage.
    * */
    getScrumConfigurationDetails = async (groupID: string) => {
        this.appInsights.trackTrace({ message: `'getScrumConfigurationDetails' - Request initiated`, severityLevel: SeverityLevel.Information });
        this.setState({ loading: true });
        const scrumConfigurationDetailsResponse = await getScrumConfigurationDetailsbyAADGroupID(this.customAPIAuthenticationToken!, groupID);
        if (scrumConfigurationDetailsResponse) {
            if (scrumConfigurationDetailsResponse.status === 200) {
                this.setState({
                    scrums: scrumConfigurationDetailsResponse.data as IScrumProps[]
                });

                // Get names of scrum members from team member details.
                this.state.scrums.forEach((scrumConfigurationDetail) => {
                    if (scrumConfigurationDetail.UserPrincipalNames) {
                        scrumConfigurationDetail.SelectedMembers = [];
                        scrumConfigurationDetail.UserPrincipalNames.split(",").forEach((user) => {
                            let member: any = this.state.teamMembers.find(member => member.content === user);
                            if (member) {
                                scrumConfigurationDetail.SelectedMembers.push(member);
                            }
                        });
                    }
                    if (scrumConfigurationDetail.TimeZone) {
                        let selectedTimeZone: any = timeZones.find(timeZone => timeZone.timeZoneId === scrumConfigurationDetail.TimeZone);
                        scrumConfigurationDetail.SelectedTimeZone = selectedTimeZone.header;
                        scrumConfigurationDetail.StartTime = moment(scrumConfigurationDetail.StartTime)
                            .format(Constants.timePickerFormat);
                    }
                });
            }
            else {
                handleError(scrumConfigurationDetailsResponse, this.customAPIAuthenticationToken);
            }
        }
        this.setState({ loading: false });
    };

    /** 
    *  Get team zone information.
    * */
    getTimeZoneDetails = async () => {
        this.appInsights.trackTrace({ message: `'getTimeZoneDetails' - Request initiated`, severityLevel: SeverityLevel.Information });
        this.setState({ loading: true });
        const timeZoneResponse = await getTimeZoneInfo(this.customAPIAuthenticationToken!);
        if (timeZoneResponse) {
            if (timeZoneResponse.status === 200) {
                timeZones = timeZoneResponse.data;
            }
            else {
                handleError(timeZoneResponse, this.customAPIAuthenticationToken);
            }
        }
        this.setState({ loading: false });
    }

    /**
    *  Gets called when user clicks on add new scrum button.
    * */
    private addNewScrum = () => {
        this.setState({
            scrums: this.state.scrums.concat([{ ScrumTeamConfigId: "", TeamId: "", ChannelId: "", ScrumTeamName: "", IsActive: true, StartTime: "", ChannelName: "", TimeZone: "", SelectedTimeZone: "", AADGroupID: "", CreatedOn: "", CreatedBy: "", UserPrincipalNames: "", SelectedMembers: [], ScrumConfigurationId: "", ServiceUrl: this.serviceUrl }])
        });
        return false;
    };

    /**
    *  Gets called when user enters team name.
    * */
    private teamNameChange = (index: number, event: any) => {
        let Scrumprop = this.state.scrums;
        Scrumprop[index].ScrumTeamName = event.target.value;
        this.setState({
            scrums: Scrumprop,
        });
    };

    /**
    *  Gets called when user changes scrum status.
    * */
    private setScrumStatus = (e: any, checkboxProps: any) => {
        let Scrumprop = this.state.scrums;
        if (checkboxProps) {
            Scrumprop[parseInt(checkboxProps!.id)].IsActive = checkboxProps.checked;
        }
        this.setState({
            scrums: Scrumprop
        });
    };

    /**
    *  Gets called when user selects team member from member list.
    * */
    private memberSelectionChange = (event: any, dropdownProps?: any) => {
        if (dropdownProps) {
            let scrumProp = this.state.scrums;
            let selectedUsers = dropdownProps.value;
            selectedUsers = selectedUsers.filter(function (el: string) {
                return el;
            });
            console.log(selectedUsers);
            scrumProp[parseInt(dropdownProps!.id)].SelectedMembers = selectedUsers;
            scrumProp[parseInt(dropdownProps!.id)].UserPrincipalNames = selectedUsers.map(user => user.content).join(",");
            this.setState({
                scrums: scrumProp,
            });

        }
    };

    /**
    *  Gets called when user selects start time.
    * */
    private setStartTime = (e: any, dropdownProps?: any) => {
        let Scrumprop = this.state.scrums;
        if (dropdownProps) {
            let selectedStartTime = dropdownProps.value;
            Scrumprop[parseInt(dropdownProps!.id)].StartTime = selectedStartTime;
            this.setState({
                scrums: Scrumprop,
            });
        }
    };

    /**
    *  Gets called when user selects time zone.
    * */
    private setTimeZone = (e: any, dropdownProps?: any) => {
        let Scrumprop = this.state.scrums;
        if (dropdownProps) {
            let selectedTimeZone = dropdownProps.value;
            Scrumprop[parseInt(dropdownProps!.id)].TimeZone = selectedTimeZone.timeZoneId;
            Scrumprop[parseInt(dropdownProps!.id)].SelectedTimeZone = selectedTimeZone;
            this.setState({
                scrums: Scrumprop,
            });
        }
    };

    /**
    *  Gets called when user selects channel.
    * */
    private setChannel = (e: any, dropdownProps?: any) => {
        let Scrumprop = this.state.scrums;
        if (dropdownProps) {
            {
                let selectedchannelid;
                let selectedchannel = dropdownProps.value;
                let channel = channels.find(channel => channel.header === selectedchannel.header);
                selectedchannelid = channel?.channelId;
                Scrumprop[parseInt(dropdownProps!.id)].ChannelId = selectedchannelid;
                Scrumprop[parseInt(dropdownProps!.id)].ChannelName = String(selectedchannel.header);
                this.setState({
                    scrums: Scrumprop,
                });
            }
        }
    };

    /**
    *  Gets called when user clicks on delete scrum button.
    * */
    private removeScrum = (id: number) => () => {
        let scrums = this.state.scrums;
        let deletedScrum = (scrums.find((scrum, scrumTeamConfigId) => id === scrumTeamConfigId)) as IScrumProps;
        this.state.deletedScrums.push(deletedScrum);

        scrums = scrums.filter((scrum, scrumTeamConfigId) => id !== scrumTeamConfigId);
        this.setState({ scrums: scrums });
    };

    /**
    *  Gets called when user clicks on save button.
    * */
    private saveScrumSettings = async () => {
        this.setState({ isSaveScrumSettingsLoading: true });
        this.setState({
            showError: false,
            errorMessage: "",
        });

        if (this.validateScrumDetails())
        {
            // Store or delete scrum configuration details in table storage.
            let response = await this.saveScrumConfigurationDetails();
            if (response) {
                this.setState({ isSaveScrumSettingsLoading: false });
                microsoftTeams.getContext((context) => {
                    microsoftTeams.tasks.submitTask();
                });
            }
        }
    };

    /**
    *  Validates user inputs before saving scrum configuration data in storage.
    * */
    private validateScrumDetails() {
        let Scrumprop = this.state.scrums;
        let errorMessage: string = "";
        for (let i = 0; i < this.state.scrums.length; i++) {
            Scrumprop[i].TeamId = teamId;
            Scrumprop[i].AADGroupID = groupID;
            let member: any = this.state.teamMembers.find(element => element.aadobjectid === this.userObjectId);
            Scrumprop[i].CreatedBy = member.header;

            if (Scrumprop[i].ScrumTeamName) {
                Scrumprop[i].ScrumTeamName = Scrumprop[i].ScrumTeamName.trim();
            }

            if (!errorMessage) {
                if (!Scrumprop[i].ScrumTeamName) {
                    errorMessage = this.state.resourceStrings.teamNameValidationText;
                }
                else if (!Scrumprop[i].UserPrincipalNames || Scrumprop[i].UserPrincipalNames.split(',').length < 2) {
                    errorMessage = this.state.resourceStrings.teamMembersValidationText;
                }
                else if (!Scrumprop[i].StartTime) {
                    errorMessage = this.state.resourceStrings.startTimeValidationText;
                }
                else if (!Scrumprop[i].TimeZone) {
                    errorMessage = this.state.resourceStrings.timeZoneValidationText;
                }
                else if (!Scrumprop[i].ChannelName) {
                    errorMessage = this.state.resourceStrings.channelNameValidationText;
                }
                else {
                    let scrums = Scrumprop.filter((scrum, scrumTeamConfigId) => i !== scrumTeamConfigId);
                    let duplicateScrum = scrums.find(scrum => scrum.ScrumTeamName === Scrumprop[i].ScrumTeamName && scrum.ChannelName === Scrumprop[i].ChannelName);
                    if (duplicateScrum) {
                        errorMessage = this.state.resourceStrings.duplicateScrumValidationText;
                    }
                }
            }
        }

        this.setState({
            scrums: Scrumprop
        });

        if (errorMessage) {
            this.setState({ showError : true, errorMessage: errorMessage, isSaveScrumSettingsLoading: false });
            return false;
        }
        else {
            this.setState({ showError: false, errorMessage: ""});
            return true;
        }
    }

    /**
    *  Stores scrum configuration details in table storage.
    * */
    private saveScrumConfigurationDetails = async () => {
        if (this.state.scrums.length > 0) {
            const saveScrumDetailsResponse = await saveScrumConfigurationDetails(this.customAPIAuthenticationToken!, this.state.scrums)
            if (saveScrumDetailsResponse.status !== 200 && saveScrumDetailsResponse.status !== 204) {
                this.setState({ isSaveScrumSettingsLoading: false, errorMessage: this.state.resourceStrings.errorMessage });
                handleError(saveScrumDetailsResponse, this.customAPIAuthenticationToken);
                return false;
            }
        }
        
        this.setState({ isSaveScrumSettingsLoading: true });
        return await this.deleteScrumConfigurationDetails();
    }

    /**
    *  Deletes scrum configuration details from table storage.
    * */
    private deleteScrumConfigurationDetails = async () => {
        if (this.state.deletedScrums.length > 0) {
            this.appInsights.trackTrace({ message: `'deleteScrumConfigurationDetails' - Request initiated`, severityLevel: SeverityLevel.Information, properties: { UserEmail: this.userEmail } });

            // Delete scrum configuration details from table storage.
            const saveScrumDetailsResponse = await deleteScrumConfigurationDetails(this.customAPIAuthenticationToken!, this.state.deletedScrums);
            if (saveScrumDetailsResponse.status === 200 || saveScrumDetailsResponse.status === 204) {
                return true;
            }
            else {
                this.setState({ isSaveScrumSettingsLoading: false, errorMessage: this.state.resourceStrings.errorMessage });
                handleError(saveScrumDetailsResponse, this.customAPIAuthenticationToken);
                return false;
            }
        }

        return true;
    }

    /**
    *  Renders settings layout on UI.
    * */
    renderSettings() {
        return (
            <>
                <Table>
                    <Header resourceStrings={this.state.resourceStrings} />
                </Table>
                <Table aria-label="table" className="create-scrum-table">
                    {this.state.scrums.map((scrum, id) => (
                        <Table.Row className="table-row" key={id} >
                            <div>
                                <Table.Cell>
                                    <Flex gap="gap.small">
                                        <Input
                                            className="table-cell1"
                                            type="text"
                                            name="teamName"
                                            maxLength={35}
                                            aria-label={this.state.resourceStrings.teamNameTitle}
                                            placeholder={this.state.resourceStrings.teamNameTitle}
                                            value={scrum.ScrumTeamName}
                                            fluid
                                            onChange={event => this.teamNameChange(id, event)}
                                            title={scrum.ScrumTeamName}
                                        />
                                    </Flex>
                                </Table.Cell>
                            </div>
                            <div>
                                <Table.Cell>
                                    <Flex gap="gap.smaller">
                                        <Checkbox className="table-cell2" toggle id={id + ""} onClick={this.setScrumStatus} checked={scrum.IsActive} />
                                    </Flex>
                                </Table.Cell>
                            </div>
                            <div>
                                <Table.Cell >
                                    <Dropdown
                                        fluid
                                        multiple
                                        search
                                        className="table-cell3"
                                        onChange={this.memberSelectionChange}
                                        items={this.state.teamMembers}
                                        placeholder={this.state.resourceStrings.selectUserPlaceholder}
                                        noResultsMessage={this.state.resourceStrings.noMatchesFoundText}
                                        value={scrum.SelectedMembers}
                                        id={id + ""}
                                    />
                                </Table.Cell>
                            </div>
                            <div>
                                <Table.Cell>
                                    <Flex gap="gap.smaller">
                                        <Flex.Item size="size.half">
                                            <TimeSuggestion /*minDate={moment()} dateTime={moment()}*/
                                                placeholder={this.state.resourceStrings.selectTimePlaceholder}
                                                onTimeChange={this.setStartTime}
                                                selectedValue={scrum.StartTime}
                                                id={id + ""}
                                            />
                                        </Flex.Item>
                                    </Flex>
                                </Table.Cell>
                            </div>
                            <div>
                                <Table.Cell >
                                    <Dropdown fluid className="table-cell5" placeholder={this.state.resourceStrings.selectTimeZonePlaceholder} onChange={this.setTimeZone} value={scrum.SelectedTimeZone} title={scrum.TimeZone} id={id + ""} items={timeZones} />
                                </Table.Cell>
                            </div>
                            <div>
                                <Table.Cell>
                                    <Dropdown fluid className="table-cell6" placeholder={this.state.resourceStrings.selectChannelPlaceholder} onChange={this.setChannel} value={scrum.ChannelName} title={scrum.ChannelName} id={id + ""} items={channels} />
                                </Table.Cell>
                            </div>
                            <div>
                                <Table.Cell>
                                    <Button size="medium" circular icon={<TrashCanIcon />} title={this.state.resourceStrings.DeleteButtonText} primary onClick={this.removeScrum(id)} />
                                </Table.Cell>
                            </div>
                            </Table.Row>
                    ))}
                </Table>
                <Footer
                    resourceStrings={this.state.resourceStrings}
                    errorMessage={this.state.errorMessage}
                    isSaveScrumSettingsLoading={this.state.isSaveScrumSettingsLoading}
                    addNewScrum={this.addNewScrum}
                    saveScrumSettings={this.saveScrumSettings}
                />
                
            </>
        );
    }

    /**
    *  Renders settings layout or loader on UI depending upon data is fetched from storage.
    * */
    render() {
        let contents = this.state.loading
            ? <p><em><Loader /></em></p>
            : this.renderSettings();
        if (this.state.resourceStringsLoaded) {
            return (
                <div className="container-div">
                    {contents}
                </div>
            );
        }
        else {
            return (
                <Loader />
            );
        }
    }
}

export default withAITracking(reactPlugin, Settings);