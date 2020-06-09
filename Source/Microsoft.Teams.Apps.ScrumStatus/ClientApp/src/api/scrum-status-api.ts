/*
    <copyright file="scrum-status-api.ts" company="Microsoft">
    Copyright (c) Microsoft. All rights reserved.
    </copyright>
*/

import axios from "./axios-decorator";
import { IScrumProps, ITeamDetails } from "../models/type";
import { AxiosResponse } from "axios";

const baseAxiosUrl = window.location.origin;

/**
* Get all team members.
* @param token {String | Null}  Custom JWT token.
* @param teamId {String} Team ID for getting members.
*/
export const getTeamDetails = async (token: string, teamId: string): Promise<AxiosResponse<ITeamDetails>> => {

    let url = baseAxiosUrl + "/api/scrumconfiguration/teamdetails?teamId=" + teamId;
    let teamDetailsResponse = await axios.get(url, token);
    return teamDetailsResponse;
}

/**
* Get scrum configuration details by Azure Active Directory group id of team.
* @param token {String | Null} Custom JWT token.
* @param groupID {String | Null} Azure Active Directory group Id.
*/
export const getScrumConfigurationDetailsbyAADGroupID = async (token: string, groupID: string): Promise<AxiosResponse<IScrumProps[]>> => {

    let url = baseAxiosUrl + "/api/scrumconfiguration/scrumconfigurationdetails?groupId=" + groupID;
    let scrumConfigurationDetailsResponse = await axios.get(url, token);
    return scrumConfigurationDetailsResponse;
}

/**
* Get system time zone information.
* @param token {String | Null} Custom JWT token.
*/
export const getTimeZoneInfo = async (token: string): Promise<any> => {

    let url = baseAxiosUrl + "/api/scrumconfiguration/timezoneinfo";
    let timeZoneInfoResponse = await axios.get(url, token);
    return timeZoneInfoResponse;
}

/**
* Get localized resource strings from API.
* @param token {String | Null} Custom JWT token.
* @param locale {String | Null} Current client culture.
*/
export const getResourceStrings = async (token: string, locale?: string | null): Promise<any> => {
    let url = baseAxiosUrl + "/api/resource/resourcestrings";
    let resourceStringsResponse = await axios.get(url, token, locale);
    return resourceStringsResponse;
}

/**
* Get localized resource strings from API.
* @param token {String | Null} Custom JWT token.
* @param scrumConfigurationDetails {Object} scrum configuration details to be stored in storage.
*/
export const saveScrumConfigurationDetails = async (token: string, scrumConfigurationDetails: {}): Promise<AxiosResponse<void>> => {
    let url = baseAxiosUrl + "/api/scrumconfiguration/scrumconfigurationdetails";
    let saveScrumConfigurationDetailsResponse = await axios.post(url, scrumConfigurationDetails, token);
    return saveScrumConfigurationDetailsResponse;
}

/**
* Get localized resource strings from API.
* @param token {String | Null} Custom JWT token.
* @param scrumConfigurationDetails {Object} Scrum configuration details to be deleted from storage.
*/
export const deleteScrumConfigurationDetails = async (token: string, scrumConfigurationDetails: {}): Promise<AxiosResponse<void>> => {
    let url = baseAxiosUrl + "/api/scrumconfiguration/scrumconfigurationdetails";
    let deleteScrumConfigurationDetailsResponse = await axios.delete(url, scrumConfigurationDetails, token);
    return deleteScrumConfigurationDetailsResponse;
}

/**
* Handle error occurred during API call.
* @param error {Object} Error response object
*/
export const handleError = (error: any, token: any): any => {
	const errorStatus = error.status;
	if (errorStatus === 403) {
        window.location.href = "/error?code=403&token=" + token;
    }
    else if (errorStatus === 401) {
        window.location.href = "/error?code=401&token=" + token;
    }
    else {
        window.location.href = "/error?token=" + token;
	}
}