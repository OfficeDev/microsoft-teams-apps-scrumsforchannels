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
    let url = baseAxiosUrl + "/api/scrummaster/teamdetails?teamId=" + teamId;
    let teamDetailsResponse = await axios.get(url, token);
    return teamDetailsResponse;
}

/**
* Get scrum master details by Azure Active Directory group id of team.
* @param token {String | Null} Custom JWT token.
* @param groupID {String | Null} Azure Active Directory group Id.
*/
export const getScrumMasterDetailsbyAADGroupID = async (token: string, groupID: string): Promise<AxiosResponse<IScrumProps[]>> => {
    let url = baseAxiosUrl + "/api/scrummaster/scrummasterdetails?groupId=" + groupID;
    let scrumMasterDetailsResponse = await axios.get(url, token);
    return scrumMasterDetailsResponse;
}

/**
* Get system time zone information.
* @param token {String | Null} Custom JWT token.
*/
export const getTimeZoneInfo = async (token: string): Promise<any> => {
    let url = baseAxiosUrl + "/api/scrummaster/timezoneinfo";
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
* @param scrumMasterDetails {Object} Scrum master details to be stored in storage.
*/
export const saveScrumMasterDetails = async (token: string, scrumMasterDetails: {}): Promise<AxiosResponse<void>> => {
    let url = baseAxiosUrl + "/api/scrummaster/scrummasterdetails";
    let saveScrumMasterDetailsResponse = await axios.post(url, scrumMasterDetails, token);
    return saveScrumMasterDetailsResponse;
}

/**
* Get localized resource strings from API.
* @param token {String | Null} Custom JWT token.
* @param scrumMasterDetails {Object} Scrum master details to be deleted from storage.
*/
export const deleteScrumMasterDetails = async (token: string, scrumMasterDetails: {}): Promise<AxiosResponse<void>> => {
    let url = baseAxiosUrl + "/api/scrummaster/scrummasterdetails";
    let deleteScrumMasterDetailsResponse = await axios.delete(url, scrumMasterDetails, token);
    return deleteScrumMasterDetailsResponse;
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