/*
    <copyright file="app.tsx" company="Microsoft">
    Copyright (c) Microsoft. All rights reserved.
    </copyright>
*/

import * as React from "react";
import { AppRoute } from "./router/router";
import { Provider, themes } from "@fluentui/react-northstar";
import Constants from "./constants";
import moment from 'moment';
import 'moment/min/locales.min';
import { getUserLocale } from './localization/translate';

moment.locale(getUserLocale());

export interface IAppState {
	theme: string;
	themeStyle: any;
}

export default class App extends React.Component<{}, IAppState> {
	theme?: string | null;
	constructor(props: any) {
		super(props);
		let search = window.location.search;
		let params = new URLSearchParams(search);
		this.theme = params.get("theme");

		this.state = {
			theme: this.theme ? this.theme : Constants.default,
			themeStyle: themes.teams,
		}
	}

	/** Called once component is mounted. */
	async componentDidMount() {
		this.setThemeComponent();
	}

	/** Set theme for all components */
	public setThemeComponent = () => {
		if (this.state.theme === Constants.dark) {
			this.setState({
				themeStyle: themes.teamsDark
			});
		} else if (this.state.theme === Constants.contrast) {
			this.setState({
				themeStyle: themes.teamsHighContrast
			});
		} else {
			this.setState({
				themeStyle: themes.teams
			});
		}
	}

	/**
	* Renders the component
	*/
	public render(): JSX.Element {
		return (
			<Provider theme={this.state.themeStyle}>
				<div>
					<AppRoute />
				</div>
			</Provider>
		);
	}
}