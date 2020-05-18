/*
    <copyright file="router.tsx" company="Microsoft">
    Copyright (c) Microsoft. All rights reserved.
    </copyright>
*/

import * as React from "react";
import { BrowserRouter, Route, Switch } from "react-router-dom";
import Settings from '../components/settings';
import ErrorPage from '../components/error-page';

export const AppRoute: React.FunctionComponent<{}> = () => {
	return (
		<BrowserRouter>
			<Switch>
				<Route path='/settings' component={Settings} />
				<Route exact path="/error" component={ErrorPage} />
			</Switch>
		</BrowserRouter>
	);
};