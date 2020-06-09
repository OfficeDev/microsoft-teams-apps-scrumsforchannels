/*
    <copyright file="index.tsx" company="Microsoft">
    Copyright (c) Microsoft. All rights reserved.
    </copyright>
*/

import * as React from "react";
import * as ReactDOM from "react-dom";
import { BrowserRouter as Router } from "react-router-dom";
import App from "./app";
import { IntlProvider } from 'react-intl';
import { getUserLocale } from './localization/translate';

const userLocale = getUserLocale();

ReactDOM.render(
    <IntlProvider locale={userLocale}>
	    <Router>
		    <App />
        </Router>
    </IntlProvider>, document.getElementById("root")
);
