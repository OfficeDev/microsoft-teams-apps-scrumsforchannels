/*
    <copyright file="error-page.tsx" company="Microsoft">
    Copyright (c) Microsoft. All rights reserved.
    </copyright>
*/

import * as React from "react";
import { Text, Flex, Label } from "@fluentui/react-northstar";
import { ErrorIcon } from '@fluentui/react-icons-northstar';
import { getResourceStrings } from "../api/scrum-status-api";

interface errorPageState {
    resourceStrings: any | "",
}

export default class ErrorPage extends React.Component<{}, errorPageState> {
    code: string | null = null;
    token: string | null = null;

    constructor(props: any) {
        super(props);
        let search = window.location.search;
        let params = new URLSearchParams(search);
        this.token = params.get("token");
        this.code = params.get("code");
        this.state = {
            resourceStrings: {}
        };
    }

    /** Called once component is mounted. */
    async componentDidMount() {
        await this.getResourceStrings();
    }

    /** 
    *  Get resource strings according to user locale.
    * */
    getResourceStrings = async () => {
        const resourceStringsResponse = await getResourceStrings(this.token!);

        if (resourceStringsResponse.status === 200) {
            this.setState({ resourceStrings: resourceStringsResponse.data });
        }
    }

    /**
     * Render error page.
     * */
    render() {
        let message = this.state.resourceStrings.errorMessage;
        if (this.code === "401") {
            message = `${this.state.resourceStrings.unauthorizedAccess}`;
        }

        return (
            <div className="container-div">
                <Flex gap="gap.small" hAlign="center" vAlign="center" className="error-container">
                    <Flex gap="gap.small" hAlign="center" vAlign="center">
                        <Flex.Item>
                            <div
                                style={{
                                    position: "relative",
                                }}
                            >
                                <Label icon={<ErrorIcon />} />
                            </div>
                        </Flex.Item>

                        <Flex.Item grow>
                            <Flex column gap="gap.small" vAlign="stretch">
                                <div>
                                    <Text weight="bold" error content={message} /><br />
                                </div>
                            </Flex>
                        </Flex.Item>
                    </Flex>
                </Flex>
            </div>
        );
    }
}