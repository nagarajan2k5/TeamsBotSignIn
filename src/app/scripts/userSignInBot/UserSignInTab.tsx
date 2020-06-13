import * as React from "react";
import { Provider, Flex, Text, Button, Header } from "@fluentui/react";
import TeamsBaseComponent, { ITeamsBaseComponentProps, ITeamsBaseComponentState } from "msteams-react-base-component";
import * as microsoftTeams from "@microsoft/teams-js";

/**
 * State for the userSignInTab React component
 */
export interface IUserSignInTabState extends ITeamsBaseComponentState {

}

/**
 * Properties for the userSignInTab React component
 */
export interface IUserSignInTabProps extends ITeamsBaseComponentProps {

}

/**
 * Implementation of the userSignIn content page
 */
export class UserSignInTab extends TeamsBaseComponent<IUserSignInTabProps, IUserSignInTabState> {

    public componentWillMount() {
        this.updateTheme(this.getQueryVariable("theme"));

        if (this.inTeams()) {
            microsoftTeams.initialize();
            microsoftTeams.registerOnThemeChangeHandler(this.updateTheme);
        }
    }

    /**
     * The render() method to create the UI of the tab
     */
    public render() {
        return (
            <Provider theme={this.state.theme}>
                <Flex fill={true} column styles={{
                    padding: ".8rem 0 .8rem .5rem"
                }}>
                    <Flex.Item>
                        <Header content="Welcome to the UserSignIn bot page" />
                    </Flex.Item>
                    <Flex.Item>
                        <div>
                            <Text content="TODO: Add you content here" />
                        </div>
                    </Flex.Item>
                    <Flex.Item styles={{
                        padding: ".8rem 0 .8rem .5rem"
                    }}>
                        <Text size="smaller" content="(C) Copyright msnextlife" />
                    </Flex.Item>
                </Flex>
            </Provider>
        );
    }
}
