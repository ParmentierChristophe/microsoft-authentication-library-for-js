import React, { Component } from "react";
import * as Msal from "msal";

const requiresInteraction = errorMessage => {
    if (!errorMessage || !errorMessage.length) {
        return false;
    }

    return (
        errorMessage.indexOf("consent_required") > -1 ||
        errorMessage.indexOf("interaction_required") > -1 ||
        errorMessage.indexOf("login_required") > -1
    );
};

const callMSGraph = async (url, accessToken) => {
    const response = await fetch(url, {
        headers: {
            Authorization: `Bearer ${accessToken}`
        }
    });

    return response.json();
};

const isIE = () => {
    const ua = window.navigator.userAgent;
    const msie = ua.indexOf("MSIE ") > -1;
    const msie11 = ua.indexOf("Trident/") > -1;

    // If you as a developer are testing using Edge InPrivate mode, please add "isEdge" to the if check
    // const isEdge = ua.indexOf("Edge/") > -1;

    return msie || msie11;
};

const loginRequest = {
    scopes: ["openid", "profile", "User.Read"]
};

const tokenRequest = {
    scopes: ["Mail.Read"]
};

const graphConfig = {
    graphMeEndpoint: "https://graph.microsoft.com/v1.0/me",
    graphMailEndpoint: "https://graph.microsoft.com/v1.0/me/messages"
};

const msalApp = new Msal.UserAgentApplication({
    auth: {
        clientId: "245e9392-c666-4d51-8f8a-bfd9e55b2456",
        authority: "https://login.microsoftonline.com/common",
        validateAuthority: true,
        postLogoutRedirectUri: "http://localhost:3000"
    },
    cache: {
        cacheLocation: "sessionStorage",
        storeAuthStateInCookie: isIE()
    }
});

// If you support IE, our recommendation is that you sign-in using Redirect APIs
const useRedirectFlow = isIE();

export default C =>
    class AuthProvider extends Component {
        constructor(props) {
            super(props);
            this.state = {
                account: msalApp.getAccount()
            };

            msalApp.handleRedirectCallback(this.authRedirectCallBack);
        }

        async acquireTokenAndCallMSGraph(endpoint, request, redirect) {
            const tokenResponse = await msalApp
                .acquireTokenSilent(request)
                .catch(error => {
                    // Call acquireTokenPopup (popup window) in case of acquireTokenSilent failure due to consent or interaction required ONLY
                    if (requiresInteraction(error.errorCode)) {
                        return redirect
                            ? msalApp.acquireTokenRedirect(request)
                            : msalApp.acquireTokenPopup(request);
                    }
                });

            if (tokenResponse) {
                const graphResponse = await callMSGraph(
                    endpoint,
                    tokenResponse.accessToken
                );

                this.graphAPICallback(graphResponse);
            }
        }

        async signIn(redirect) {
            if (redirect) {
                return msalApp.loginRedirect(loginRequest);
            }

            await msalApp.loginPopup(loginRequest);

            await this.acquireTokenAndCallMSGraph(
                graphConfig.graphMeEndpoint,
                loginRequest
            );

            this.setState({
                account: msalApp.getAccount()
            });
        }

        signOut() {
            msalApp.logout();
        }

        async readMail() {
            return this.acquireTokenAndCallMSGraph(
                graphConfig.graphMailEndpoint,
                tokenRequest,
                useRedirectFlow
            );
        }

        async authRedirectCallBack(error, response) {
            if (error) {
                console.error(error);
                return;
            }

            if (response.tokenType === "access_token") {
                return callMSGraph(
                    graphConfig.graphMeEndpoint,
                    response.accessToken
                ).then(this.graphAPICallback);
            }
        }

        graphAPICallback(graphData) {
            this.setState({
                graphData
            });
        }

        componentDidMount() {
            if (
                this.state.account &&
                !msalApp.isCallback(window.location.hash)
            ) {
                this.acquireTokenAndCallMSGraph(
                    graphConfig.graphMeEndpoint,
                    loginRequest,
                    useRedirectFlow
                );
            }
        }

        render() {
            return (
                <C
                    {...this.props}
                    signIn={() => this.signIn(useRedirectFlow)}
                    signOut={() => this.signOut()}
                    account={this.state.account}
                    readMail={() => this.readMail()}
                    graphData={this.state.graphData}
                />
            );
        }
    };
