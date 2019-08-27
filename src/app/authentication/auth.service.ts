import * as Msal from 'msal';
import { AppComponent } from '../app.component';
const { oneDS } = (window as any);

const loginType = getLoginType();

const appInsights = new oneDS.ApplicationInsights();

const config = {
    instrumentationKey: 'a800ae98-f89e-4f96-b491-cf1b8a989bff',
    channelConfiguration: { // Post channel configuration
        eventsLimitInMem: 50,
    },
    propertyConfiguration: { // Properties Plugin configuration
        userAgent: 'Custom User Agent',
    },
    webAnalyticsConfiguration: { // Web Analytics Plugin configuration
        autoCapture: {
            jsError: true,
        },
    },
};

//Initialize SDK
appInsights.initialize(config, []);

export const collectLogs = (error: any): void => {
    if (appInsights) {
        appInsights.trackException({ exception: error });
    }
};

export function logout(userAgentApp: Msal.UserAgentApplication) {
    userAgentApp.logout();
}

// tslint:disable-next-line: max-line-length
export async function getTokenSilent(userAgentApp: Msal.UserAgentApplication, scopes: string[]): Promise<Msal.AuthResponse> {
    return userAgentApp.acquireTokenSilent({ scopes: generateUserScopes(scopes) });
}

export async function login(userAgentApp: Msal.UserAgentApplication) {
    const loginRequest = {
        scopes: generateUserScopes(),
        prompt: 'select_account',
    };
    if (loginType === 'POPUP') {
        try {
            const response = await userAgentApp.loginPopup(loginRequest);
            return response;
        } catch (error) {
            throw error;
        }
    } else if (loginType === 'REDIRECT') {
        await userAgentApp.loginRedirect(loginRequest);
    }
}

export async function acquireNewAccessToken(userAgentApp: Msal.UserAgentApplication, scopes: string[] = []) {
    const hasScopes = (scopes.length > 0);
    let listOfScopes = AppComponent.Options.DefaultUserScopes;
    if (hasScopes) {
        listOfScopes = scopes;
    }
    return getTokenSilent(userAgentApp, generateUserScopes(listOfScopes)).catch((error) => {
        if (requiresInteraction(error.errorCode)) {
            if (loginType === 'POPUP') {
                try {
                    return userAgentApp.acquireTokenPopup({ scopes: generateUserScopes(listOfScopes) });
                } catch (error) {
                    throw error;
                }
            } else if (loginType === 'REDIRECT') {
                userAgentApp.acquireTokenRedirect({ scopes: generateUserScopes(listOfScopes) });
            }
        }
    });
}

export function getAccount(userAgentApp: Msal.UserAgentApplication) {
    return userAgentApp.getAccount();
}

export function generateUserScopes(userScopes = AppComponent.Options.DefaultUserScopes) {
    const graphMode = JSON.parse(localStorage.getItem('GRAPH_MODE'));
    if (graphMode === null) {
        return userScopes;
    }
    const graphUrl = localStorage.getItem('GRAPH_URL');
    const reducedScopes = userScopes.reduce((newScopes, scope) => {
        if (scope === 'openid' || scope === 'profile') {
            return newScopes += scope + ' ';
        }
        return newScopes += graphUrl + '/' + scope + ' ';
    }, '');

    const scopes = reducedScopes.split(' ').filter((scope) => {
        return scope !== '';
    });
    return scopes;
}

export function requiresInteraction(errorCode) {
    if (!errorCode || !errorCode.length) {
        return false;
    }
    return errorCode === 'consent_required' ||
        errorCode === 'interaction_required' ||
        errorCode === 'login_required' ||
        errorCode === 'token_renewal_error';
}

export function getLoginType() {
    /**
     * Always redirects because of transient issues caused by showing a pop up. Graph Explorer
     * loses hold of the iframe Pop Up
     */
    return 'REDIRECT';
}
