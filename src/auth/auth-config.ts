
import { LogLevel } from '@azure/msal-browser';
export const msalConfig = {
    auth: {
        clientId: '5b7b3cf8-2b5e-49df-8f33-e62f2a17eeba', // This is the ONLY mandatory field that you need to supply.
        authority: 'https://login.microsoftonline.com/01c6c7ba-3525-4f90-ae9a-4dc8f5f82790', // Replace the placeholder with your tenant subdomain 
        redirectUri: '/', // Points to window.location.origin. You must register this URI on Microsoft Entra admin center/App Registration.
        postLogoutRedirectUri: '/', // Indicates the page to navigate after logout.
        navigateToLoginRequestUrl: false, // If "true", will navigate back to the original request location before processing the auth code response.
    },
    cache: {
        cacheLocation: 'sessionStorage', // Configures cache location. "sessionStorage" is more secure, but "localStorage" gives you SSO between tabs.
        storeAuthStateInCookie: false, // Set this to "true" if you are having issues on IE11 or Edge
    },
    system: {
        loggerOptions: {
            loggerCallback: (level: any, message: any, containsPii: any) => {
                if (containsPii) {
                    return;
                }
                switch (level) {
                    case LogLevel.Error:
                        console.error(message);
                        return;
                    case LogLevel.Info:
                        console.info(message);
                        return;
                    case LogLevel.Verbose:
                        console.debug(message);
                        return;
                    case LogLevel.Warning:
                        console.warn(message);
                        return;
                    default:
                        return;
                }
            },
        },
    },
};
export const loginRequest = {
    scopes: [],
};
