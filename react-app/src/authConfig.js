export const msalConfig = {
  auth: {
    clientId: "e8e675d2-e3bd-4bea-8cb2-25609539a365", // Replace with your Azure AD app registration client ID
    authority: "https://login.microsoftonline.com/common",
    redirectUri: "http://localhost:3000", // Replace with your redirect URI
  },
  cache: {
    cacheLocation: "sessionStorage",
    storeAuthStateInCookie: false,
  },
};

export const graphConfig = {
  graphMeEndpoint: "https://graph.microsoft.com/v1.0/me",
};
