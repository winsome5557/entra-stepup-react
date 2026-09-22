export const settings = {
  clientId: import.meta.env.VITE_ENTRA_CLIENT_ID,
  tenantId: import.meta.env.VITE_ENTRA_TENANT_ID,
  redirectUri: import.meta.env.VITE_ENTRA_REDIRECT_URI || window.location.origin,
  scopes: (import.meta.env.VITE_ENTRA_SCOPES || "openid profile User.Read").split(/\s+/),
  authContext: (import.meta.env.VITE_ENTRA_AUTH_CONTEXT || "c3").toLowerCase()
};

export const msalConfig = {
  auth: {
    clientId: settings.clientId,
    authority: `https://login.microsoftonline.com/${settings.tenantId}`,
    redirectUri: settings.redirectUri,
    postLogoutRedirectUri: settings.redirectUri,
    clientCapabilities: ["CP1"]
  },
  cache: { cacheLocation: "sessionStorage" }
};

export const normalLoginRequest = {
  scopes: settings.scopes
};

export function stepUpRequest(account) {
  return {
    scopes: settings.scopes,
    account,
    // Authentication Context is requested via the OAuth/OIDC "claims" parameter.
    claims: JSON.stringify({
      access_token: {
        acrs: {
          essential: true,
          value: settings.authContext
        }
      }
    }),
  };
}
