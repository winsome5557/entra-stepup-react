import React from "react";
import ReactDOM from "react-dom/client";
import { PublicClientApplication } from "@azure/msal-browser";
import { MsalProvider } from "@azure/msal-react";
import { msalConfig } from "./authConfig";
import App from "./App";
import "./style.css";

const msal = new PublicClientApplication(msalConfig);
await msal.initialize();

if (!msal.getActiveAccount() && msal.getAllAccounts().length) {
  msal.setActiveAccount(msal.getAllAccounts()[0]);
}

ReactDOM.createRoot(document.getElementById("root")).render(
  <React.StrictMode>
    <MsalProvider instance={msal}><App /></MsalProvider>
  </React.StrictMode>
);

msal.handleRedirectPromise()
  .then((redirectResult) => {
    if (redirectResult?.account) msal.setActiveAccount(redirectResult.account);
  })
  .catch((error) => {
    console.error("MSAL redirect handling failed", error);
  });
