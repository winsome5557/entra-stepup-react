import React, { useEffect, useMemo, useState } from "react";
import { EventType } from "@azure/msal-browser";
import { useIsAuthenticated, useMsal } from "@azure/msal-react";
import { normalLoginRequest, settings, stepUpRequest } from "./authConfig";

function decodeJwt(token) {
  try {
    const payload = token.split(".")[1].replace(/-/g, "+").replace(/_/g, "/");
    return JSON.parse(decodeURIComponent(atob(payload).split("").map(c =>
      "%" + ("00" + c.charCodeAt(0).toString(16)).slice(-2)).join("")));
  } catch { return {}; }
}

function eventName(eventType) {
  return Object.entries(EventType).find(([, value]) => value === eventType)?.[0] || String(eventType);
}

function authorizeDetails(request) {
  return {
    method: "GET",
    endpoint: `https://login.microsoftonline.com/${settings.tenantId}/oauth2/v2.0/authorize`,
    query: {
      client_id: settings.clientId,
      redirect_uri: settings.redirectUri,
      response_type: "code",
      scopes: request.scopes,
      prompt: request.prompt || "default",
      claims: request.claims ? JSON.parse(request.claims) : "not requested"
    },
    note: "MSAL adds state, nonce, PKCE, and other transient parameters at runtime."
  };
}

export default function App() {
  const { instance, accounts } = useMsal();
  const authenticated = useIsAuthenticated();
  const [status, setStatus] = useState("Ready");
  const [claims, setClaims] = useState(null);
  const [consoleEntries, setConsoleEntries] = useState([]);
  const account = instance.getActiveAccount() || accounts[0];

  function recordConsole(type, details) {
    setConsoleEntries((entries) => [...entries, {
      time: new Date().toLocaleTimeString(),
      type,
      details
    }].slice(-100));
  }

  useEffect(() => {
    const callbackId = instance.addEventCallback((message) => {
      recordConsole("MSAL event", {
        event: eventName(message.eventType),
        interactionType: message.interactionType || "n/a",
        correlationId: message.correlationId || "n/a",
        error: message.error?.errorCode || null
      });
    });

    const seenResources = new Set();
    const captureResources = () => {
      performance.getEntriesByType("resource")
        .map((entry) => entry.name)
        .filter((name) => name.includes("login.microsoftonline.com"))
        .filter((name) => !seenResources.has(name))
        .forEach((name) => {
          seenResources.add(name);
          recordConsole("Observed Entra resource", { url: name });
        });
    };
    const resourceTimer = window.setInterval(captureResources, 500);
    captureResources();

    return () => {
      instance.removeEventCallback(callbackId);
      window.clearInterval(resourceTimer);
    };
  }, [instance]);

  const configView = useMemo(() => ({
    tenantId: settings.tenantId,
    clientId: settings.clientId,
    redirectUri: settings.redirectUri,
    scopes: settings.scopes.join(" "),
    authenticationContext: settings.authContext
  }), []);

  async function signIn() {
    try {
      recordConsole("Authorize request", authorizeDetails(normalLoginRequest));
      recordConsole("Token exchange", {
        method: "POST",
        endpoint: `https://login.microsoftonline.com/${settings.tenantId}/oauth2/v2.0/token`,
        note: "MSAL performs this code-to-token exchange after the authorization response. Parameters are POST body fields, not query string parameters."
      });
      setStatus("Opening standard Entra sign-in...");
      // Standard Authentication
      const result = await instance.loginPopup(normalLoginRequest);
      instance.setActiveAccount(result.account);
      setClaims(decodeJwt(result.accessToken || result.idToken));
      setStatus("Standard authentication completed.");
    } catch (e) {
      setStatus(`Sign-in failed: ${e.errorCode || e.message}`);
    }
  }

  async function stepUp() {
    const current = instance.getActiveAccount() || accounts[0];
    if (!current) {
      setStatus("Sign in first.");
      return;
    }
    try {
      const request = stepUpRequest(current);
      recordConsole("Authorize request", authorizeDetails(request));
      recordConsole("Authentication Context", {
        requested: settings.authContext,
        claims: JSON.parse(request.claims),
        note: "This is sent as the OAuth claims parameter on /authorize."
      });
      recordConsole("Token exchange", {
        method: "POST",
        endpoint: `https://login.microsoftonline.com/${settings.tenantId}/oauth2/v2.0/token`,
        note: "The token request uses POST body fields. The returned access-token claims are shown below."
      });
      setStatus(`Requesting Authentication Context ${settings.authContext}...`);
      // Stepup Authentication
      const result = await instance.acquireTokenPopup(request);
      const tokenClaims = decodeJwt(result.accessToken);
      setClaims(tokenClaims);
      const acrs = tokenClaims.acrs;
      const satisfied = Array.isArray(acrs)
        ? acrs.map(String).map(x => x.toLowerCase()).includes(settings.authContext)
        : String(acrs || "").toLowerCase() === settings.authContext;
      setStatus(satisfied
        ? `Authentication Context ${settings.authContext} was returned. This confirms the requested context claim, but not that a Conditional Access policy was enforced. Verify Entra sign-in logs or a protected API.`
        : `Token returned, but ${settings.authContext} was not visible as an exact value in the access-token acrs claim. Check CA/authentication-context configuration.`);
    } catch (e) {
      setStatus(`Step-up failed: ${e.errorCode || e.message}`);
    }
  }

  async function signOut() {
    try {
      recordConsole("Logout request", {
        method: "GET",
        endpoint: `https://login.microsoftonline.com/${settings.tenantId}/oauth2/v2.0/logout`,
        note: "MSAL opens the logout endpoint in a popup."
      });
      await instance.logoutPopup({ account });
      setClaims(null);
      setStatus("Signed out.");
    } catch (e) {
      setStatus(`Sign-out failed: ${e.errorCode || e.message}`);
    }
  }

  function clearConsole() {
    setConsoleEntries([]);
  }

  return <main>
    <header>
      <div><span className="eyebrow">MICROSOFT ENTRA POC</span><h1>Conditional Access<br/>Step-up Authentication</h1></div>
      <div className="header-actions">
        <div className={`pill ${authenticated ? "ok" : ""}`}>{authenticated ? "● Signed in" : "○ Not signed in"}</div>
        {authenticated && <button className="sign-out" onClick={signOut}>Sign out</button>}
      </div>
    </header>

    <section className="flow">
      <article>
        <span className="number">01</span>
        <h2>Standard authentication</h2>
        <p>Authenticate normally with Microsoft Entra ID. No Authentication Context is explicitly requested.</p>
        <button onClick={signIn}>Sign in with Entra</button>
        {account && <small>{account.username}</small>}
      </article>
      <div className="arrow">→</div>
      <article>
        <span className="number">02</span>
        <h2>Protected operation</h2>
        <p>Request a new token with Authentication Context <b>{settings.authContext}</b>. Conditional Access evaluates the policy mapped to it.</p>
        <button className="step" disabled={!authenticated} onClick={stepUp}>Trigger {settings.authContext.toUpperCase()} step-up</button>
      </article>
    </section>

    <section className="status"><b>Test status</b><span>{status}</span></section>

    <section className="details">
      <div><h3>Runtime configuration</h3><pre>{JSON.stringify(configView, null, 2)}</pre></div>
      <div><h3>Latest token claims</h3><pre>{claims ? JSON.stringify(claims, null, 2) : "No token captured yet."}</pre></div>
    </section>

    <section className="entra-console">
      <div className="console-heading"><div><h3>Entra request console</h3><p>Request plans, MSAL events, and observable Entra URLs. Tokens are never logged.</p></div><button className="clear-console" onClick={clearConsole}>Clear</button></div>
      <pre>{consoleEntries.length ? consoleEntries.map((entry) => `[${entry.time}] ${entry.type}\n${JSON.stringify(entry.details, null, 2)}`).join("\n\n") : "No Entra calls recorded yet. Start a sign-in or step-up flow."}</pre>
    </section>
  </main>;
}
