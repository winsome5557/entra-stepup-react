# Microsoft Entra Step-up Authentication POC

A deliberately small React/Vite SPA for testing Microsoft Entra Conditional Access Authentication Context.

## Flow

1. **Sign in with Entra** performs a normal interactive MSAL sign-in.
2. **Trigger C3 step-up** requests a new access token with this claims request:

```json
{
  "access_token": {
    "acrs": {
      "essential": true,
      "value": "c3"
    }
  }
}
```

Entra evaluates the Conditional Access policy associated with Authentication Context `c3`. The UI then displays the returned access-token claims so you can inspect `acrs`.

## 1. Entra app registration

Create an **App registration** in Microsoft Entra ID.

- Supported account type: normally **Accounts in this organizational directory only** for an enterprise POC.
- Authentication > Add platform > **Single-page application**
- Redirect URI: `http://localhost:5173`
- Do **not** create a client secret. A browser SPA is a public client.
- Add delegated API permission `Microsoft Graph / User.Read` if you retain the default scope used by this sample.

Copy:
- Directory (tenant) ID
- Application (client) ID

## 2. Authentication Context / Conditional Access

In Entra Conditional Access, create or enable an Authentication Context and use **c3** as its ID/reference.

Create a Conditional Access policy whose target is that Authentication Context (`c3`) and configure the grant control you actually want to test, for example an Authentication Strength or MFA requirement.

Use a test user/group first and exclude emergency/break-glass accounts according to your organisation's CA governance.

Important: `c3` itself does not mean MFA. It is simply an Authentication Context identifier. The CA policy associated with `c3` determines the controls Entra requires.

## 3. Configure the app

Copy `.env.example` to `.env` and replace the values:

```text
VITE_ENTRA_CLIENT_ID=<application-client-id>
VITE_ENTRA_TENANT_ID=<directory-tenant-id>
VITE_ENTRA_REDIRECT_URI=http://localhost:5173
VITE_ENTRA_SCOPES=openid profile User.Read
VITE_ENTRA_AUTH_CONTEXT=c3
```

You can change `VITE_ENTRA_AUTH_CONTEXT` to another configured value such as `c1`.

## 4. Run

```bash
npm install
npm run dev
```

Open `http://localhost:5173`.

## Expected test

- Click **Sign in with Entra** and complete ordinary authentication.
- Click **Trigger C3 step-up**.
- The second request sends the `acrs=c3` claims request.
- Entra evaluates the CA policy bound to `c3`.
- If the existing session already satisfies that policy, Entra may not require an additional authentication method. Step-up means satisfying the requested context, not necessarily showing MFA every time.
- Inspect **Latest token claims** for the resulting `acrs` value.

## Production architecture note

This front-end-only implementation is ideal for demonstrating the protocol request. In a production application, the stronger pattern is:

`React SPA -> protected API -> API checks acrs -> API returns claims challenge -> SPA passes challenge to MSAL -> Entra CA -> new token -> retry API`

That prevents a user interface button alone from deciding which server-side operation requires elevated authentication.
