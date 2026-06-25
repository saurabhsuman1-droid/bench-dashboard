import { StrictMode } from 'react'
import { createRoot } from 'react-dom/client'
import { MsalProvider } from '@azure/msal-react'
import { PublicClientApplication } from '@azure/msal-browser'
import './index.css'
import App from './App.jsx'

const pca = new PublicClientApplication({
  auth: {
    clientId: import.meta.env.VITE_AAD_CLIENT_ID || "00000000-0000-0000-0000-000000000000",
    authority: `https://login.microsoftonline.com/${import.meta.env.VITE_AAD_TENANT_ID || "common"}`,
    redirectUri: window.location.origin + window.location.pathname,
  },
  cache: { cacheLocation: "sessionStorage" },
});

createRoot(document.getElementById('root')).render(
  <StrictMode>
    <MsalProvider instance={pca}>
      <App />
    </MsalProvider>
  </StrictMode>,
)
