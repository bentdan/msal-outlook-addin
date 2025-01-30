import React from 'react';
import { Root, createRoot } from 'react-dom/client';
import { Provider as StoreProvider } from 'react-redux';
import { ChrThemeProvider } from '@chr/eds-react';
import { initializeIcons } from '@fluentui/react/lib/Icons';
import { jwtDecode } from 'jwt-decode';
import { createStore } from './createStore';
import { AppConfig } from './react-app-env';
import { setOfficeInitialized, setOfficeTheme } from './areas/office/officeActions';
import { AuthClient } from './auth/AuthClient';
import { DecodedJwt, LoginResponse } from './auth/authTypes';
import App from './components/App';
import { getConfiguration } from './utils/configuration';
import { AjaxClient } from '@chr/common-web-ui-ajax-client';

let root: Root;

// Initialize Fluent UI icons
initializeIcons();

function handleException(error: any): never {
  const errorMessage = typeof error === 'string' ? error : JSON.stringify(error);
  throw new Error(`Error: ${errorMessage}`);
}

// Handle dialog responses and return the access token or throw an error
async function handleDialogResponse(args: any, loginDialog: any): Promise<LoginResponse> {
  if ('error' in args) {
    handleException(args.error);
  }

  const message = JSON.parse(args.message);
  if (message.status === 'failure') {
    handleException(message.error);
  }

  loginDialog.close();
  return message;
}

const getTokenOfficeApi = (): Promise<string> => {
  console.log('Using office api for auth...');
  return Office.auth.getAccessToken({
    forMSGraphAccess: false,
    allowSignInPrompt: true,
    allowConsentPrompt: true,
  })
    .then((token) => {
      console.log('Token retrieved from Office:', token);
      return token;
    })
    .catch((error) => {
      console.error('GetAccessToken from Office failed.', error);
      throw new Error('Token retrieval failed.');
    });
};

// Get access token by displaying a login dialog
async function startLoginDialog(): Promise<LoginResponse> {
  console.log('Using okta for auth...');
  const dialogLoginUrl = `${location.origin}/login.html`;

  return new Promise((resolve, reject) => {
    if (Office.context?.ui?.displayDialogAsync) {
      Office.context.ui.displayDialogAsync(dialogLoginUrl, { height: 70, width: 35, promptBeforeOpen: false }, (result) => {
        if (result.status === Office.AsyncResultStatus.Failed) {
          return reject(`Failed to open dialog: ${result.error.message}`);
        }

        const loginDialog = result.value;
        loginDialog.addEventHandler(Office.EventType.DialogMessageReceived, async (args) => {
          try {
            resolve(await handleDialogResponse(args, loginDialog));
          } catch (error) {
            reject(error);
          }
        });
      });
    } else {
      reject('Office.context.ui.displayDialogAsync is not available.');
    }
  });
}

// Initialize the Auth client with token management
function createAuthClient(): AuthClient {
  let accessToken: string | undefined;
  let identityToken: string | undefined;

  const TOKEN_EXPIRY_BUFFER_SECONDS = 10;

  function isTokenExpiring(token: DecodedJwt, bufferInSeconds: number = TOKEN_EXPIRY_BUFFER_SECONDS): boolean {
    return token.exp <= Date.now() / 1000 + bufferInSeconds;
  }

  async function refreshTokens(): Promise<void> {
    try {
      accessToken = await getTokenOfficeApi();
      identityToken = accessToken; // this token is both an accessToken and identityToken
    } catch {
      const tokenResponse = await startLoginDialog();
      accessToken = tokenResponse.accessToken;
      identityToken = tokenResponse.identityToken;
    }
  }

  async function ensureAccessToken(): Promise<string> {
    if (accessToken) {
      const decodedToken: DecodedJwt = jwtDecode(accessToken);
      if (!isTokenExpiring(decodedToken)) {
        return accessToken;
      }
    }
    await refreshTokens();
    if (!accessToken) {
      throw new Error('Failed to retrieve a valid access token.');
    }
    return accessToken;
  }

  async function ensureIdentityToken(): Promise<string> {
    if (!identityToken) {
      await refreshTokens();
      if (!identityToken) throw new Error('Failed to retrieve a valid identity token.');
    }
    return identityToken;
  }

  return {
    getAccessToken: ensureAccessToken,
    getIdentityToken: ensureIdentityToken,
  };
}

// Create and initialize the store
async function initializeStore(authClient: AuthClient) {
  const config: AppConfig = await getConfiguration();
  const ajaxClient = AjaxClient.createDefault(authClient);
  return createStore(config, authClient, ajaxClient);
}

// Initialize Office context and get access token
async function initializeOffice(store: any, authClient: AuthClient): Promise<void> {
  try {
    let userName;
    await authClient.getAccessToken();
    const idToken = await authClient.getIdentityToken();
    if (idToken) {
      const decodedToken: DecodedJwt = jwtDecode(idToken);
      if ("preferred_username" in decodedToken && decodedToken.preferred_username) {
        // preferred_username is users email.
        userName = decodedToken.preferred_username
      } else if ("samAccountName" in decodedToken && decodedToken.samAccountName) {
        // samAccountName is users 7 letter.
        userName = decodedToken.samAccountName;
      } else {
        userName = "unknown";
      }
      console.log('User name:', userName);
    } else {
      // Fallback to Office context if no identity token is available
      // This will return the name of the mailbox which will often be a shared mailbox and not the user's name
      userName = Office.context.mailbox.userProfile.displayName;
    }
    store.dispatch(setOfficeInitialized(userName));
  } catch (error) {
    console.error('Error getting access token', error);
  }
}

// Render the main React component
function renderMainComponent(store: any) {
  const { officeReducer } = store.getState();
  if (!root) {
    root = createRoot(document.getElementById('container')!);
  } 

  root.render(
    <StoreProvider store={store}>
      <ChrThemeProvider colorPreference={officeReducer.isDarkMode ? 'dark' : 'light'}>
        <App />
      </ChrThemeProvider>
    </StoreProvider>
  );
}

async function renderOfficeReady() {
  if (Office.context.requirements.isSetSupported("IdentityAPI", "1.3")) {
    console.log("Supported");
  }
  const authClient = createAuthClient();
  const store = await initializeStore(authClient);
  await initializeOffice(store, authClient);
  Office.context?.mailbox?.addHandlerAsync(Office.EventType.OfficeThemeChanged, () => {
    store.dispatch(setOfficeTheme());
  });
  renderMainComponent(store);
}

async function renderBeforeOfficeReady() {
  const authClient = createAuthClient();
  const store = await initializeStore(authClient);
  renderMainComponent(store);
}

// Initialize the application when Office is ready
Office.onReady(renderOfficeReady);

// Initial render before Office is ready
renderBeforeOfficeReady();
