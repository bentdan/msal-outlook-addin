import { AjaxClient, AsyncInterceptorAuthClient } from '@chr/common-web-ui-ajax-client';
import { initializeIcons } from '@fluentui/react/lib/Icons';
import * as React from 'react';
import { createRoot } from 'react-dom/client';
import { Provider } from 'react-redux';
import { createStore } from './createStore';
import { officeInitialized } from './areas/office';
import Main from './Main';

initializeIcons();

const config: AppConfig = await (await fetch('/app-config.json')).json();
console.log('Configuration loaded:', config);

let accessToken: string = '';

const getAccessTokenOfficeApi = (): Promise<string> => {
  console.log('Using office api for auth...');
  return Office.auth.getAccessToken({
    forMSGraphAccess: false,
    allowSignInPrompt: true,
    allowConsentPrompt: true,
  })
    .then((token) => {
      console.log('Token retrieved:', token);
      accessToken = token;
      return token;
    })
    .catch((error) => {
      console.error('GetAccessToken from Office failed.', error);
      throw new Error('Token retrieval failed.');
    });
};

const getAccessTokenDialog = (): Promise<string> => {
  console.log('Using login dialog for auth...');
  
  // Open a dialog for manual login as a fallback
  const dialogLoginUrl = `${location.protocol}//${location.hostname}${location.port ? ':' + location.port : ''}/login.html`;
  
  return new Promise((resolve, reject) => {
    Office.context.ui.displayDialogAsync(dialogLoginUrl, { height: 40, width: 30 }, (result) => {
      if (result.status === Office.AsyncResultStatus.Failed) {
        reject(`Failed to open dialog: ${result.error.message}`);
      } else {
        const loginDialog = result.value;
        
        loginDialog.addEventHandler(Office.EventType.DialogMessageReceived, (args: any) => {
          if ('error' in args) {
            reject(`Error in dialog: ${args.error}`);
          } else {
            const message = JSON.parse(args.message);
            if (message.status === 'failure') {
              reject(`Failure to get access token: ${JSON.stringify(message.error)}`);
            } else {
              accessToken = message.token;  // Cache the new token
              loginDialog.close();
              resolve(accessToken);
            }
          }
        });
      }
    });
  });
};

const authClient : AsyncInterceptorAuthClient = {
  getAccessToken() {
    return getAccessTokenOfficeApi();
    //return getAccessTokenOfficeApi().catch(() => getAccessTokenDialog());
  }
}

const ajaxClient = AjaxClient.createDefault(authClient);
const store = createStore(config, authClient, ajaxClient);

Office.initialize = () => {
  console.log(`User: ${Office.context.mailbox.userProfile.displayName}`);
  console.log(`Email: ${Office.context.mailbox.userProfile.emailAddress}`);

  authClient.getAccessToken().then(token => {
    console.log('Access token:', token);
    store.dispatch(officeInitialized());
  }).catch(error => {    
    console.error('Error getting access token:', error);
  });
};

createRoot(document.getElementById('container')!).render(
  <Provider store={store}>
    <Main title={document.title} />
  </Provider>
);
