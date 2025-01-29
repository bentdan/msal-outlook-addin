import { PublicClientApplication } from "@azure/msal-browser";

Office.onReady(async () => {
  const clientId = "6e453c5d-acae-46ed-859d-2e466f91f4c4";
  const pca = new PublicClientApplication({
    auth: {
      clientId: clientId,
      authority: 'https://login.microsoftonline.com/d441ad83-6235-46d6-ab1a-89744a91b1d8',
      redirectUri: `${window.location.origin}/login.html`, // Must be registered as "spa" type.
      navigateToLoginRequestUrl: false,
    },
    cache: {
      cacheLocation: 'localStorage',
      storeAuthStateInCookie: false,
    },
  });
  await pca.initialize();

  try {
    const response = await pca.handleRedirectPromise();
    if (response) {
      console.log("Login successful", response);
      Office.context.ui.messageParent(JSON.stringify({ status: 'success', token: response.accessToken, userName: response.account.username}));
    } else {
      console.log("No response, triggering login.");
      // A problem occurred, so invoke login.
      await pca.loginRedirect({
        scopes: ['api://19b0b9cd-4d8f-4d36-b139-1a3308e8fa69/EmailRouter.Read'],
        //extraScopesToConsent: ["user.read"],
      });
    }
  } catch (error) {
    const errorData = {
      errorMessage: error.errorCode,
      message: error.errorMessage,
      errorCode: error.stack
    };
    // Display error message in the dialog UI so its easy for users to see and report.
    const errorMessageDiv = document.getElementById('errorMessage');
    if (errorMessageDiv) {
      errorMessageDiv.innerText = `Error: ${errorData.errorMessage}\nMessage: ${errorData.message}\nStack: ${errorData.errorCode}`;
    }
    // Send error data back to the parent
    Office.context.ui.messageParent(JSON.stringify({ status: 'failure', error: errorData }));
  }
});