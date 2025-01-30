import { OfficeAction, OfficeState } from './officeTypes';
import {
  SET_OFFICE_INITIALIZED,
  SET_OFFICE_THEME,
  COMPLETE_OUTLOOK_MESSAGE_RETRIEVAL,
  FAIL_OUTLOOK_MESSAGE_RETRIEVAL,
  START_OUTLOOK_MESSAGE_RETRIEVAL
} from './officeActions';

export const initialState: OfficeState = {
  host: undefined,
  isDarkMode: false,
  isOfficeInitialized: false,
  userName: undefined,
  multiSelect: false,
  noSelection: true,
  outlookMessageId: undefined,
  messageAgeInMinutes: undefined,
  outlookErrorMessage: undefined,
};

// this is a hack; officeTheme.isDarkTheme does not work in Outlook
// https://learn.microsoft.com/en-us/javascript/api/office/office.officetheme?view=common-js-preview
function isDarkTheme() {
  const officeTheme = Office.context.officeTheme;
  if (!officeTheme) {
    return window.matchMedia ? window.matchMedia('(prefers-color-scheme: dark)').matches : false;
  }
  
  return officeTheme.isDarkTheme.toString() === '#000001' || officeTheme.isDarkTheme.toString() === 'true'; // officeTheme.bodyBackgroundColor !== '#ffffff'
}

export default function reduce(state: OfficeState, action: OfficeAction) {
  if (!state) {
    state = initialState;
  }

  switch (action.type) {
  case SET_OFFICE_INITIALIZED:
  {
    return {
      ...state,
      host: Office.context.mailbox.diagnostics.hostName,
      isDarkMode: isDarkTheme(),
      isOfficeInitialized: true,
      userName: action.userName,
    };
  }
  case SET_OFFICE_THEME:
  {
    return {
      ...state,
      isDarkMode: isDarkTheme()
    };
  }
  case START_OUTLOOK_MESSAGE_RETRIEVAL:
  {
    return {
      ...state,
      outlookMessageId: undefined,
      messageAgeInMinutes: undefined,
      multiSelect: false,
      noSelection: true,
      outlookErrorMessage: undefined,
    };
  }
  case COMPLETE_OUTLOOK_MESSAGE_RETRIEVAL:
  {
    return {
      ...state,
      outlookMessageId: action.messageId,
      messageAgeInMinutes: action.messageAgeInMinutes,
      multiSelect: action.multiSelect,
      noSelection: action.noSelection,
      outlookErrorMessage: undefined,
    };
  }
  case FAIL_OUTLOOK_MESSAGE_RETRIEVAL:
  {
    return {
      ...state,
      outlookMessageId: undefined,
      messageAgeInMinutes: undefined,
      multiSelect: false,
      noSelection: true,
      outlookErrorMessage: action.error,
    };
  }
  default:
  {
    return state;
  }}
}
