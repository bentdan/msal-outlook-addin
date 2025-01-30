import type { Action } from 'redux';

export const SET_OFFICE_INITIALIZED =
  'office/SET_OFFICE_INITIALIZED';
export const SET_OFFICE_THEME =
  'office/SET_OFFICE_THEME';
export const START_OUTLOOK_MESSAGE_RETRIEVAL =
  'office/START_OUTLOOK_MESSAGE_RETRIEVAL';
export const COMPLETE_OUTLOOK_MESSAGE_RETRIEVAL =
  'office/COMPLETE_OUTLOOK_MESSAGE_RETRIEVAL';
export const FAIL_OUTLOOK_MESSAGE_RETRIEVAL =
  'office/FAIL_OUTLOOK_MESSAGE_RETRIEVAL';

export type SetOfficeInitialized =
  Action<typeof SET_OFFICE_INITIALIZED> & {
    userName: string;
  };

export type SetOfficeTheme = Action<typeof SET_OFFICE_THEME>;

export type StartOutlookMessageRetrieval =
  Action<typeof START_OUTLOOK_MESSAGE_RETRIEVAL>;

export type CompleteOutlookMessageRetrieval =
  Action<typeof COMPLETE_OUTLOOK_MESSAGE_RETRIEVAL> & {
    noSelection: boolean;
    multiSelect: boolean;
    messageId: string | undefined;
    messageAgeInMinutes: number | undefined;
  };

export type FailOutlookMessageRetrieval =
  Action<typeof FAIL_OUTLOOK_MESSAGE_RETRIEVAL> & {
    error: string;
  };

export function setOfficeInitialized(
  userName: string,
): SetOfficeInitialized {
  return {
    type: SET_OFFICE_INITIALIZED,
    userName,
  };
}

export function setOfficeTheme(): SetOfficeTheme {
  return {
    type: SET_OFFICE_THEME,
  };
}

export function startOutlookMessageRetrieval(): StartOutlookMessageRetrieval {
  return {
    type: START_OUTLOOK_MESSAGE_RETRIEVAL,
  };
}

export function completeOutlookMessageRetrieval(
  noSelection: boolean,
  multiSelect: boolean,
  messageId: string | undefined,
  messageAgeInMinutes: number | undefined,
): CompleteOutlookMessageRetrieval {
  return {
    type: COMPLETE_OUTLOOK_MESSAGE_RETRIEVAL,
    noSelection,
    multiSelect,
    messageId,
    messageAgeInMinutes
  };
}

export function failOutlookMessageRetrieval(
  error: string
): FailOutlookMessageRetrieval {
  return {
    type: FAIL_OUTLOOK_MESSAGE_RETRIEVAL,
    error
  };
}
