import { CompleteOutlookMessageRetrieval, FailOutlookMessageRetrieval, SetOfficeInitialized, SetOfficeTheme, StartOutlookMessageRetrieval } from "./officeActions";

export interface OfficeState {
  // https://learn.microsoft.com/en-us/javascript/api/outlook/office.mailbox?view=outlook-js-preview&preserve-view=true#outlook-office-mailbox-diagnostics-member
  host: 'Outlook' | 'newOutlookWindows' | 'OutlookWebApp' | 'OutlookIOS' | 'OutlookAndroid' | undefined;
  isDarkMode: boolean | undefined;
  isOfficeInitialized: boolean;
  userName: string | undefined;
  multiSelect: boolean;
  noSelection: boolean;
  outlookMessageId: string | undefined;
  messageAgeInMinutes: number | undefined;
  outlookErrorMessage?: string;
}

export type OfficeAction =
  | StartOutlookMessageRetrieval
  | SetOfficeInitialized
  | SetOfficeTheme
  | CompleteOutlookMessageRetrieval
  | FailOutlookMessageRetrieval;