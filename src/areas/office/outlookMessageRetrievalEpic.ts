import { ofType } from 'redux-observable';
import { Observable, of } from 'rxjs';
import { catchError, map, mergeMap } from 'rxjs/operators';
import {
  completeOutlookMessageRetrieval,
  failOutlookMessageRetrieval,
  START_OUTLOOK_MESSAGE_RETRIEVAL
} from './officeActions';
import { differenceInMinutes } from 'date-fns';
import { AppEpic } from '../../react-app-env';

interface MessageRetrievalData {
  multiSelect: boolean;
  noSelection: boolean;
  messageId: string | undefined;
  messageAgeInMinutes: number | undefined;
}

// Helper to calculate the message age in minutes
const calculateMessageAgeInMinutes = (messageCreatedUTC: Date | null): number => {
  if (!messageCreatedUTC) return 0;
  const nowUTC = new Date(Date.now());
  return differenceInMinutes(nowUTC, messageCreatedUTC);
};

export const outlookMessageRetrievalEpic: AppEpic = (action$) =>
  action$.pipe(
    ofType(START_OUTLOOK_MESSAGE_RETRIEVAL),
    mergeMap(() => new Observable<MessageRetrievalData>(subscriber => {
      try {
        if (!Office.context.mailbox) {
          subscriber.error(new Error('Office mailbox not available; must be loaded in Outlook'));
        } else {
          const item = Office.context.mailbox.item;
          if (item) {
            item.getAllInternetHeadersAsync(result => {
              if (result.status === Office.AsyncResultStatus.Succeeded) {
                let messageId = result.value.match(/^\s*Message-ID:\s*(\S+)/im)?.[1];
                if (!messageId) {
                  // fall back to using messageId from item
                  messageId = item.internetMessageId
                  if (!messageId) {
                    subscriber.error(new Error('Message ID could not be found for this email'));
                  }
                }
                const messageCreatedUTC = item.dateTimeCreated;
                const messageAgeInMinutes = calculateMessageAgeInMinutes(messageCreatedUTC);
                subscriber.next({ noSelection: false, multiSelect: false, messageId, messageAgeInMinutes });
                subscriber.complete();
              } else {
                subscriber.error(result.error);
              }
            });
          } else {
            subscriber.next({ noSelection: true, multiSelect: false, messageId: undefined, messageAgeInMinutes: undefined });
          }
        }
      } catch (error) {
        subscriber.error(`Failed to get Message ID. Error: ${error.message || 'Unknown error'}`);
      }
    }).pipe(
      map(({ noSelection, multiSelect, messageId, messageAgeInMinutes }) => completeOutlookMessageRetrieval(noSelection, multiSelect, messageId, messageAgeInMinutes)),
      catchError((error: Error) => {
        return of(failOutlookMessageRetrieval(error.message || 'Unknown error'));
      })
    )
    )
  );
