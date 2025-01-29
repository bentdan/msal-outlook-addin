import { ofType } from 'redux-observable';
import { mergeMap, catchError } from 'rxjs/operators';
import { of, Observable } from 'rxjs';
import { 
    START_MESSAGE_AUDIT_RETRIEVAL, 
    completeMessageAuditRetrieval, 
    failMessageAuditRetrieval, 
} from './routerActions';
import { Action } from 'redux';

export const routerEpics: AppEpic = (action$, state$, { ajaxClient, routerApiUrl }) => {
    return action$.pipe(
        ofType(START_MESSAGE_AUDIT_RETRIEVAL),
        mergeMap((action: any): Observable<Action<any>> => {
            const { type, payload } = action;
            switch (type) {
                case START_MESSAGE_AUDIT_RETRIEVAL:
                    return getAudit(payload.messageId, routerApiUrl);
                default:
                    return of();
            }
        })
    );

    function getAudit(messageId: string, routerApiUrl: string): Observable<Action<any>> {
        return ajaxClient.get({
            url: `${routerApiUrl}/Audit/EmailEvents/${messageId}/Status`
        })
        .pipe(
            mergeMap((response: any) => {
                console.log(response)
                const data = response.response;
                if (response.status === 200 || response.status === 404) {
                    return of(completeMessageAuditRetrieval(messageId, data));
                } else {
                    return of(failMessageAuditRetrieval(messageId, JSON.stringify(data)));
                }
            }),
            catchError((error: any) => {
                let errorMessage = 'Unknown error';
                if (error instanceof Error) {
                    errorMessage = error.message || 'Error with no message';
                } else if (typeof error === 'string') {
                    errorMessage = error;
                } else if (error instanceof Object) {
                    errorMessage = JSON.stringify(error);
                }
                return of(failMessageAuditRetrieval(messageId, errorMessage));
            })
        );
    }
};
