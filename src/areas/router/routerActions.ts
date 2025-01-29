export const START_MESSAGE_AUDIT_RETRIEVAL ='router/START_MESSAGE_AUDIT_RETRIEVAL';
export const COMPLETE_MESSAGE_AUDIT_RETRIEVAL = 'router/COMPLETE_MESSAGE_AUDIT_RETRIEVAL';
export const FAIL_MESSAGE_AUDIT_RETRIEVAL = 'router/FAIL_MESSAGE_AUDIT__RETRIEVAL';

export function requestMessageAuditRetrieval(
  messageId: string
) {
  return {
    type: START_MESSAGE_AUDIT_RETRIEVAL,
    payload : {
      messageId
    },
  }
};

export function completeMessageAuditRetrieval(messageId: string, audit: any) {
  return {
    type: COMPLETE_MESSAGE_AUDIT_RETRIEVAL,
    messageId,
    audit
  };
}

export function failMessageAuditRetrieval(messageId: string, error: any) {
  return {
    type: FAIL_MESSAGE_AUDIT_RETRIEVAL,
    messageId,
    error
  };
}

