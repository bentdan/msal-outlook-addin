import {
  START_MESSAGE_AUDIT_RETRIEVAL,
  COMPLETE_MESSAGE_AUDIT_RETRIEVAL,
  FAIL_MESSAGE_AUDIT_RETRIEVAL,
} from './routerActions';

export default function reduce(state: any = {}, action: any) {
  switch (action.type) {

    case START_MESSAGE_AUDIT_RETRIEVAL:
      return {
        ...state,
        audit: undefined,
        errorMessage: '',
      };

    case COMPLETE_MESSAGE_AUDIT_RETRIEVAL:
      return {
        ...state,
        audit: action.audit,
        errorMessage: '',
      };

    case FAIL_MESSAGE_AUDIT_RETRIEVAL:
      return {
        ...state,
        audit: undefined,
        errorMessage: action.error,
      };

    default:
      return state;
  }
}
