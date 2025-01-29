import {
    OFFICE_INITIALIZED
} from './officeActions';

export default function reduce(state: any, action: any) {
  if (!state) {
    state = {};
  }

  switch (action.type) {

    case OFFICE_INITIALIZED:
      return {
        ...state,
        isOfficeInitialized: true,
      };

    default:
      return state;
  }
}
