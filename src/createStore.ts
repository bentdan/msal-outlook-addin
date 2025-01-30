import { configureStore, Action, Middleware, combineReducers } from '@reduxjs/toolkit';
import { createLogger } from 'redux-logger';
import { createEpicMiddleware, combineEpics } from 'redux-observable';
import { LoggingMiddleware } from '@chr/common-javascript-logging-redux-middleware';
import { officeReducer, outlookMessageRetrievalEpic } from './areas/office';
import { AjaxClient, AsyncInterceptorAuthClient } from '@chr/common-web-ui-ajax-client';
import { AppConfig, AppDependencies, AppState } from './react-app-env';

export const createStore = (config: AppConfig, authClient: AsyncInterceptorAuthClient, ajaxClient: AjaxClient) => {
  const epicMiddleware = createEpicMiddleware<Action, Action, AppState, AppDependencies>({
    dependencies: {
      ajaxClient,
      authClient,
      ...config
    }
  });

  const rootEpic = combineEpics<Action, Action, AppState, AppDependencies>(
    outlookMessageRetrievalEpic,
  );

  const rootReducer = combineReducers({
    officeReducer,
  });

  // Note: Logging is off test environments due to the amount of logging sometimes being obtrusive. Turn it on if its helpful.
  const logRedux = process.env.NODE_ENV !== 'production' && process.env.NODE_ENV !== 'test';
  const middlewares: Middleware<{}, any, any>[] = logRedux
    ? [epicMiddleware as Middleware<{}, any, any>, LoggingMiddleware as Middleware<{}, any, any>, createLogger() as Middleware<{}, any, any>]
    : [epicMiddleware as Middleware<{}, any, any>, LoggingMiddleware as Middleware<{}, any, any>];

  const store = configureStore({
    reducer: rootReducer,
    middleware: (getDefaultMiddleware) => getDefaultMiddleware().concat(middlewares),
  });

  epicMiddleware.run(rootEpic);

  return store;
};
