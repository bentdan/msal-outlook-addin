import { createStore as createReduxStore, applyMiddleware, Action, Middleware, combineReducers } from 'redux';
import { createLogger } from 'redux-logger';
import { createEpicMiddleware, combineEpics } from 'redux-observable';
import { LoggingMiddleware } from '@chr/common-javascript-logging-redux-middleware';
import { officeReducers } from './areas/office';
import { routerEpics } from './areas/router';
import { AjaxClient, AsyncInterceptorAuthClient } from '@chr/common-web-ui-ajax-client';

export const createStore = (config: AppConfig, authClient: AsyncInterceptorAuthClient, ajaxClient: AjaxClient) => {
  const epicMiddleware = createEpicMiddleware<Action, Action, AppState, AppDependencies>({
    dependencies: {
      ajaxClient,
      authClient,
      ...config
    }
  });

  const rootEpic = combineEpics<Action, Action, AppState, AppDependencies>(
    routerEpics
  );

  const rootReducer = combineReducers({
    officeReducers,
  });

  const middlewares: Middleware[] = process.env.NODE_ENV === 'production'
    ? [epicMiddleware, LoggingMiddleware]
    : [epicMiddleware, LoggingMiddleware, createLogger() as Middleware];

  const store = createReduxStore(
    rootReducer,
    applyMiddleware(...middlewares)
  );

  epicMiddleware.run(rootEpic);

  return store;
};
