declare interface AppConfig {
  routerApiUrl: string;
}

declare interface RouterState {
  currentState?: string;
  audit?: any; // audit response on get
  errorMessage?: string;
}

declare interface OfficeState {
  isOfficeInitialized: boolean;
}

declare interface AppState {
  routerReducers: RouterState;
  officeReducers: OfficeState;
}

declare type AppEpic = import('redux-observable').Epic<
  import('redux').Action,
  import('redux').Action,
  AppState,
  AppDependencies
>

declare interface AppDependencies extends AppConfig {
  ajaxClient: import('@chr/common-web-ui-ajax-client').AjaxClient
  authClient: import('@chr/common-web-ui-ajax-client').AsyncInterceptorAuthClient
}
