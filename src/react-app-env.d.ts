import { OfficeState } from './areas/office';
import { CommonConfig } from '@chr/common-web-ui-configuration';

declare interface AppConfig extends CommonConfig{
  routerApiUrl: string;
}

declare interface AppState {
  officeReducer: OfficeState;
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
