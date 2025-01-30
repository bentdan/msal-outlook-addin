import { AsyncInterceptorAuthClient } from "@chr/common-web-ui-ajax-client";

export interface AuthClient extends AsyncInterceptorAuthClient {
    getIdentityToken(): Promise<string>;
}