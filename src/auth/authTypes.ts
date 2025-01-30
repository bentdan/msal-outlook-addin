export interface TokenResult {
  accessToken?: string,
  identityToken?: string,
  refreshToken?: string;
}

export interface LoginResponse {
  status: string;
  accessToken: string | undefined;
  identityToken: string | undefined;
  refreshToken: string | undefined;
}

export interface BaseDecodedJwt {
  aud: string;
  iss: string;
  iat: number;
  exp: number;
  sub: string;
  name: string;
  preferred_username: string;
  ver: number | string; // Adjusted to support both number and string versions
}

export interface AzureDecodedJwt extends BaseDecodedJwt {
  nbf: number;
  aio: string;
  azp: string;
  azpacr: string;
  oid: string;
  rh: string;
  scp: string;
  sid: string;
  tid: string;
  uti: string;
}

export interface ChrDecodedJwt extends BaseDecodedJwt {
  email: string;
  jti: string;
  amr: string[];
  idp: string;
  nonce: string;
  auth_time: number;
  at_hash: string;
  samAccountName: string;
}

export type DecodedJwt = AzureDecodedJwt | ChrDecodedJwt;
