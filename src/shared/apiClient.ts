/**
 * HTTP client for the target API.
 *
 * Thin Excel adapter over @macrothrust/api-client — attaches Authentik bearer
 * token from shared state and targets the configured API base URL.
 */

import {
  apiRequest as sharedApiRequest,
  ApiError,
  type ApiRequestOptions,
} from "@macrothrust/api-client";

import { refreshAccessToken } from "../auth/authConfig";
import { getConfig } from "./config";
import { getAuthState, isTokenExpired } from "./state";

export { ApiError, type ApiRequestOptions };

const excelTokenProvider = {
  async getAccessToken(): Promise<string | null> {
    const auth = getAuthState();
    if (!auth.isAuthenticated || !auth.authentikAccessToken) {
      return null;
    }
    return auth.authentikAccessToken;
  },

  async refreshIfNeeded(): Promise<boolean> {
    if (!isTokenExpired()) return true;
    return refreshAccessToken();
  },
};

function getClientConfig() {
  return {
    baseUrl: getConfig().apiBaseUrl,
    tokenProvider: excelTokenProvider,
    notAuthenticatedMessage: "Not authenticated. Please sign in first.",
    unauthorizedMessage: "Unauthorized — sign in again from the MT task pane.",
  };
}

export async function apiRequest<T = unknown>(options: ApiRequestOptions): Promise<T> {
  return sharedApiRequest<T>(getClientConfig(), options);
}
