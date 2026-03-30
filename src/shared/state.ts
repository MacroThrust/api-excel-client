/**
 * Shared runtime state accessible from custom functions, taskpane, and commands.
 *
 * Because the add-in uses a shared runtime (lifetime="long"), all entry points
 * (taskpane.ts, functions.ts, commands.ts) share the same JavaScript context.
 * Module-level state here is the canonical store for auth tokens and user info.
 */

export interface AuthState {
  isAuthenticated: boolean;
  msAccessToken: string | null;
  authentikAccessToken: string | null;
  authentikRefreshToken: string | null;
  tokenExpiry: number | null;
  userDisplayName: string | null;
  userEmail: string | null;
}

const state: AuthState = {
  isAuthenticated: false,
  msAccessToken: null,
  authentikAccessToken: null,
  authentikRefreshToken: null,
  tokenExpiry: null,
  userDisplayName: null,
  userEmail: null,
};

export function getAuthState(): Readonly<AuthState> {
  return state;
}

export function setAuthenticated(params: {
  msAccessToken: string;
  authentikAccessToken: string;
  authentikRefreshToken?: string;
  tokenExpiry?: number;
  userDisplayName?: string;
  userEmail?: string;
}): void {
  state.isAuthenticated = true;
  state.msAccessToken = params.msAccessToken;
  state.authentikAccessToken = params.authentikAccessToken;
  state.authentikRefreshToken = params.authentikRefreshToken ?? null;
  state.tokenExpiry = params.tokenExpiry ?? null;
  state.userDisplayName = params.userDisplayName ?? null;
  state.userEmail = params.userEmail ?? null;

  notifyListeners();
}

export function clearAuth(): void {
  state.isAuthenticated = false;
  state.msAccessToken = null;
  state.authentikAccessToken = null;
  state.authentikRefreshToken = null;
  state.tokenExpiry = null;
  state.userDisplayName = null;
  state.userEmail = null;

  notifyListeners();
}

export function isTokenExpired(): boolean {
  if (!state.tokenExpiry) return true;
  return Date.now() >= state.tokenExpiry;
}

type AuthChangeListener = (state: Readonly<AuthState>) => void;
const listeners: AuthChangeListener[] = [];

export function onAuthChange(listener: AuthChangeListener): () => void {
  listeners.push(listener);
  return () => {
    const idx = listeners.indexOf(listener);
    if (idx >= 0) listeners.splice(idx, 1);
  };
}

function notifyListeners(): void {
  const snapshot = { ...state };
  listeners.forEach((fn) => fn(snapshot));
}
