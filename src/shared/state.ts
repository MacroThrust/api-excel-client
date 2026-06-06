/**
 * Shared runtime state accessible from custom functions, taskpane, and commands.
 *
 * Webpack emits separate bundles (taskpane.js vs functions.js), so in-memory
 * module state is not shared between entry points even under Office shared
 * runtime. Auth tokens are persisted to localStorage and re-read on every
 * getAuthState() call so all bundles see the same session.
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

const STORAGE_KEY_AUTH = "mt_auth_state";

const state: AuthState = {
  isAuthenticated: false,
  msAccessToken: null,
  authentikAccessToken: null,
  authentikRefreshToken: null,
  tokenExpiry: null,
  userDisplayName: null,
  userEmail: null,
};

function persistState(): void {
  try {
    localStorage.setItem(STORAGE_KEY_AUTH, JSON.stringify(state));
  } catch {
    // localStorage may be unavailable in some Office hosts
  }
}

function hydrateFromStorage(): void {
  try {
    const raw = localStorage.getItem(STORAGE_KEY_AUTH);
    if (!raw) return;

    const parsed = JSON.parse(raw) as Partial<AuthState>;
    if (!parsed.isAuthenticated || !parsed.authentikAccessToken) return;

    state.isAuthenticated = true;
    state.msAccessToken = parsed.msAccessToken ?? null;
    state.authentikAccessToken = parsed.authentikAccessToken;
    state.authentikRefreshToken = parsed.authentikRefreshToken ?? null;
    state.tokenExpiry = parsed.tokenExpiry ?? null;
    state.userDisplayName = parsed.userDisplayName ?? null;
    state.userEmail = parsed.userEmail ?? null;
  } catch {
    // Ignore corrupt storage
  }
}

export function getAuthState(): Readonly<AuthState> {
  hydrateFromStorage();
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

  persistState();
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

  try {
    localStorage.removeItem(STORAGE_KEY_AUTH);
  } catch {
    // ignore
  }

  notifyListeners();
}

export function isTokenExpired(): boolean {
  hydrateFromStorage();
  if (!state.tokenExpiry) return false;
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

// Hydrate once at module load so the taskpane shows the last session immediately.
hydrateFromStorage();
