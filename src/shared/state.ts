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
const AUTH_BROADCAST_CHANNEL = "mt_auth_state_sync";

const state: AuthState = createEmptyAuthState();

const authBroadcast =
  typeof BroadcastChannel !== "undefined"
    ? new BroadcastChannel(AUTH_BROADCAST_CHANNEL)
    : null;

if (authBroadcast) {
  authBroadcast.addEventListener("message", () => {
    hydrateFromStorage();
    notifyListeners();
  });
}

function broadcastAuthChange(): void {
  authBroadcast?.postMessage("changed");
}

function createEmptyAuthState(): AuthState {
  return {
    isAuthenticated: false,
    msAccessToken: null,
    authentikAccessToken: null,
    authentikRefreshToken: null,
    tokenExpiry: null,
    userDisplayName: null,
    userEmail: null,
  };
}

function resetInMemoryState(): void {
  Object.assign(state, createEmptyAuthState());
}

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
    if (!raw) {
      resetInMemoryState();
      return;
    }

    const parsed = JSON.parse(raw) as Partial<AuthState>;
    if (!parsed.isAuthenticated || !parsed.authentikAccessToken) {
      resetInMemoryState();
      return;
    }

    state.isAuthenticated = true;
    state.msAccessToken = parsed.msAccessToken ?? null;
    state.authentikAccessToken = parsed.authentikAccessToken;
    state.authentikRefreshToken = parsed.authentikRefreshToken ?? null;
    state.tokenExpiry = parsed.tokenExpiry ?? null;
    state.userDisplayName = parsed.userDisplayName ?? null;
    state.userEmail = parsed.userEmail ?? null;
  } catch {
    resetInMemoryState();
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
  broadcastAuthChange();
}

export function clearAuth(): void {
  resetInMemoryState();

  try {
    localStorage.removeItem(STORAGE_KEY_AUTH);
  } catch {
    // ignore
  }

  notifyListeners();
  broadcastAuthChange();
}

function getJwtExpiryMs(token: string): number | null {
  try {
    const parts = token.split(".");
    if (parts.length !== 3) return null;
    const payload = JSON.parse(atob(parts[1].replace(/-/g, "+").replace(/_/g, "/"))) as {
      exp?: number;
    };
    return payload.exp ? payload.exp * 1000 : null;
  } catch {
    return null;
  }
}

export function isTokenExpired(): boolean {
  hydrateFromStorage();
  const now = Date.now();

  if (state.tokenExpiry && now >= state.tokenExpiry - 30_000) {
    return true;
  }

  if (state.authentikAccessToken) {
    const jwtExpiry = getJwtExpiryMs(state.authentikAccessToken);
    if (jwtExpiry && now >= jwtExpiry - 30_000) {
      return true;
    }
  }

  return false;
}

export function updateUserProfile(params: {
  userDisplayName?: string | null;
  userEmail?: string | null;
}): void {
  if (params.userDisplayName !== undefined) {
    state.userDisplayName = params.userDisplayName;
  }
  if (params.userEmail !== undefined) {
    state.userEmail = params.userEmail;
  }
  persistState();
  notifyListeners();
  broadcastAuthChange();
}

export function updateAccessToken(params: {
  authentikAccessToken: string;
  authentikRefreshToken?: string | null;
  tokenExpiry?: number | null;
}): void {
  state.authentikAccessToken = params.authentikAccessToken;
  if (params.authentikRefreshToken !== undefined) {
    state.authentikRefreshToken = params.authentikRefreshToken;
  }
  if (params.tokenExpiry !== undefined) {
    state.tokenExpiry = params.tokenExpiry;
  }
  persistState();
  notifyListeners();
  broadcastAuthChange();
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
