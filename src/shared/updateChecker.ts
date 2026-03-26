/**
 * Automatic update notification system.
 *
 * On a periodic interval, fetches `version.json` from the add-in hosting
 * origin (the same server that serves the HTML/JS). If the server's version
 * is newer than the version embedded in the running bundle, the update state
 * is set and listeners (the taskpane UI) are notified so they can display a
 * banner with instructions.
 *
 * Because Office caches add-in web resources aggressively, a user may be
 * running stale JavaScript even after a new version has been deployed. This
 * check detects that situation.
 */

import { ADDIN_VERSION } from "./version";
import { getConfig } from "./config";

export interface UpdateInfo {
  available: boolean;
  currentVersion: string;
  latestVersion: string | null;
  releaseNotesUrl: string | null;
  checkedAt: number;
}

interface VersionManifest {
  version: string;
  buildTimestamp: string;
  minimumSupported: string;
  releaseNotesUrl: string;
}

const CHECK_INTERVAL_MS = 30 * 60 * 1000; // 30 minutes
const CACHE_BUST_PARAM = "_t";

let updateState: UpdateInfo = {
  available: false,
  currentVersion: ADDIN_VERSION,
  latestVersion: null,
  releaseNotesUrl: null,
  checkedAt: 0,
};

let intervalHandle: ReturnType<typeof setInterval> | null = null;

type UpdateListener = (info: UpdateInfo) => void;
const listeners: UpdateListener[] = [];

export function onUpdateAvailable(listener: UpdateListener): () => void {
  listeners.push(listener);
  if (updateState.available) {
    listener(updateState);
  }
  return () => {
    const idx = listeners.indexOf(listener);
    if (idx >= 0) listeners.splice(idx, 1);
  };
}

export function getUpdateInfo(): Readonly<UpdateInfo> {
  return updateState;
}

function notifyListeners(): void {
  const snapshot = { ...updateState };
  listeners.forEach((fn) => fn(snapshot));
}

/**
 * Semver comparison: returns 1 if a > b, -1 if a < b, 0 if equal.
 * Handles standard major.minor.patch format.
 */
function compareSemver(a: string, b: string): number {
  const pa = a.split(".").map(Number);
  const pb = b.split(".").map(Number);
  const len = Math.max(pa.length, pb.length);
  for (let i = 0; i < len; i++) {
    const na = pa[i] ?? 0;
    const nb = pb[i] ?? 0;
    if (na > nb) return 1;
    if (na < nb) return -1;
  }
  return 0;
}

async function fetchVersionManifest(): Promise<VersionManifest | null> {
  const config = getConfig();
  const url = `${config.addinHost}/version.json?${CACHE_BUST_PARAM}=${Date.now()}`;

  try {
    const response = await fetch(url, {
      cache: "no-store",
      headers: { Accept: "application/json" },
    });
    if (!response.ok) return null;
    return await response.json();
  } catch {
    return null;
  }
}

export async function checkForUpdate(): Promise<UpdateInfo> {
  const manifest = await fetchVersionManifest();

  if (!manifest?.version) {
    updateState = {
      ...updateState,
      checkedAt: Date.now(),
    };
    return updateState;
  }

  const isNewer = compareSemver(manifest.version, ADDIN_VERSION) > 0;

  updateState = {
    available: isNewer,
    currentVersion: ADDIN_VERSION,
    latestVersion: manifest.version,
    releaseNotesUrl: manifest.releaseNotesUrl || null,
    checkedAt: Date.now(),
  };

  if (isNewer) {
    notifyListeners();
  }

  return updateState;
}

export function startPeriodicCheck(): void {
  if (intervalHandle) return;

  checkForUpdate();

  intervalHandle = setInterval(checkForUpdate, CHECK_INTERVAL_MS);
}

export function stopPeriodicCheck(): void {
  if (intervalHandle) {
    clearInterval(intervalHandle);
    intervalHandle = null;
  }
}

/**
 * Instructions the user should follow to get the updated version.
 * This varies by platform:
 *  - Excel Online: clear browser cache or hard-refresh (Ctrl+Shift+R)
 *  - Excel Desktop: clear the Office cache and restart
 *  - Both: an admin may need to re-deploy the manifest if the update
 *    includes manifest changes
 */
export function getUpdateInstructions(latestVersion: string): string {
  return [
    `A new version (v${latestVersion}) of MT Data Connector is available. You are running v${ADDIN_VERSION}.`,
    "",
    "To update:",
    "",
    "Excel Online (Microsoft 365 web):",
    "  1. Close all workbooks using this add-in",
    "  2. Hard-refresh your browser (Ctrl+Shift+R / Cmd+Shift+R)",
    "  3. Re-open your workbook — the updated add-in will load automatically",
    "",
    "Excel Desktop (Windows/Mac):",
    "  1. Close Excel completely",
    "  2. Clear the Office web cache:",
    "     Windows: delete contents of %LOCALAPPDATA%\\Microsoft\\Office\\16.0\\Wef\\",
    "     Mac: ~/Library/Containers/com.microsoft.Excel/Data/Library/Caches/",
    "  3. Re-open Excel and your workbook",
    "",
    "If your organization deployed this add-in centrally, your admin may",
    "need to update the deployment. Contact your IT administrator.",
  ].join("\n");
}
