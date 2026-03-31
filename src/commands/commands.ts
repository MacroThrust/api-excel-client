/**
 * Ribbon command handlers.
 *
 * These functions are referenced by <FunctionName> elements in manifest.xml.
 * They execute in the shared runtime, so they have full access to auth state.
 */

import { signIn, signOut, initAuth } from "../auth/authConfig";
import { getAuthState } from "../shared/state";
import { getConfig, setApiBaseUrl } from "../shared/config";
import { reloadFunctions } from "../functions/dynamicRegistry";

async function ensureAuthInitialized(): Promise<void> {
  try {
    await initAuth();
  } catch {
    // Already initialized or failed — handled elsewhere
  }
}

/**
 * Sign In command — triggered from the ribbon menu.
 */
async function loginCommand(event: Office.AddinCommands.Event): Promise<void> {
  try {
    await ensureAuthInitialized();
    await signIn();

    const state = getAuthState();
    if (state.isAuthenticated) {
      showNotification("Signed In", `Welcome, ${state.userDisplayName ?? state.userEmail ?? "User"}!`);
    }
  } catch (err) {
    showNotification("Sign In Failed", err instanceof Error ? err.message : "Unknown error");
  } finally {
    event.completed();
  }
}

/**
 * Sign Out command — triggered from the ribbon menu.
 */
async function logoutCommand(event: Office.AddinCommands.Event): Promise<void> {
  try {
    await signOut();
    showNotification("Signed Out", "You have been signed out successfully.");
  } catch (err) {
    showNotification("Sign Out Failed", err instanceof Error ? err.message : "Unknown error");
  } finally {
    event.completed();
  }
}

/**
 * Settings command — opens the taskpane to the settings view.
 */
async function settingsCommand(event: Office.AddinCommands.Event): Promise<void> {
  try {
    await Office.addin.showAsTaskpane();
    const config = getConfig();
    showNotification("Settings", `API endpoint: ${config.apiBaseUrl}`);
  } catch (err) {
    showNotification("Settings Error", err instanceof Error ? err.message : "Unknown error");
  } finally {
    event.completed();
  }
}

/**
 * Refresh command — triggers recalculation of all MT custom functions.
 */
async function refreshCommand(event: Office.AddinCommands.Event): Promise<void> {
  try {
    const state = getAuthState();
    if (!state.isAuthenticated) {
      showNotification("Refresh", "Please sign in first.");
      event.completed();
      return;
    }

    await Excel.run(async (context) => {
      context.workbook.application.calculate(Excel.CalculationType.full);
      await context.sync();
    });
    showNotification("Refreshed", "All data has been recalculated.");
  } catch (err) {
    showNotification("Refresh Failed", err instanceof Error ? err.message : "Unknown error");
  } finally {
    event.completed();
  }
}

/**
 * Reload Functions command — fetches the OpenAPI spec and registers
 * dynamic functions filtered by user permissions/scopes.
 */
async function reloadFunctionsCommand(event: Office.AddinCommands.Event): Promise<void> {
  try {
    const state = getAuthState();
    if (!state.isAuthenticated) {
      showNotification("Reload Functions", "Please sign in first.");
      event.completed();
      return;
    }

    const result = await reloadFunctions();
    if (result.error) {
      showNotification("Reload Functions", `Error: ${result.error}`);
    } else {
      showNotification(
        "Functions Loaded",
        `${result.permitted} of ${result.total} endpoint(s) registered.`
          + (result.denied > 0 ? ` ${result.denied} denied by scope restrictions.` : ""),
      );
    }
  } catch (err) {
    showNotification("Reload Functions Failed", err instanceof Error ? err.message : "Unknown error");
  } finally {
    event.completed();
  }
}

function showNotification(title: string, message: string): void {
  try {
    if (Office.context.mailbox) return;
    console.log(`[MT] ${title}: ${message}`);
  } catch {
    console.log(`[MT] ${title}: ${message}`);
  }
}

/* Register command functions on the global scope for the shared runtime */
(globalThis as Record<string, unknown>).loginCommand = loginCommand;
(globalThis as Record<string, unknown>).logoutCommand = logoutCommand;
(globalThis as Record<string, unknown>).settingsCommand = settingsCommand;
(globalThis as Record<string, unknown>).refreshCommand = refreshCommand;
(globalThis as Record<string, unknown>).reloadFunctionsCommand = reloadFunctionsCommand;
