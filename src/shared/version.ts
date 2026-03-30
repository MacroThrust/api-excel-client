/**
 * Build-embedded version information.
 *
 * Values are injected at compile time by webpack DefinePlugin from package.json.
 * They are literal string constants in the compiled bundle — no runtime file
 * reads or network requests required.
 */

declare const __ADDIN_VERSION__: string;
declare const __ADDIN_NAME__: string;
declare const __BUILD_TIMESTAMP__: string;
declare const __ADDIN_HOST__: string;

export const ADDIN_VERSION: string = __ADDIN_VERSION__;
export const ADDIN_NAME: string = __ADDIN_NAME__;
export const BUILD_TIMESTAMP: string = __BUILD_TIMESTAMP__;
export const ADDIN_HOST: string = __ADDIN_HOST__;
