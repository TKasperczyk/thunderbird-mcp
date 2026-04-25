/**
 * Module-singleton mutable state shared across the MCP server.
 *
 * ESM modules are cached per resource:// URL, so a Set/WeakSet exported from
 * here is the same instance everywhere it's imported. This is the shared-
 * state strategy chosen during the modular split (see api.js header).
 *
 * Cleared by api.js onShutdown(). Do not put process-lifecycle handles
 * (server, startPromise) here -- those live on globalThis so they survive
 * api.js sandbox reloads.
 */

// Active attach-injection timers. Held to prevent GC before they fire.
// compose code inserts; onShutdown clears.
export const _attachTimers = new Set();

// Filesystem paths of temporary files created for inline base64 attachments.
// compose code inserts; onShutdown deletes the files and clears.
export const _tempAttachFiles = new Set();

// Compose windows already claimed by an in-flight replyToMessage call, so
// concurrent replies to the same original message never bind two observers
// to the same compose window (which would double-inject body/attachments).
//
// MUST stay a WeakSet -- entries are auto-collected when the window dies,
// and that's the only mechanism preventing this from leaking. Do not switch
// to a regular Set for "easier debugging".
//
// @type {WeakSet<Window>}
export const _claimedReplyComposeWindows = new WeakSet();
