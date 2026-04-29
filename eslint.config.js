import js from "@eslint/js";
import globals from "globals";

// Identifiers exported from utils.js into the global scope (file-shared `var`
// and `function` declarations). Page scripts read these; utils.js declares
// them, so they're listed as readonly globals only in the page-script block.
const UTILS_EXPORTS = {
  CLIPGEN_ANIMATED_BG: "readonly",
  CLIPGEN_CONFIG: "writable",
  clipgenApplyConfig: "readonly",
  qs: "readonly",
  qsa: "readonly",
  el: "readonly",
  isVideoLoop: "readonly",
  createLoopVideo: "readonly",
  createTooltip: "readonly",
  attachHoverTooltip: "readonly",
  positionTooltipAnchored: "readonly",
  pad2: "readonly",
  formatTime: "readonly",
  formatDuration: "readonly",
  artifactDurationSec: "readonly",
  truncate: "readonly",
  parseTimestamp: "readonly",
  parseClockTimestamp: "readonly",
  parseClipSegmentsForCell: "readonly",
  clamp: "readonly",
  debounce: "readonly",
  escapeHtml: "readonly",
  hexToRgba: "readonly",
  SHOW_TOAST_DEFAULT_MS: "readonly",
  showToast: "readonly",
  severityClass: "readonly",
  severityRank: "readonly",
  severityRankByClass: "readonly",
  apiGet: "readonly",
  apiPost: "readonly",
  apiPut: "readonly",
  apiDelete: "readonly",
  POLL_INTERVAL: "readonly",
  MARK_CATEGORIES: "writable",
  setMarkCategories: "readonly",
  XREF_BADGES: "readonly",
  DETECTOR_COLORS: "writable",
  refreshDetectorColors: "readonly",
  THEME_STORAGE_KEY: "readonly",
  TOOLTIP_STORAGE_KEY: "readonly",
  applyStoredThemePreference: "readonly",
  toggleThemePreference: "readonly",
  updateThemeToggleButton: "readonly",
  initThemeToggle: "readonly",
  initFrontendSwitcher: "readonly",
  getStoredTooltipPref: "readonly",
  setStoredTooltipPref: "readonly",
  UI_STATE_STORAGE_KEY: "readonly",
  getStoredUIState: "readonly",
  setStoredUIStateField: "readonly",
};

// Identifiers exported by other shared scripts loaded before page scripts.
const SHARED_PAGE_GLOBALS = {
  CLIPGEN_DATA: "readonly",
  // settings-modal.js attaches to window.openSettingsModal; some page scripts
  // call it bare, others use the window. form. Both are valid.
  openSettingsModal: "readonly",
};

const SHARED_RULES = {
  "no-console": "off",
  "no-empty": ["error", { allowEmptyCatch: true }],
  "no-unused-vars": [
    "error",
    {
      argsIgnorePattern: "^_",
      varsIgnorePattern: "^_",
      caughtErrors: "all",
      caughtErrorsIgnorePattern: "^_",
    },
  ],
  // Allow `x == null` / `x != null` — used as a deliberate "null or undefined"
  // check throughout the codebase. Strict equality everywhere else.
  eqeqeq: ["error", "always", { null: "ignore" }],
};

export default [
  { ignores: ["build/**", "dist/**", "output/**"] },
  js.configs.recommended,

  // Page scripts: utils helpers are pre-declared globals.
  {
    files: ["assets/web/**/*.js"],
    ignores: ["assets/web/utils.js"],
    languageOptions: {
      ecmaVersion: 2020,
      sourceType: "script",
      globals: {
        ...globals.browser,
        ...UTILS_EXPORTS,
        ...SHARED_PAGE_GLOBALS,
      },
    },
    rules: SHARED_RULES,
  },

  // utils.js: declares the helpers itself, so those names are NOT pre-declared
  // globals. `vars: "local"` keeps unused-var warnings active for in-function
  // dead code without flagging top-level exports that other files consume.
  {
    files: ["assets/web/utils.js"],
    languageOptions: {
      ecmaVersion: 2020,
      sourceType: "script",
      globals: { ...globals.browser },
    },
    rules: {
      ...SHARED_RULES,
      "no-unused-vars": [
        "error",
        {
          args: "after-used",
          argsIgnorePattern: "^_",
          vars: "local",
          caughtErrors: "all",
          caughtErrorsIgnorePattern: "^_",
        },
      ],
    },
  },
];
