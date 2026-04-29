import js from "@eslint/js";
import globals from "globals";

export default [
  {
    ignores: ["build/**", "dist/**", "output/**"],
  },
  js.configs.recommended,
  {
    files: ["assets/web/**/*.js"],
    languageOptions: {
      ecmaVersion: 2020,
      sourceType: "script",
      globals: {
        ...globals.browser,
        CLIPGEN_DATA: "readonly",
        CLIPGEN_CONFIG: "writable",
        clipgenApplyConfig: "readonly",
        createLoopVideo: "readonly",
        el: "readonly",
        getArtifactType: "readonly",
        getSeverityClass: "readonly",
        getSeverityRank: "readonly",
        getSeverityRankByClass: "readonly",
        getTaskColor: "readonly",
        iconSpan: "readonly",
        isVideoLoop: "readonly",
        normalizeSeverityLabel: "readonly",
        parseTimestampToSeconds: "readonly",
        qs: "readonly",
        qsa: "readonly",
        renderGridBackground: "readonly",
        setupSettingsModal: "readonly",
        severityClassFromRank: "readonly",
        severityLabelFromRank: "readonly",
        showSettingsModal: "readonly",
        sortSeverityRanks: "readonly",
        truncateText: "readonly",
      },
    },
    rules: {
      "no-console": "off",
      "no-unused-vars": [
        "warn",
        {
          argsIgnorePattern: "^_",
          varsIgnorePattern: "^CLIPGEN_",
        },
      ],
    },
  },
];
