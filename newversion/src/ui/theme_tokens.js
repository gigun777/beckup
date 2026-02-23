/**
 * Central design tokens for UI themes.
 *
 * NOTE:
 * - Tokens are declared here as a source of truth for JS-side tooling.
 * - Runtime theme switching is performed by changing `data-theme` on `<html>`.
 * - CSS contains token declarations for each theme selector.
 */
(function attachThemeTokens(global) {
  const TOKENS = Object.freeze({
    light: Object.freeze({
      '--bg': '#f5f7fb',
      '--text': '#1f2937',
      '--panel': '#ffffff',
      '--border': '#d7deea',
      '--accent': '#3b82f6',
      '--danger': '#dc2626',
      '--tableBg': '#ffffff',
      '--tableHeadBg': '#edf2f8',
      '--tableHeadBg2': '#e4ebf4',
      '--tableGrid': '#d7deea',
      '--radius': '12px',
      '--gap': '12px',
      '--fontSize': '14px',
      '--controlH': '36px'
    }),
    dark: Object.freeze({
      '--bg': '#1f242c',
      '--text': '#e6e9ef',
      '--panel': '#2a313c',
      '--border': '#414b5b',
      '--accent': '#63a4ff',
      '--danger': '#f27d7d',
      '--tableBg': '#252c36',
      '--tableHeadBg': '#303a47',
      '--tableHeadBg2': '#384455',
      '--tableGrid': '#485568',
      '--radius': '12px',
      '--gap': '12px',
      '--fontSize': '14px',
      '--controlH': '36px'
    }),
    gray: Object.freeze({
      '--bg': '#eceff3',
      '--text': '#222831',
      '--panel': '#f8fafc',
      '--border': '#bcc6d4',
      '--accent': '#4b7abf',
      '--danger': '#c84d4d',
      '--tableBg': '#f8fafc',
      '--tableHeadBg': '#dde4ed',
      '--tableHeadBg2': '#d2dae6',
      '--tableGrid': '#b8c2d1',
      '--radius': '12px',
      '--gap': '12px',
      '--fontSize': '14px',
      '--controlH': '36px'
    }),
    contrast: Object.freeze({
      '--bg': '#101214',
      '--text': '#f8fafc',
      '--panel': '#1a1e22',
      '--border': '#6b7280',
      '--accent': '#00cfff',
      '--danger': '#ff6363',
      '--tableBg': '#12161a',
      '--tableHeadBg': '#252c34',
      '--tableHeadBg2': '#313a45',
      '--tableGrid': '#707b8a',
      '--radius': '12px',
      '--gap': '12px',
      '--fontSize': '14px',
      '--controlH': '36px'
    })
  });

  function getThemeNames() {
    return Object.keys(TOKENS);
  }

  global.UI_THEME_TOKENS = TOKENS;
  global.UI_THEME_META = Object.freeze({ getThemeNames });
})(typeof window !== 'undefined' ? window : globalThis);
