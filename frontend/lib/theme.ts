// Light / dark theme. The choice is remembered per browser; "system" follows the OS setting.

export type Theme = "light" | "dark" | "system";
const KEY = "sheetcraft:theme";

export function getTheme(): Theme {
  try {
    const t = localStorage.getItem(KEY);
    return t === "light" || t === "dark" ? t : "system";
  } catch {
    return "system";
  }
}

export function isDark(theme: Theme = getTheme()) {
  return theme === "dark" || (theme === "system" && window.matchMedia("(prefers-color-scheme: dark)").matches);
}

export function applyTheme(theme: Theme) {
  try {
    if (theme === "system") localStorage.removeItem(KEY);
    else localStorage.setItem(KEY, theme);
  } catch {}
  document.documentElement.classList.toggle("dark", isDark(theme));
}

/** Runs before the page paints (inlined in app/layout.tsx) so there's no flash of the wrong theme. */
export const THEME_BOOT = `try{var t=localStorage.getItem("${KEY}");if(t==="dark"||(t!=="light"&&matchMedia("(prefers-color-scheme: dark)").matches))document.documentElement.classList.add("dark")}catch(e){}`;
