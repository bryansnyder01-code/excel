// Builds the screens as static files for the desktop app (into out/), served by the backend
// from the same address, so API calls use relative URLs. Works the same on Windows, Mac and Linux.
import { spawnSync } from "node:child_process";

const env = { ...process.env, STATIC_EXPORT: "1", NEXT_PUBLIC_API_URL: "" };
const r = spawnSync("npx", ["next", "build"], { stdio: "inherit", env, shell: process.platform === "win32" });
process.exit(r.status ?? 1);
