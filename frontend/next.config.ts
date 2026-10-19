import type { NextConfig } from "next";

const nextConfig: NextConfig = {
  reactStrictMode: true,
  // Self-contained server bundle for the Docker image (frontend/Dockerfile), or plain static
  // files for the desktop app, where the backend serves them (STATIC_EXPORT=1 npm run build).
  output: process.env.STATIC_EXPORT ? "export" : "standalone",
};

export default nextConfig;
