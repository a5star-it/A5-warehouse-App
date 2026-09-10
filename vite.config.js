import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

export default defineConfig({
  plugins: [react()],
  build: {
    outDir: "dist",
  },
  server: {
    // In local dev, proxy API calls to the Express server on :3000
    proxy: {
      "/api": "http://localhost:3000",
    },
  },
});
