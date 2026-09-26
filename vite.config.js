import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

export default defineConfig({
  // GitHub Pages serves this repo at /BenchPortal_ElevateCross/, so
  // asset URLs need that prefix in the production build. Local dev
  // and preview keep serving from the root.
  base: process.env.GITHUB_PAGES ? "/BenchPortal_ElevateCross/" : "/",
  plugins: [react()],
});
