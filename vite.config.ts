import { defineConfig, loadEnv } from 'vite'
import react from '@vitejs/plugin-react'
import { nodePolyfills } from 'vite-plugin-node-polyfills'
import { recommendationsPlugin } from './vite-plugin-recommendations'

export default defineConfig(({ mode }) => {
  // Load .env into process.env so the middleware can read ANTHROPIC_API_KEY
  Object.assign(process.env, loadEnv(mode, process.cwd(), ''))
  return {
  plugins: [
    react(),
    recommendationsPlugin(),
    nodePolyfills({
      include: ['buffer', 'stream', 'util', 'events'],
      globals: { Buffer: true, process: true },
    }),
  ],
  optimizeDeps: {
    include: ['exceljs'],
  },
  build: {
    target: 'esnext',
    rollupOptions: {
      output: {
        manualChunks: {
          exceljs: ['exceljs'],
        },
      },
    },
  },
  }
})
