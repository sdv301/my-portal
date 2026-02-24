import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'
import path from 'path'

export default defineConfig(({ mode }) => {
  return {
    plugins: [react()],

    server: {
      port: 3000,
      host: true,

      proxy: mode === 'development' ? {
        '/api': 'http://localhost:4000',
      } : undefined,
    },

    resolve: {
      alias: {
        '@': path.resolve(__dirname, './src'),
      },
    },

    build: {
      outDir: 'dist',
      emptyOutDir: true,
      // если хочешь, чтобы билд был максимально оптимизирован
      minify: 'esbuild',
      target: 'esnext',
    },
  }
})