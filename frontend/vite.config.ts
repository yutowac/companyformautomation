import { defineConfig } from 'vite'

export default defineConfig({
  publicDir: 'public',
  server: {
    port: 3000,
    proxy: {
      '/generate-word': {
        target: 'http://localhost:10000',
        changeOrigin: true,
      },
      '/generate-word2': {
        target: 'http://localhost:10000',
        changeOrigin: true,
      },
      '/generate-excel': {
        target: 'http://localhost:10000',
        changeOrigin: true,
      },
      '/get-created-word': {
        target: 'http://localhost:10000',
        changeOrigin: true,
      },
      '/get-created-word2': {
        target: 'http://localhost:10000',
        changeOrigin: true,
      },
      '/get-created-excel': {
        target: 'http://localhost:10000',
        changeOrigin: true,
      },
    },
  },
})

