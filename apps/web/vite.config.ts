import { defineConfig } from 'vite'
import vue from '@vitejs/plugin-vue'
import path from 'node:path'

export default defineConfig({
  plugins: [vue()],
  resolve: {
    alias: {
      '@': path.resolve(__dirname, 'src'),
    },
  },
  server: {
    port: 5176,
    strictPort: true,
    proxy: {
      // SPA route /auth/callback must not be proxied to the API
      '/auth': {
        target: 'http://localhost:3000',
        bypass(req) {
          if (req.url?.startsWith('/auth/callback')) {
            return req.url
          }
        },
      },
      '/files': 'http://localhost:3000',
      '/users': 'http://localhost:3000',
      '/ai': 'http://localhost:3000',
    },
  },
  optimizeDeps: {
    include: [
      'react',
      'react-dom',
      'rxjs',
      '@mertdeveci55/univer-import-export',
      'xlsx',
    ],
  },
})
