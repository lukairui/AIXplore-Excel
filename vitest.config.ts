import { defineConfig } from 'vitest/config'
import vue from '@vitejs/plugin-vue'
import path from 'node:path'

export default defineConfig({
  test: {
    environment: 'jsdom',
    globals: false,
    include: ['src/renderer/src/excel/lib/**/*.test.ts'],
    setupFiles: ['src/renderer/src/excel/lib/test-setup.ts']
  },
  plugins: [vue()],
  resolve: {
    alias: {
      '@': path.resolve(__dirname, './src/renderer/src')
    }
  }
})
