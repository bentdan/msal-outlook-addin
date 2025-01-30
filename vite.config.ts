import { getHttpsServerOptions } from 'office-addin-dev-certs';
import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';
import tsconfigPaths from 'vite-tsconfig-paths';
import { viteStaticCopy } from 'vite-plugin-static-copy';
import path from 'path';
import { coverageConfigDefaults } from 'vitest/config';

export default defineConfig(async ({ command, mode }) => {
  const isDev = command === 'serve' && mode === 'development';

  return {
    plugins: [
      react(),
      tsconfigPaths(),
      viteStaticCopy({
        targets: [{
          src: 'node_modules/@microsoft/office-js/dist/*',
          dest: 'assets/office-js',
        }],
      }),
    ],
    build: {
      outDir: 'build',
      target: 'esnext',
      rollupOptions: {
        input: {
          main: path.resolve(__dirname, 'index.html'),
          login: path.resolve(__dirname, 'login.html'),
        },
      },
    },
    server: {
      https: isDev ? await getHttpsServerOptions() : undefined,
      port: 3000
    },
    test: {
      coverage: {
        exclude: [
          'src/setupTests.ts',
          ...coverageConfigDefaults.exclude,
        ],
        include: ['src/**'],
        reporter: ['lcov'],
      },
      environment: 'jsdom',
      globals: true,
      testTimeout: 60000,
      teardownTimeout: 30000,
      include: ['src/**/*.spec.*'],
      reporters: [
        'default',
        ['vitest-sonar-reporter', { outputFile: 'coverage/genericcoverage.xml' }],
      ],
      setupFiles: ['./src/setupTests.ts'],
    },
  };
});
