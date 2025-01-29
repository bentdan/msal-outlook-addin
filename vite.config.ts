import { getHttpsServerOptions } from 'office-addin-dev-certs';
import { defineConfig } from 'vite';
import { coverageConfigDefaults } from 'vitest/config';

export default defineConfig(async ({ command, mode }) => ({
  build: {
    outDir: 'build',
    target: 'esnext',
  },
  server: {
    https: command === 'serve' && mode === 'development'
      ? await getHttpsServerOptions()
      : undefined, // skip dev cert installation in build/CI
    port: 3000
  },
  test: {
    coverage: {
      exclude: [
        'src/index.tsx',
        ...coverageConfigDefaults.exclude
      ],
      include: [
        'src/**'
      ],
      reporter: ['lcov']
    },
    environment: 'jsdom',
    globals: true,
    include: [
      'src/**/*.spec.*'
    ],
    reporters: [
      'default',
      ['vitest-sonar-reporter', { outputFile: 'coverage/genericcoverage.xml' }]
    ]
  }
}));
