/** @type {import('next').NextConfig} */
const isGithubPages = process.env.GITHUB_PAGES === 'true';

const nextConfig = {
  reactStrictMode: true,
  output: isGithubPages ? 'export' : undefined,
  basePath: isGithubPages ? '/Playwright' : '',
  assetPrefix: isGithubPages ? '/Playwright/' : undefined,
  env: {
    NEXT_PUBLIC_BASE_PATH: isGithubPages ? '/Playwright' : '',
  },
  images: {
    unoptimized: true,
  },
}

module.exports = nextConfig
