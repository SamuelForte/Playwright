/** @type {import('next').NextConfig} */
const isGithubPages = process.env.GITHUB_PAGES === 'true';

const nextConfig = {
  reactStrictMode: true,
  output: isGithubPages ? 'export' : undefined,
  trailingSlash: isGithubPages,
  basePath: isGithubPages ? '/Playwright' : '',
  assetPrefix: isGithubPages ? '/Playwright/' : undefined,
  env: {
    NEXT_PUBLIC_GH_PAGES_MODE: isGithubPages ? 'true' : 'false',
  },
  images: {
    unoptimized: true,
  },
}

module.exports = nextConfig
