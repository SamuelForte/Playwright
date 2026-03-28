/** @type {import('next').NextConfig} */
const isGithubPages = process.env.GITHUB_PAGES === 'true';

const nextConfig = {
  reactStrictMode: true,
  output: isGithubPages ? 'export' : undefined,
  trailingSlash: isGithubPages,
  basePath: isGithubPages ? '/Playwright' : '',
  assetPrefix: isGithubPages ? '/Playwright/' : undefined,
  images: {
    unoptimized: true,
  },
}

module.exports = nextConfig
