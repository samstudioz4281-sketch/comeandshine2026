/** @type {import('next').NextConfig} */
const nextConfig = {
  async rewrites() {
    return [
      {
        source: '/',
        destination: '/api/page',
      },
    ];
  },
};

module.exports = nextConfig;
