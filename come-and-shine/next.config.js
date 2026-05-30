/** @type {import('next').NextConfig} */
const nextConfig = {
  async rewrites() {
    return {
      // Run BEFORE Next.js page routing
      beforeFiles: [
        {
          source: '/',
          destination: '/index.html',
        },
      ],
    };
  },
};

module.exports = nextConfig;
