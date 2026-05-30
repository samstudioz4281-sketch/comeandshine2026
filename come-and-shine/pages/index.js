// This file exists as a fallback.
// The actual page is served by /api/page via rewrites in next.config.js
export default function Home() {
  if (typeof window !== 'undefined') {
    window.location.href = '/api/page';
  }
  return null;
}
