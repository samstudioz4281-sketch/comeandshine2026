// This page serves the full static HTML directly.
// The API route /api/scan handles all Groq calls server-side.
export default function Home() {
  // Rendered server-side — see getServerSideProps
  return null;
}

export async function getServerSideProps({ res }) {
  const { getPageHTML } = await import('../lib/pageHTML');
  res.setHeader('Content-Type', 'text/html; charset=utf-8');
  res.write(getPageHTML());
  res.end();
  return { props: {} };
}
