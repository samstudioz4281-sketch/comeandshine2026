import { getPageHTML } from '../../lib/pageHTML';

export default function handler(req, res) {
  res.setHeader('Content-Type', 'text/html; charset=utf-8');
  res.setHeader('Cache-Control', 'no-store');
  res.status(200).end(getPageHTML());
}

export const config = {
  api: {
    bodyParser: false,
    responseLimit: '10mb',
  },
};
