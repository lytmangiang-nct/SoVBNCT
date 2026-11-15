import type { IncomingMessage, ServerResponse } from 'http';
import app from '../server.js';

export default function handler(req: IncomingMessage & { url?: string }, res: ServerResponse) {
  if (req.url) {
    try {
      const parsed = new URL(req.url, 'http://localhost');
      const vercelPath = parsed.searchParams.get('__vercel_path');
      if (vercelPath !== null) {
        parsed.searchParams.delete('__vercel_path');
        const qs = parsed.searchParams.toString();
        if (parsed.pathname === '/api/index' || parsed.pathname === '/api') {
          const cleanRoute = vercelPath.replace(/^\/+/, '');
          req.url = `/api/${cleanRoute}${qs ? `?${qs}` : ''}`;
        } else {
          req.url = `${parsed.pathname}${qs ? `?${qs}` : ''}`;
        }
      }
    } catch {
      // ignore malformed URL
    }
  }

  return app(req as any, res as any);
}
