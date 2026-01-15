import http from 'http';
import dotenv from 'dotenv';
import webhookHandler from './api/webhook.js';

dotenv.config({ path: '.env.local' });

const PORT = process.env.PORT || 3000;

// Helper to add Express-like methods to response
function enhanceResponse(res) {
  res.status = (code) => {
    res.statusCode = code;
    return res;
  };
  res.json = (data) => {
    res.setHeader('Content-Type', 'application/json');
    res.end(JSON.stringify(data));
  };
  return res;
}

const server = http.createServer(async (req, res) => {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type, Authorization');

  enhanceResponse(res);

  if (req.method === 'OPTIONS') {
    res.status(200).end();
    return;
  }

  if (req.url === '/health') {
    res.status(200).json({ status: 'ok', timestamp: new Date().toISOString() });
    return;
  }

  if (req.url === '/webhook' && req.method === 'POST') {
    let body = '';
    req.on('data', chunk => body += chunk);
    req.on('end', async () => {
      try {
        req.body = JSON.parse(body);
        await webhookHandler(req, res);
      } catch (error) {
        console.error('Webhook error:', error);
        res.status(500).json({ error: 'Internal server error' });
      }
    });
    return;
  }

  res.status(404).json({ error: 'Not found' });
});

server.listen(PORT, () => {
  console.log(`🤖 Gent-Chat running on port ${PORT}`);
  console.log(`📡 Webhook: http://localhost:${PORT}/webhook`);
  console.log(`❤️  Health: http://localhost:${PORT}/health`);
});
