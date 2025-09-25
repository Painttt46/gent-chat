module.exports = (req, res) => {
  if (req.method === 'POST') {
    res.json({ text: `Echo: ${req.body?.text || 'Hello!'}` });
  } else {
    res.json({ status: 'Webhook server running' });
  }
};