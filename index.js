module.exports = (req, res) => {
  if (req.method === 'POST') {
    console.log('Webhook received:', req.body);
    res.json({ text: `Echo: ${req.body?.text || 'Hello!'}` });
  } else {
    res.json({ status: 'Webhook server running' });
  }
};