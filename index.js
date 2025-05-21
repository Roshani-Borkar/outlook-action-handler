const express = require('express');
const app = express();
const port = process.env.PORT || 3000;

app.use(express.json());

// Endpoint to receive Outlook action responses
app.post('/action', (req, res) => {
  console.log('Received Action:', req.body);

  // Example: extract info from request
  const action = req.body && req.body.action || "Unknown";

  // Return a basic confirmation
  res.json({
    type: "MessageCard",
    text: `✅ Action "${action}" received and processed.`
  });
});

app.get('/', (req, res) => {
  res.send('Outlook Actionable Message handler is running!');
});

app.listen(port, () => {
  console.log(`Server running on port ${port}`);
});

