const express = require("express");  // ✅ Import Express
const app = express();               // ✅ Create an Express app instance

app.use(express.json());            // ✅ Middleware to parse JSON body

app.get("/", (req, res) => {
  res.send("✅ Outlook Actionable Message Handler is running!");
});

app.post("/action", async (req, res) => {
  const { ApprovalStatus, Description, ID, listName } = req.body;

  console.log("✅ Received a response:");
  console.log("ApprovalStatus:", ApprovalStatus);
  console.log("Description:", Description);
  console.log("ID:", ID);
  console.log("List Name:", listName);

  res.status(200).send({
    type: "MessageCard",
    text: `✅ You selected: ${ApprovalStatus}. Thank you!`
  });
});

// ✅ Start the server
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`🚀 Server running on port ${PORT}`);
});
