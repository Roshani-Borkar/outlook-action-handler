app.post("/action", async (req, res) => {
  const { ApprovalStatus, Description, ID, listName } = req.body;

  console.log("✅ Received a response:");
  console.log("ApprovalStatus:", ApprovalStatus);
  console.log("Description:", Description);
  console.log("ID:", ID);
  console.log("List Name:", listName);

  // Here you can later add code to update SharePoint using Graph API

  res.status(200).send({
    type: "MessageCard",
    text: `✅ You selected: ${status}. Thank you!`
  });
});
