// index.js
const express = require("express");
const bodyParser = require("body-parser");
const axios = require("axios");
const getAccessToken = require("./auth");

const app = express();
app.use(bodyParser.json());

const tenantId = process.env.TENANT_ID;
const clientId = process.env.CLIENT_ID;
const clientSecret = process.env.CLIENT_SECRET;
const siteId = process.env.SITE_ID;
const listId = process.env.LIST_ID;

app.post("/response", async (req, res) => {
  const { ApprovalStatus, Description, ID } = req.body;

  console.log("🔄 Received Adaptive Card response:", req.body);

  try {
    console.log("Received body:", req.body);
    console.log("Environment:", {
      tenantId,
      clientId,
      siteId,
      listId
    });

    const token = await getAccessToken(tenantId, clientId, clientSecret);

    const url = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/items/${ID}/fields`;

    const updatePayload = {
      Status: ApprovalStatus,
      Comments: Description
    };
console.log("PATCH URL:", url);
console.log("Payload:", updatePayload);

    await axios.patch(url, updatePayload, {
      headers: {
        Authorization: `Bearer ${token}`,
        "Content-Type": "application/json"
      }
    });

    console.log("✅ SharePoint list item updated successfully");

    res.status(200).send({
      type: "MessageCard",
      text: `✅ SharePoint item updated with status: ${ApprovalStatus}`
    });

  } catch (error) {
    console.error("Full error:", error);

    console.error("❌ Error updating SharePoint:", error.response?.data || error.message);
    res.status(500).send({
      type: "MessageCard",
      text: `❌ Error updating SharePoint: ${error.message}`
    });
  }
});

app.get("/", (req, res) => {
  res.send("✅ Outlook Handler is running!");
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`✅ Server running on port ${PORT}`));
