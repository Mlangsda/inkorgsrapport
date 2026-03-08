import { ConfidentialClientApplication } from '@azure/msal-node';

const USER_EMAIL = 'marzena@marzenalangsdale.com';

export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') return res.status(200).end();
  if (req.method !== 'POST') return res.status(405).json({ error: 'Method not allowed' });

  const { messageId } = req.body;
  if (!messageId) return res.status(400).json({ error: 'Missing messageId' });

  try {
    const cca = new ConfidentialClientApplication({
      auth: {
        clientId: process.env.MS_CLIENT_ID,
        authority: `https://login.microsoftonline.com/${process.env.MS_TENANT_ID}`,
        clientSecret: process.env.MS_CLIENT_SECRET,
      },
    });

    const tokenResult = await cca.acquireTokenByClientCredential({
      scopes: ['https://graph.microsoft.com/.default'],
    });

    const response = await fetch(
      `https://graph.microsoft.com/v1.0/users/${USER_EMAIL}/messages/${messageId}`,
      {
        method: 'PATCH',
        headers: {
          Authorization: `Bearer ${tokenResult.accessToken}`,
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({ isRead: true }),
      }
    );

    if (!response.ok) {
      const error = await response.text();
      return res.status(response.status).json({ error });
    }

    return res.status(200).json({ success: true });
  } catch (err) {
    return res.status(500).json({ error: err.message });
  }
}
