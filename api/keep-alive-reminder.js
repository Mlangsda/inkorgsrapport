import { ConfidentialClientApplication } from '@azure/msal-node';

const USER_EMAIL = 'marzena@marzenalangsdale.com';
const TARGET_DATE = '2026-08-07';

function stockholmDateString() {
  return new Intl.DateTimeFormat('sv-SE', {
    timeZone: 'Europe/Stockholm',
    year: 'numeric',
    month: '2-digit',
    day: '2-digit',
  }).format(new Date());
}

export default async function handler(req, res) {
  const today = stockholmDateString();

  if (today !== TARGET_DATE) {
    return res.status(200).json({ skipped: true, today, targetDate: TARGET_DATE });
  }

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
      `https://graph.microsoft.com/v1.0/users/${USER_EMAIL}/sendMail`,
      {
        method: 'POST',
        headers: {
          Authorization: `Bearer ${tokenResult.accessToken}`,
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          message: {
            subject: 'Supabase keep-alive — dags att stänga av?',
            body: {
              contentType: 'HTML',
              content:
                '<p>Hej Marzena,</p>' +
                '<p>Nu har det gått ca 4 veckor sedan jag satte upp en keep-alive-routine som pingar Supabase var 5:e dag, så att projektet inte pausas av inaktivitet medan du var på semester.</p>' +
                '<p>Är du tillbaka nu? Säg till mig i Claude Code om routinen ska stängas av, eller om du vill att den fortsätter ett tag till.</p>' +
                '<p>/ Claude</p>',
            },
            toRecipients: [{ emailAddress: { address: USER_EMAIL } }],
            from: { emailAddress: { address: USER_EMAIL, name: 'Marzena Langsdale' } },
          },
        }),
      }
    );

    if (!response.ok) {
      const error = await response.text();
      return res.status(response.status).json({ error });
    }

    return res.status(200).json({ sent: true, today });
  } catch (err) {
    return res.status(500).json({ error: err.message });
  }
}
