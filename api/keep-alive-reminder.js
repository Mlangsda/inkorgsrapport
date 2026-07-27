import { ConfidentialClientApplication } from '@azure/msal-node';

const USER_EMAIL = 'marzena@marzenalangsdale.com';
// Påminner den 7:e varje månad (första gången 2026-08-07) så länge cronen är igång,
// så Marzena aktivt får ta ställning till om Supabase keep-alive ska fortsätta.
const REMINDER_DAY = '07';
const FIRST_REMINDER = '2026-08-07';

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

  if (today < FIRST_REMINDER || !today.endsWith(`-${REMINDER_DAY}`)) {
    return res.status(200).json({ skipped: true, today, remindsOn: `den ${REMINDER_DAY}:e varje månad fr.o.m. ${FIRST_REMINDER}` });
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
            subject: 'Supabase keep-alive — ska den fortsätta?',
            body: {
              contentType: 'HTML',
              content:
                '<p>Hej Marzena,</p>' +
                '<p>Månadspåminnelse: en Vercel-cron i ideas-dashboard pingar Supabase varje morgon (skriv-anrop mot tabellen _keep_alive) så att projektet inte pausas av inaktivitet på free-tier.</p>' +
                '<p>Vill du att den fortsätter? Den är gratis och harmlös, så standardvalet är att låta den rulla. Om du vill stänga av den, eller uppgradera Supabase till Pro istället, säg till mig i Claude Code.</p>' +
                '<p>Du får den här påminnelsen den 7:e varje månad tills du ber mig ta bort den.</p>' +
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
