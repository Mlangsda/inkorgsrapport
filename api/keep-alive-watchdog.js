import { ConfidentialClientApplication } from '@azure/msal-node';

const USER_EMAIL = 'marzena@marzenalangsdale.com';
// Larmgräns: keep-alive-cronen i ideas-dashboard skriver varje morgon 05:00 UTC.
// Är stämpeln äldre än 26 h har minst en körning missats → larma via mejl.
const MAX_AGE_HOURS = 26;

async function sendAlarm(subject, htmlBody) {
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
          subject,
          body: { contentType: 'HTML', content: htmlBody },
          toRecipients: [{ emailAddress: { address: USER_EMAIL } }],
          from: { emailAddress: { address: USER_EMAIL, name: 'Marzena Langsdale' } },
        },
      }),
    }
  );

  if (!response.ok) {
    throw new Error(`sendMail misslyckades: ${response.status} ${await response.text()}`);
  }
}

export default async function handler(req, res) {
  const url = process.env.SUPABASE_URL;
  const anonKey = process.env.SUPABASE_ANON_KEY;

  if (!url || !anonKey) {
    return res.status(500).json({ error: 'Supabase env vars saknas' });
  }

  let problem = null;
  let pingedAt = null;

  try {
    const response = await fetch(`${url}/rest/v1/_keep_alive?select=pinged_at&id=eq.1`, {
      headers: { apikey: anonKey, Authorization: `Bearer ${anonKey}` },
    });

    if (!response.ok) {
      problem = `Supabase svarade ${response.status} — projektet kan vara pausat eller otillgängligt.`;
    } else {
      const rows = await response.json();
      pingedAt = rows[0]?.pinged_at ?? null;
      if (!pingedAt) {
        problem = 'Heartbeat-raden i _keep_alive saknas.';
      } else {
        const ageHours = (Date.now() - new Date(pingedAt).getTime()) / 3600000;
        if (ageHours > MAX_AGE_HOURS) {
          problem = `Senaste keep-alive-pingen är ${ageHours.toFixed(1)} timmar gammal (gräns ${MAX_AGE_HOURS} h) — cronen i ideas-dashboard verkar inte köra.`;
        }
      }
    }
  } catch (err) {
    problem = `Kunde inte nå Supabase alls: ${err.message}`;
  }

  if (!problem) {
    return res.status(200).json({ ok: true, pingedAt });
  }

  try {
    await sendAlarm(
      'LARM: Supabase keep-alive fungerar inte',
      '<p>Hej Marzena,</p>' +
        `<p>Vakthunden hittade ett problem med Supabase keep-alive i morse:</p>` +
        `<p><strong>${problem}</strong></p>` +
        '<p>Risk: Supabase free-tier pausar projektet efter 7 dagars inaktivitet. Öppna Claude Code och be mig felsöka, eller kolla själv i Supabase-dashboarden.</p>' +
        '<p>/ Claude (keep-alive-vakthunden i inkorgsrapport)</p>'
    );
    return res.status(200).json({ ok: false, problem, alarmSent: true });
  } catch (err) {
    return res.status(500).json({ ok: false, problem, alarmSent: false, alarmError: err.message });
  }
}
