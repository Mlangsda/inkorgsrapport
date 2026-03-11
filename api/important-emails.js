import { ConfidentialClientApplication } from '@azure/msal-node';

const USER_EMAIL = 'marzena@marzenalangsdale.com';

// Avsändare/domäner som ALLTID filtreras bort (nyhetsbrev, notiser, reklam)
const JUNK_DOMAINS = [
  'noreply', 'no-reply', 'notifications', 'newsletter', 'marketing',
  'mailer-daemon', 'postmaster', 'donotreply', 'do-not-reply',
  'accounts.google', 'facebookmail', 'linkedin.com', 'twitter.com',
  'spotify.com', 'netflix.com', 'apple.com', 'microsoft.com',
  'github.com', 'vercel.com', 'supabase.io', 'heroku.com',
  'mailchimp.com', 'sendinblue', 'hubspot', 'klaviyo',
  'shopify.com', 'squarespace.com', 'wordpress.com',
  'zoom.us', 'calendly.com', 'meetup.com',
  'paypal.com', 'stripe.com', 'klarna.com', 'swish',
  'postnord', 'dhl', 'ups.com', 'fedex.com',
  'booking.com', 'hotels.com', 'airbnb.com', 'sas.se', 'norwegian.com',
  'google.com', 'youtube.com', 'instagram.com', 'tiktok.com',
];

// Ämnesord som indikerar automatiska/oviktiga mejl
const JUNK_SUBJECTS = [
  'unsubscribe', 'nyhetsbrev', 'newsletter', 'your order',
  'din beställning', 'kvitto', 'receipt', 'shipping confirmation',
  'leveransbekräftelse', 'welcome to', 'välkommen till',
  'password reset', 'lösenord', 'verify your', 'verifiera',
  'notification', 'avisering', 'automatic reply', 'autosvar',
  'out of office', 'frånvaro',
];

// Nyckelord som indikerar VIKTIGA mejl (bevaras alltid)
const IMPORTANT_KEYWORDS = [
  'faktura', 'invoice', 'offert', 'quote', 'avtal', 'contract',
  'deadline', 'brådskande', 'urgent', 'viktigt', 'important',
  'evolan', 'yobedoo', 'scalex', 'bolagsverket', 'skatteverket',
  'kronofogden', 'försäkringskassan', 'advokat', 'juridik',
  'förfallen', 'past due', 'betalningspåminnelse', 'payment reminder',
  'samarbete', 'cooperation', 'partnership', 'möte', 'meeting',
];

// Kategorisering baserat på avsändare/ämne
function categorize(email) {
  const from = (email.from?.emailAddress?.address || '').toLowerCase();
  const name = (email.from?.emailAddress?.name || '').toLowerCase();
  const subject = (email.subject || '').toLowerCase();
  const all = from + ' ' + name + ' ' + subject;

  if (all.includes('yobedoo')) return { category: 'yobedoo', icon: '◆', label: 'Yobedoo' };
  if (all.includes('scalex')) return { category: 'scalex', icon: '▲', label: 'Scalex / Företagsutveckling' };
  if (all.includes('bolagsverket') || all.includes('skatteverket') || all.includes('mlc nest'))
    return { category: 'bolag', icon: '■', label: 'MLC Nest AB / Bolagsverket' };

  if (all.includes('faktura') || all.includes('invoice') || all.includes('kronofogden') ||
      all.includes('advokat') || all.includes('förfallen') || all.includes('betalning') ||
      all.includes('skuld') || all.includes('försäkring') || all.includes('kredit'))
    return { category: 'ekonomi', icon: '●', label: 'Ekonomi / Juridik / Fakturor' };

  if (all.includes('evolan') || all.includes('janine') || all.includes('cecilia haldorsen') ||
      all.includes('jonathan brady') || all.includes('domotion') || all.includes('convendum') ||
      all.includes('rawnice') || all.includes('sparkcomm'))
    return { category: 'affar', icon: '★', label: 'Affär / Kund' };

  if (all.includes('försäkringskassan') || all.includes('polisen') || all.includes('myndighet'))
    return { category: 'ovrigt', icon: '◉', label: 'Personligt / Myndigheter' };

  // Default: affärskontakt eller övrigt
  return { category: 'affar', icon: '★', label: 'Affär / Kund' };
}

// Prioritet: high, medium, low
function getPriority(email) {
  const subject = (email.subject || '').toLowerCase();
  const from = (email.from?.emailAddress?.address || '').toLowerCase();
  const importance = email.importance || 'normal';
  const age = Date.now() - new Date(email.receivedDateTime).getTime();
  const daysSince = age / (1000 * 60 * 60 * 24);

  // Exchange-flaggad som hög
  if (importance === 'high') return 'high';

  // Viktiga nyckelord
  if (['brådskande', 'urgent', 'deadline', 'förfallen', 'past due', 'kronofogden'].some(k => subject.includes(k)))
    return 'high';

  // Evolan, Yobedoo = alltid hög
  if (['evolan', 'yobedoo'].some(k => subject.includes(k) || from.includes(k)))
    return 'high';

  // Nyligen (senaste 3 dagarna) = hög
  if (daysSince < 3) return 'high';

  // Senaste veckan = medium
  if (daysSince < 7) return 'medium';

  return 'medium';
}

// Kolla om mejlet ska filtreras bort
function isJunk(email) {
  const from = (email.from?.emailAddress?.address || '').toLowerCase();
  const name = (email.from?.emailAddress?.name || '').toLowerCase();
  const subject = (email.subject || '').toLowerCase();

  // Kolla avsändare
  if (JUNK_DOMAINS.some(d => from.includes(d) || name.includes(d))) {
    // Men behåll om det innehåller viktiga nyckelord
    if (IMPORTANT_KEYWORDS.some(k => subject.includes(k) || name.includes(k))) {
      return false;
    }
    return true;
  }

  // Kolla ämnesord
  if (JUNK_SUBJECTS.some(s => subject.includes(s))) {
    if (IMPORTANT_KEYWORDS.some(k => subject.includes(k))) {
      return false;
    }
    return true;
  }

  return false;
}

export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  // Kort cache — 5 minuter
  res.setHeader('Cache-Control', 's-maxage=300, stale-while-revalidate=60');

  if (req.method === 'OPTIONS') return res.status(200).end();

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

    const token = tokenResult.accessToken;
    const headers = { Authorization: `Bearer ${token}` };

    // Hämta mappnamn
    const foldersResp = await fetch(
      `https://graph.microsoft.com/v1.0/users/${USER_EMAIL}/mailFolders?$top=200`,
      { headers }
    );
    const foldersData = await foldersResp.json();
    const folderMap = {};
    for (const f of (foldersData.value || [])) {
      folderMap[f.id] = f.displayName;
    }

    // Hämta olästa mejl — senaste 200, paginerat
    const filter = encodeURIComponent("isRead eq false");
    const select = 'id,subject,from,receivedDateTime,bodyPreview,importance,parentFolderId,isRead';
    let url = `https://graph.microsoft.com/v1.0/users/${USER_EMAIL}/messages?$filter=${filter}&$orderby=receivedDateTime desc&$top=100&$select=${select}`;

    let allMessages = [];
    let pages = 0;

    while (url && pages < 3) {
      const resp = await fetch(url, { headers });
      if (!resp.ok) {
        const errText = await resp.text();
        return res.status(resp.status).json({ error: errText });
      }
      const data = await resp.json();
      allMessages = allMessages.concat(data.value || []);
      url = data['@odata.nextLink'] || null;
      pages++;
    }

    // Filtrera bort skräp och kategorisera
    const important = [];
    const skippedCount = { junk: 0, junkMail: 0 };

    for (const msg of allMessages) {
      const folderName = folderMap[msg.parentFolderId] || '';

      // Skippa mejl i skräppost/borttaget
      if (['Skräppost', 'Junk Email', 'Borttaget', 'Deleted Items', 'Drafts', 'Utkast', 'Sent Items', 'Skickat'].some(f =>
        folderName.toLowerCase().includes(f.toLowerCase())
      )) {
        skippedCount.junkMail++;
        continue;
      }

      // Filtrera automatiska/oviktiga mejl
      if (isJunk(msg)) {
        skippedCount.junk++;
        continue;
      }

      const cat = categorize(msg);
      const priority = getPriority(msg);
      const date = new Date(msg.receivedDateTime);

      important.push({
        id: msg.id,
        subject: msg.subject || '(Inget ämne)',
        from: msg.from?.emailAddress?.name || msg.from?.emailAddress?.address || 'Okänd',
        fromEmail: msg.from?.emailAddress?.address || '',
        date: date.toLocaleDateString('sv-SE', { day: 'numeric', month: 'short', year: 'numeric' }),
        dateISO: msg.receivedDateTime,
        preview: (msg.bodyPreview || '').substring(0, 200),
        folder: folderName || 'Inkorg',
        priority,
        category: cat.category,
        categoryLabel: cat.label,
        categoryIcon: cat.icon,
      });
    }

    // Sortera: hög prioritet först, sedan datum (nyast först)
    const priorityOrder = { high: 0, medium: 1, low: 2 };
    important.sort((a, b) => {
      if (priorityOrder[a.priority] !== priorityOrder[b.priority]) {
        return priorityOrder[a.priority] - priorityOrder[b.priority];
      }
      return new Date(b.dateISO) - new Date(a.dateISO);
    });

    return res.status(200).json({
      emails: important,
      total: allMessages.length,
      filtered: skippedCount.junk + skippedCount.junkMail,
      important: important.length,
      fetchedAt: new Date().toISOString(),
    });
  } catch (err) {
    return res.status(500).json({ error: err.message });
  }
}
