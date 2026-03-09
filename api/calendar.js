import { google } from 'googleapis';
import { ConfidentialClientApplication } from '@azure/msal-node';

const USER_EMAIL = 'marzena@marzenalangsdale.com';

const YOBEDOO_ICS_URL = 'https://outlook.office365.com/owa/calendar/7784fa426c2145ee866c3914d19bc670@yobedoo.com/2fa37a7a158349dca13ff5b5fd60a8307320543417251081616/calendar.ics';

const GOOGLE_CALENDARS = {
  arbete: 'rd77r3g1nq0m48g3vntp3u1uhk@group.calendar.google.com',
  privat: 'aout8htgc2m1k9mjsiae3vi7mk@group.calendar.google.com',
  familj: 'g72eeu1f0hk0gbe7be9th52h7k@group.calendar.google.com',
  evolan: 'rhk7ggl5tvgvp8qrf3k3ok3o58@group.calendar.google.com',
};

export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  res.setHeader('Cache-Control', 's-maxage=300, stale-while-revalidate=60');

  if (req.method === 'OPTIONS') return res.status(200).end();

  // Use start of today in Europe/Stockholm timezone (handles DST)
  const now = new Date();
  const stockholmDate = now.toLocaleDateString('sv-SE', { timeZone: 'Europe/Stockholm' });
  const offsetPart = new Intl.DateTimeFormat('en', {
    timeZone: 'Europe/Stockholm', timeZoneName: 'longOffset'
  }).formatToParts(now).find(p => p.type === 'timeZoneName').value;
  const offset = offsetPart.replace('GMT', '') || '+00:00';
  const startOfToday = new Date(stockholmDate + 'T00:00:00' + offset);
  const future = new Date(startOfToday.getTime() + 7 * 24 * 60 * 60 * 1000);
  const events = [];
  const errors = [];

  // 1. Google Calendars — fetch all in parallel
  try {
    const credentials = JSON.parse(process.env.GOOGLE_CALENDAR_KEY);
    const auth = new google.auth.GoogleAuth({
      credentials,
      scopes: ['https://www.googleapis.com/auth/calendar.readonly'],
    });
    const calendar = google.calendar({ version: 'v3', auth });

    const googleResults = await Promise.allSettled(
      Object.entries(GOOGLE_CALENDARS).map(async ([name, calId]) => {
        const result = await calendar.events.list({
          calendarId: calId,
          timeMin: startOfToday.toISOString(),
          timeMax: future.toISOString(),
          singleEvents: true,
          orderBy: 'startTime',
          maxResults: 30,
        });
        return { name, items: result.data.items || [] };
      })
    );

    for (const result of googleResults) {
      if (result.status === 'fulfilled') {
        for (const e of result.value.items) {
          events.push({
            title: e.summary || '(Ingen rubrik)',
            start: e.start.dateTime || e.start.date,
            end: e.end.dateTime || e.end.date,
            allDay: !e.start.dateTime,
            calendar: result.value.name,
            location: e.location || null,
          });
        }
      } else {
        errors.push(`Google: ${result.reason.message}`);
      }
    }
  } catch (err) {
    errors.push(`Google auth: ${err.message}`);
  }

  // 2. Exchange Calendar
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

    const url = `https://graph.microsoft.com/v1.0/users/${USER_EMAIL}/calendarView?startDateTime=${startOfToday.toISOString()}&endDateTime=${future.toISOString()}&$select=subject,start,end,isAllDay,location&$orderby=start/dateTime&$top=50`;
    const response = await fetch(url, {
      headers: {
        Authorization: `Bearer ${tokenResult.accessToken}`,
        'Prefer': 'outlook.timezone="Europe/Stockholm"',
      },
    });

    if (response.ok) {
      const data = await response.json();
      for (const e of (data.value || [])) {
        // With Europe/Stockholm timezone header, dateTime is in local time
        // Parse as local time by appending Stockholm offset
        const startDt = e.start.dateTime.replace(/\.0+$/, '');
        const endDt = e.end.dateTime.replace(/\.0+$/, '');
        events.push({
          title: e.subject || '(Ingen rubrik)',
          start: startDt,
          end: endDt,
          allDay: e.isAllDay,
          calendar: 'exchange',
          location: e.location?.displayName || null,
        });
      }
    } else {
      errors.push(`Exchange: ${response.status}`);
    }
  } catch (err) {
    errors.push(`Exchange: ${err.message}`);
  }

  // 3. Yobedoo Calendar (public ICS feed)
  try {
    const icsResp = await fetch(YOBEDOO_ICS_URL, {
      headers: { 'User-Agent': 'MLC-Calendar/1.0' },
    });
    if (icsResp.ok) {
      const icsText = await icsResp.text();
      // Parse VEVENT blocks from ICS
      const vevents = icsText.split('BEGIN:VEVENT');
      for (let i = 1; i < vevents.length; i++) {
        const block = vevents[i].split('END:VEVENT')[0];
        const get = (key) => {
          const m = block.match(new RegExp('^' + key + '[^:]*:(.+)', 'm'));
          return m ? m[1].trim() : null;
        };

        const summary = get('SUMMARY') || '(Ingen rubrik)';
        const dtstart = get('DTSTART');
        const dtend = get('DTEND');
        if (!dtstart) continue;

        // Parse ICS date formats: 20260313T140000 or with TZID
        const parseIcsDate = (val) => {
          // Strip any trailing \r
          const clean = val.replace(/\r/g, '');
          // Format: YYYYMMDD (all-day)
          if (/^\d{8}$/.test(clean)) {
            return { date: clean.slice(0,4)+'-'+clean.slice(4,6)+'-'+clean.slice(6,8), allDay: true };
          }
          // Format: YYYYMMDDTHHmmss or YYYYMMDDTHHmmssZ
          const m = clean.match(/^(\d{4})(\d{2})(\d{2})T(\d{2})(\d{2})(\d{2})Z?$/);
          if (m) {
            const isUtc = clean.endsWith('Z');
            const iso = `${m[1]}-${m[2]}-${m[3]}T${m[4]}:${m[5]}:${m[6]}`;
            if (isUtc) {
              // Convert UTC to Stockholm local time
              const d = new Date(iso + 'Z');
              return { date: d.toISOString(), allDay: false };
            }
            // Already local time (TZID specified in ICS)
            return { date: iso, allDay: false };
          }
          return null;
        };

        const startParsed = parseIcsDate(dtstart);
        if (!startParsed) continue;

        // Filter to our date range
        const evDate = new Date(startParsed.date);
        if (evDate < startOfToday || evDate > future) continue;

        const endParsed = dtend ? parseIcsDate(dtend) : startParsed;

        events.push({
          title: summary,
          start: startParsed.date,
          end: endParsed ? endParsed.date : startParsed.date,
          allDay: startParsed.allDay,
          calendar: 'yobedoo',
          location: get('LOCATION') || null,
        });
      }
    } else {
      errors.push(`Yobedoo ICS: ${icsResp.status}`);
    }
  } catch (err) {
    errors.push(`Yobedoo ICS: ${err.message}`);
  }

  // 4. Deduplicate — keep primary source when title+start match within 2 min
  // Priority: Google > Yobedoo > Exchange
  const deduped = [];
  const primaryEvents = events.filter(e => e.calendar !== 'exchange' && e.calendar !== 'yobedoo');
  const yobedooEvents = events.filter(e => e.calendar === 'yobedoo');
  const exchangeEvents = events.filter(e => e.calendar === 'exchange');

  deduped.push(...primaryEvents);
  deduped.push(...yobedooEvents);

  for (const ex of exchangeEvents) {
    const exStart = new Date(ex.start).getTime();
    const isDuplicate = [...primaryEvents, ...yobedooEvents].some(g => {
      const gStart = new Date(g.start).getTime();
      return g.title.toLowerCase() === ex.title.toLowerCase()
        && Math.abs(gStart - exStart) < 2 * 60 * 1000;
    });
    if (!isDuplicate) {
      deduped.push(ex);
    }
  }

  // 4. Sort by start time
  deduped.sort((a, b) => new Date(a.start) - new Date(b.start));

  return res.status(200).json({ events: deduped, errors });
}
