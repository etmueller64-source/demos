/**
 * Newmarket Church of the Nazarene — Event Photos: Form -> Website Data
 * ---------------------------------------------------------------------------
 * The website's actual event LIST (dates, titles, upcoming vs past) comes
 * straight from a public Google Calendar — this script has nothing to do
 * with that. All this handles is PHOTOS: someone fills in a simple form
 * after a "FEATURED:" calendar event happens, and this script turns that
 * into a clean row the website can read, matched back to the calendar
 * event by title.
 *
 * SETUP (once):
 *   1. Extensions -> Apps Script, paste this in, Save.
 *   2. Run  installTrigger  once. Approve the permissions prompt.
 *      (It asks for Drive access — that's the file-sharing step in publishFile.)
 *   3. Run  rebuildAll  once to backfill any existing responses.
 *
 * After that it runs itself on every submission.
 */

// ============================================================ configuration

const RESPONSES_SHEET = 'Form Responses 1';
const OUTPUT_SHEET    = 'Website Data';
const OUTPUT_HEADERS  = ['Timestamp', 'Event Title', 'Cover Photo', 'Gallery Photos'];

const IMAGE_WIDTH       = 2000;   // the =w2000 on the end of each image URL
const MAKE_FILES_PUBLIC = true;   // set uploads to "anyone with the link can view"

/**
 * Question titles, spelled EXACTLY as they appear in the form.
 * >>> This MUST match the calendar event's title, minus the "FEATURED:"
 * prefix <<< — that's how the website links a photo submission back to the
 * right calendar event. A mismatch here just means that event's photos
 * won't show up; nothing breaks, it just silently has no gallery yet.
 */
const Q = {
  title:   'Event Title',
  cover:   'Cover Photo',
  gallery: 'Gallery Photos'
};

// ============================================================ entry points

/** Run once, by hand, to start listening for submissions. */
function installTrigger() {
  const ss = SpreadsheetApp.getActive();
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'onFormSubmit')
    .forEach(t => ScriptApp.deleteTrigger(t));          // no duplicate triggers
  ScriptApp.newTrigger('onFormSubmit').forSpreadsheet(ss).onFormSubmit().create();
  SpreadsheetApp.getActive().toast('Trigger installed.');
}

/** Fires automatically on every form submission. */
function onFormSubmit() {
  rebuildAll();   // cheap at this volume, and it can never drift out of sync
}

/** Rebuilds the Website Data tab from every response. Safe to run any time. */
function rebuildAll() {
  const ss  = SpreadsheetApp.getActive();
  const src = ss.getSheetByName(RESPONSES_SHEET);
  if (!src) throw new Error('No sheet named "' + RESPONSES_SHEET + '"');

  const out = ss.getSheetByName(OUTPUT_SHEET) || ss.insertSheet(OUTPUT_SHEET);
  const values = src.getDataRange().getValues();
  if (values.length < 2) return;

  const headers = values[0].map(h => String(h).trim());
  const rows = [];

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    if (row.every(c => String(c).trim() === '')) continue;   // skip blank rows

    const get = title => {
      if (!title) return '';
      const key = title.trim().toLowerCase();
      for (let c = 0; c < headers.length; c++) {
        if (headers[c].toLowerCase() === key) {
          const v = String(row[c] ?? '').trim();
          if (v) return v;
        }
      }
      return '';
    };

    const title = get(Q.title);
    if (!title) continue; // no way to match this to a calendar event - skip it

    const coverUrls = toDirectUrls(get(Q.cover));
    const galleryUrls = toDirectUrls(get(Q.gallery));
    if (!coverUrls.length && !galleryUrls.length) continue; // nothing to show

    rows.push([
      timestamp(),
      title,
      coverUrls.join(', '),
      galleryUrls.join(', ')
    ]);
  }

  out.clearContents();
  out.getRange(1, 1, 1, OUTPUT_HEADERS.length).setValues([OUTPUT_HEADERS]);
  if (rows.length) out.getRange(2, 1, rows.length, OUTPUT_HEADERS.length).setValues(rows);

  out.getRange(1, 1, 1, OUTPUT_HEADERS.length).setFontWeight('bold');
  out.setFrozenRows(1);
}

/**
 * ISO-style timestamp. Written with a "T" so every browser parses it —
 * Safari does not reliably accept "2026-07-28 19:37:56" with a space.
 */
function timestamp() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss");
}

// ============================================================ Drive links

/**
 * Form uploads arrive as  https://drive.google.com/open?id=FILEID
 * A browser can't render that — it's a Drive viewer page, not an image.
 * This rewrites each one to the direct form and, importantly, makes the file
 * publicly viewable. Without that step the photos load for you (you own them)
 * and show as broken for everyone else — the classic Drive hotlinking trap.
 */
function toDirectUrls(cell) {
  const raw = String(cell || '').trim();
  if (!raw) return [];

  return raw.split(',')
    .map(s => s.trim())
    .filter(Boolean)
    .map(url => {
      const id = extractFileId(url);
      if (!id) return url;                   // not a Drive link — pass through as-is
      if (MAKE_FILES_PUBLIC) publishFile(id);
      return 'https://lh3.googleusercontent.com/d/' + id + '=w' + IMAGE_WIDTH;
    })
    .filter(Boolean);
}

/** Pulls the file ID out of any of the Drive URL shapes. */
function extractFileId(url) {
  const s = String(url);
  let m = s.match(/[?&]id=([a-zA-Z0-9_-]+)/);
  if (m) return m[1];
  m = s.match(/\/file\/d\/([a-zA-Z0-9_-]+)/);
  if (m) return m[1];
  m = s.match(/lh3\.googleusercontent\.com\/d\/([a-zA-Z0-9_-]+)/);
  if (m) return m[1];
  return null;
}

/** Sets one file to "anyone with the link can view". Never throws. */
function publishFile(fileId) {
  try {
    DriveApp.getFileById(fileId)
      .setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (err) {
    console.warn('Could not publish file ' + fileId + ': ' + err);
  }
}
