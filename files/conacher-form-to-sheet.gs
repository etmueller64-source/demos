/**
 * Conacher Photography — Form → "Website Data" sheet bridge
 * ============================================================
 * This script must live on the FORM (Form editor → ⋮ menu → Script editor),
 * not the spreadsheet — see the wiring steps sent alongside this file for
 * why, and for exactly how to build the matching Form sections.
 *
 * Six branches, driven by "What are you editing?":
 *   1. Collections and Events        — create a new collection/event
 *   2. Adding to Collections and Events — add photos to an existing one
 *   3. New Carousel Group            — create a new Home carousel group
 *   4. Adding to Carousel Group      — add photos to an existing one
 *   5. New Genre                     — create a new Portfolio genre,
 *                                       optionally featuring it on Home
 *   6. Adding to Genre               — add photos to an existing genre
 *
 * Every branch is pulled by QUESTION TITLE (not column position), so
 * reordering questions in the Form later won't break this — but the
 * TITLES MUST MATCH EXACTLY (see the Q_* constants below) between the
 * Form and this script.
 *
 * Column contract written to the "Website Data" tab (unchanged from
 * before — every page's JS already expects this shape):
 *   Timestamp | Page | Section | Name | Label | Cover Photo | Album Photos | Story
 *
 * Featuring a new genre on Home doesn't need its own column — it just
 * writes a SECOND row (Section: "featured") alongside the genre row,
 * reusing Home's existing Featured-tile mechanism as-is.
 */

// ---- Sheet destination ----------------------------------------------------
const DATA_SHEET_NAME = 'Website Data';

// ---- Router ----------------------------------------------------------------
const Q_ROUTER = 'What are you editing?';

const ROUTE_NEW_COLLECTION = 'Collections and Events';
const ROUTE_ADD_COLLECTION = 'Adding to Collections and Events';
const ROUTE_NEW_CAROUSEL   = 'New Carousel Group';
const ROUTE_ADD_CAROUSEL   = 'Adding to Carousel Group';
const ROUTE_NEW_GENRE      = 'New Genre';
const ROUTE_ADD_GENRE      = 'Adding to Genre';

// ---- 1. Collections and Events (new) ---------------------------------------
const Q_COLLECTION_NAME     = 'Whats the name of the collection?';
const Q_COLLECTION_STORY    = 'A few words on the photo (location, setting, etc..)';
const Q_COLLECTION_COVER    = 'Cover Photo';
const Q_COLLECTION_PHOTOS   = 'Photos in the collection';
const Q_COLLECTION_OVERFLOW = 'If more than 10 photos';

// ---- 2. Adding to Collections and Events (existing) -------------------------
const Q_ADD_COLLECTION_WHICH    = 'Which collection or event?'; // dropdown of existing names
const Q_ADD_COLLECTION_PHOTOS   = 'Photos to add';
const Q_ADD_COLLECTION_OVERFLOW = 'If more than 10 photos';

// ---- 3. New Carousel Group ---------------------------------------------------
const Q_CAROUSEL_NAME   = 'Name';
const Q_CAROUSEL_COVER  = 'Cover Photo';
const Q_CAROUSEL_PHOTOS = 'Photos in the group';

// ---- 4. Adding to Carousel Group (existing) ----------------------------------
const Q_ADD_CAROUSEL_WHICH  = 'Which carousel group?'; // dropdown of existing group names
const Q_ADD_CAROUSEL_PHOTOS = 'Photos to add';

// ---- 5. New Genre -------------------------------------------------------------
const Q_NEW_GENRE_NAME     = 'Whats the genre called?';
const Q_NEW_GENRE_FEATURED = 'Feature this on the home page?'; // Yes/No
const Q_NEW_GENRE_PHOTOS   = 'Upload Photos';
const Q_NEW_GENRE_OVERFLOW = 'If more than 10 photos';
const FEATURED_YES = 'Yes'; // must match the Form's exact option text

// ---- 6. Adding to Genre (existing) ---------------------------------------------
const Q_ADD_GENRE_WHICH    = 'Which genre?'; // dropdown of existing genres
const Q_ADD_GENRE_PHOTOS   = 'Upload Photos';
const Q_ADD_GENRE_OVERFLOW = 'If more than 10 photos';

// ---- Resolves the linked response spreadsheet, regardless of container ----
// SpreadsheetApp.getActiveSpreadsheet() ONLY works in a script bound to a
// Sheet. This script is bound to the Form, so that call returns nothing
// here — going through the Form's destination ID works no matter which
// container the script lives in.
function getDataSpreadsheet() {
  const form = FormApp.getActiveForm();
  if (!form) {
    throw new Error('This script isn\'t bound to a Form. Open it from the FORM editor (⋮ menu → Script editor), not from the Sheet.');
  }
  const destId = form.getDestinationId();
  if (!destId) {
    throw new Error('This Form has no linked response Spreadsheet yet. In the Form, go to Responses → click the green Sheets icon → link/create one, then try again.');
  }
  return SpreadsheetApp.openById(destId);
}

/**
 * Manual test helper — run THIS from the function picker (▶ dropdown at
 * the top of the editor), not onFormSubmit itself, to test without
 * submitting a real response every time. It grabs your most recent real
 * response and runs it through the exact same code path.
 *
 * (Running onFormSubmit directly always throws "Cannot read properties of
 * undefined (reading 'response')" — Apps Script doesn't invent a fake
 * event object; only a real submission via the installed trigger does.)
 */
function testWithLatestResponse() {
  const form = FormApp.getActiveForm();
  if (!form) { Logger.log('Not bound to a Form — open this script from the Form editor, not the Sheet.'); return; }
  const responses = form.getResponses();
  if (!responses.length) { Logger.log('No responses on this form yet — submit one first.'); return; }
  const latest = responses[responses.length - 1];
  onFormSubmit({ response: latest });
}

function onFormSubmit(e) {
  try {
    const itemResponses = e.response.getItemResponses();
    const byTitle = {};
    itemResponses.forEach(ir => { byTitle[ir.getItem().getTitle()] = ir; });

    const getText = title => (byTitle[title] ? String(byTitle[title].getResponse() || '').trim() : '');
    const getFileIds = title => {
      if (!byTitle[title]) return [];
      const val = byTitle[title].getResponse();
      // File-upload questions return an array of Drive file IDs; guard for
      // the single-ID/no-answer cases too, just in case.
      if (Array.isArray(val)) return val.filter(Boolean);
      return val ? [val] : [];
    };
    const filesToUrls = ids => ids.map(driveFileToPublicUrl).filter(Boolean);

    const route = getText(Q_ROUTER);
    const timestamp = e.response.getTimestamp();

    Logger.log('Routing on: "' + route + '"');

    // Each branch pushes one or more {page, section, name, label, cover,
    // album, story} rows onto this list — usually just one, except New
    // Genre with Featured=Yes, which pushes two.
    const rows = [];

    if (route === ROUTE_NEW_COLLECTION) {
      const coverIds = getFileIds(Q_COLLECTION_COVER);
      const photoIds = [...getFileIds(Q_COLLECTION_PHOTOS), ...getFileIds(Q_COLLECTION_OVERFLOW)];
      rows.push({
        page: 'portfolio', section: 'collection',
        name: getText(Q_COLLECTION_NAME), label: '',
        cover: filesToUrls(coverIds)[0] || '',
        album: filesToUrls(photoIds),
        story: getText(Q_COLLECTION_STORY)
      });

    } else if (route === ROUTE_ADD_COLLECTION) {
      const photoIds = [...getFileIds(Q_ADD_COLLECTION_PHOTOS), ...getFileIds(Q_ADD_COLLECTION_OVERFLOW)];
      rows.push({
        page: 'portfolio', section: 'collection',
        name: getText(Q_ADD_COLLECTION_WHICH), label: '',
        cover: '', // leave existing cover untouched - this row only adds photos
        album: filesToUrls(photoIds),
        story: ''
      });

    } else if (route === ROUTE_NEW_CAROUSEL) {
      const coverIds = getFileIds(Q_CAROUSEL_COVER);
      const photoIds = getFileIds(Q_CAROUSEL_PHOTOS);
      rows.push({
        page: 'home', section: 'carousel',
        name: getText(Q_CAROUSEL_NAME), label: '',
        cover: filesToUrls(coverIds)[0] || '',
        album: filesToUrls(photoIds),
        story: ''
      });

    } else if (route === ROUTE_ADD_CAROUSEL) {
      const photoIds = getFileIds(Q_ADD_CAROUSEL_PHOTOS);
      rows.push({
        page: 'home', section: 'carousel',
        name: getText(Q_ADD_CAROUSEL_WHICH), label: '',
        cover: '', // leave existing cover untouched - this row only adds photos
        album: filesToUrls(photoIds),
        story: ''
      });

    } else if (route === ROUTE_NEW_GENRE) {
      const genreName = getText(Q_NEW_GENRE_NAME);
      const photoIds = [...getFileIds(Q_NEW_GENRE_PHOTOS), ...getFileIds(Q_NEW_GENRE_OVERFLOW)];
      const photoUrls = filesToUrls(photoIds);
      rows.push({
        page: 'portfolio', section: 'genre',
        name: '', label: genreName,
        cover: '', album: photoUrls, story: ''
      });
      const wantsFeatured = getText(Q_NEW_GENRE_FEATURED) === FEATURED_YES;
      if (wantsFeatured) {
        rows.push({
          page: 'home', section: 'featured',
          name: genreName, label: genreName,
          cover: photoUrls[0] || '', // Home's Featured tile needs a Cover Photo to display at all
          album: [], story: ''
        });
      }

    } else if (route === ROUTE_ADD_GENRE) {
      const photoIds = [...getFileIds(Q_ADD_GENRE_PHOTOS), ...getFileIds(Q_ADD_GENRE_OVERFLOW)];
      rows.push({
        page: 'portfolio', section: 'genre',
        name: '', label: getText(Q_ADD_GENRE_WHICH),
        cover: '', album: filesToUrls(photoIds), story: ''
      });

    } else {
      Logger.log('Unrecognized "What are you editing?" value: "' + route + '" — check it matches ROUTE_* exactly (including capitalization).');
      return;
    }

    rows.forEach(row => appendRow(timestamp, row));
    Logger.log('Appended ' + rows.length + ' row(s) OK: ' + JSON.stringify(rows));

  } catch (err) {
    Logger.log('onFormSubmit error: ' + err.message + '\n' + err.stack);
    // Re-throw so a failure shows up as a red "failure" in Apps Script's
    // trigger history instead of silently vanishing.
    throw err;
  }
}

// ---- Writes one normalized row into the Website Data tab -------------------
function appendRow(timestamp, row) {
  const ss = getDataSpreadsheet();
  const sheet = ss.getSheetByName(DATA_SHEET_NAME);
  if (!sheet) throw new Error('No tab named "' + DATA_SHEET_NAME + '" found in the response spreadsheet — run runSetup() once to create it.');

  // Column order must match the header row already on the sheet:
  // Timestamp | Page | Section | Name | Label | Cover Photo | Album Photos | Story
  sheet.appendRow([
    timestamp,
    row.page,
    row.section,
    row.name,
    row.label,
    row.cover,
    row.album.join(', '),
    row.story
  ]);
}

// ---- Turns a Drive file ID into a public, hotlink-able image URL -----------
function driveFileToPublicUrl(fileId) {
  try {
    const file = DriveApp.getFileById(fileId);
    // Make sure the site (and anyone visiting it) can actually load the
    // image — Form-uploaded files default to private, visible only to you.
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    // This lh3.googleusercontent.com form is far more reliable for <img src>
    // hotlinking than the old drive.google.com/uc?export=view link, which
    // Google increasingly rate-limits/blocks for direct embedding.
    return 'https://lh3.googleusercontent.com/d/' + fileId;
  } catch (err) {
    Logger.log('Could not process file ' + fileId + ': ' + err.message);
    return '';
  }
}

/**
 * One-time setup helper — run this once manually (pick runSetup in the
 * function dropdown, then ▶ Run) to create the Website Data tab with the
 * correct header row if it doesn't exist yet. Safe to run again later —
 * it won't duplicate the tab or touch existing rows.
 */
function runSetup() {
  const ss = getDataSpreadsheet();
  let sheet = ss.getSheetByName(DATA_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(DATA_SHEET_NAME);
    sheet.appendRow(['Timestamp', 'Page', 'Section', 'Name', 'Label', 'Cover Photo', 'Album Photos', 'Story']);
    sheet.setFrozenRows(1);
    Logger.log('Created "' + DATA_SHEET_NAME + '" tab with header row.');
  } else {
    Logger.log('"' + DATA_SHEET_NAME + '" tab already exists — nothing to do.');
  }
}
