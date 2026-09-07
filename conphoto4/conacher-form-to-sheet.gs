/**
 * Conacher Photography — Form → "Website Data" sheet bridge
 * ============================================================
 * This script must live on the FORM (Form editor → ⋮ menu → Script editor),
 * not the spreadsheet — see the wiring steps below for why.
 *
 * What it does on every submission:
 *   1. Reads "What are you editing?" to figure out which of the 4 branches
 *      was used (New collection / New Carousel Group / Adding to Carousel
 *      Group / Adding to Genre).
 *   2. Pulls that branch's answers by QUESTION TITLE (not column position),
 *      so re-ordering questions in the Form later won't break this.
 *   3. Converts any uploaded files (Cover Photo, Photos in the
 *      collection/group, Upload Photos, overflow fields) from Drive file
 *      IDs into public, hotlink-able image URLs.
 *   4. Appends ONE normalized row to the "Website Data" tab, in the exact
 *      column shape every page's JS already expects:
 *      Timestamp | Page | Section | Name | Label | Cover Photo | Album Photos | Story
 *
 * "Adding to Carousel Group" isn't built out in the Form yet — that branch
 * below just logs and exits cleanly so a submission there never throws.
 * Mirror the ROUTE_NEW_CAROUSEL block once it exists.
 */

// ---- Sheet destination ----------------------------------------------------
const DATA_SHEET_NAME = 'Website Data';

// ---- Exact question titles, copied from the Form ---------------------------
const Q_ROUTER = 'What are you editing?';

const Q_COLLECTION_NAME     = 'Whats the name of the collection?';
const Q_COLLECTION_STORY    = 'A few words on the photo (location, setting, etc..)';
const Q_COLLECTION_COVER    = 'Cover Photo';           // shared title, disambiguated by branch below
const Q_COLLECTION_PHOTOS   = 'Photos in the collection';
const Q_COLLECTION_OVERFLOW = 'If more than 10 photos';

const Q_CAROUSEL_NAME   = 'Name';
const Q_CAROUSEL_COVER  = 'Cover Photo';
const Q_CAROUSEL_PHOTOS = 'Photos in the group';

const Q_GENRE_WHICH    = 'Which genre?';
const Q_GENRE_PHOTOS   = 'Upload Photos';
const Q_GENRE_OVERFLOW = 'If more than 10 photos';

// Router option values as they appear in the dropdown.
const ROUTE_NEW_COLLECTION = 'New collection';
const ROUTE_NEW_CAROUSEL   = 'New Carousel Group';
const ROUTE_ADD_CAROUSEL   = 'Adding to Carousel Group'; // not built out yet
const ROUTE_ADD_GENRE      = 'Adding to Genre';

// ---- Resolves the linked response spreadsheet, regardless of container ----
// SpreadsheetApp.getActiveSpreadsheet() ONLY works in a script bound to a
// Sheet. This script is bound to the Form, so that call would return
// nothing here — going through the Form's destination ID works no matter
// which container the script lives in.
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

    let row = null; // {page, section, name, label, cover, album, story}

    if (route === ROUTE_NEW_COLLECTION) {
      const coverIds = getFileIds(Q_COLLECTION_COVER);
      const photoIds = [...getFileIds(Q_COLLECTION_PHOTOS), ...getFileIds(Q_COLLECTION_OVERFLOW)];
      row = {
        page: 'portfolio',
        section: 'collection',
        name: getText(Q_COLLECTION_NAME),
        label: '',
        cover: filesToUrls(coverIds)[0] || '',
        album: filesToUrls(photoIds),
        story: getText(Q_COLLECTION_STORY)
      };

    } else if (route === ROUTE_NEW_CAROUSEL) {
      const coverIds = getFileIds(Q_CAROUSEL_COVER);
      const photoIds = getFileIds(Q_CAROUSEL_PHOTOS);
      row = {
        page: 'home',
        section: 'carousel',
        name: getText(Q_CAROUSEL_NAME),
        label: '',
        cover: filesToUrls(coverIds)[0] || '',
        album: filesToUrls(photoIds),
        story: ''
      };

    } else if (route === ROUTE_ADD_CAROUSEL) {
      // Section not built out in the Form yet — nothing to safely map.
      // Once it exists, mirror the ROUTE_NEW_CAROUSEL block above, matching
      // the existing group by Name so the site's "combine by Name" logic
      // picks it up automatically.
      Logger.log('Adding to Carousel Group submitted, but this branch is not built out yet — skipped.');
      return;

    } else if (route === ROUTE_ADD_GENRE) {
      const photoIds = [...getFileIds(Q_GENRE_PHOTOS), ...getFileIds(Q_GENRE_OVERFLOW)];
      row = {
        page: 'portfolio',
        section: 'genre',
        name: '',
        label: getText(Q_GENRE_WHICH),
        cover: '',
        album: filesToUrls(photoIds),
        story: ''
      };

    } else {
      Logger.log('Unrecognized "What are you editing?" value: "' + route + '" — check it matches ROUTE_* exactly (including capitalization).');
      return;
    }

    appendRow(timestamp, row);
    Logger.log('Row appended OK: ' + JSON.stringify(row));

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
