/**
 * Conacher Photography — Form → "Website Data" sheet bridge
 * ============================================================
 * Install this as an installable "On form submit" trigger on the FORM
 * itself (Form editor → ⋮ → Script editor), not the spreadsheet. It needs
 * to be an installable trigger (Triggers → + Add Trigger → onFormSubmit →
 * From form → On form submit) rather than the simple built-in trigger,
 * because setting Drive file permissions requires full authorization.
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
 * "Adding to Carousel Group" isn't built out in the Form yet (per Eric) —
 * the branch below is a stub that just logs and exits cleanly so a
 * submission there never throws. Fill in ADD_TO_CAROUSEL_TITLES once that
 * section exists.
 */

// ---- Sheet destination ----------------------------------------------------
const DATA_SHEET_NAME = 'Website Data';

// ---- Exact question titles, copied from the Form ---------------------------
const Q_ROUTER = 'What are you editing?';

const Q_COLLECTION_NAME   = 'Whats the name of the collection?';
const Q_COLLECTION_STORY  = 'A few words on the photo (location, setting, etc..)';
const Q_COLLECTION_COVER  = 'Cover Photo';           // shared title, disambiguated by section below
const Q_COLLECTION_PHOTOS = 'Photos in the collection';
const Q_COLLECTION_OVERFLOW = 'If more than 10 photos';

const Q_CAROUSEL_NAME   = 'Name';
const Q_CAROUSEL_COVER  = 'Cover Photo';
const Q_CAROUSEL_PHOTOS = 'Photos in the group';

const Q_GENRE_WHICH     = 'Which genre?';
const Q_GENRE_PHOTOS    = 'Upload Photos';
const Q_GENRE_OVERFLOW  = 'If more than 10 photos';

// Router option values as they appear in the dropdown.
const ROUTE_NEW_COLLECTION = 'New collection';
const ROUTE_NEW_CAROUSEL   = 'New Carousel Group';
const ROUTE_ADD_CAROUSEL   = 'Adding to Carousel Group'; // not built out yet
const ROUTE_ADD_GENRE      = 'Adding to Genre';

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
      Logger.log('Unrecognized "What are you editing?" value: ' + route);
      return;
    }

    appendRow(timestamp, row);

  } catch (err) {
    Logger.log('onFormSubmit error: ' + err.message + '\n' + err.stack);
    // Re-throw so a failure shows up as a red "failure" in Apps Script's
    // trigger history instead of silently vanishing.
    throw err;
  }
}

// ---- Writes one normalized row into the Website Data tab -------------------
function appendRow(timestamp, row) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(DATA_SHEET_NAME);
  if (!sheet) throw new Error('No sheet named "' + DATA_SHEET_NAME + '" found — create it first with the header row.');

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
 * One-time setup helper — run this once manually (▶ in the Script editor,
 * with runSetup selected) to create the Website Data tab with the correct
 * header row if it doesn't already exist. Safe to run again later; it
 * won't duplicate the tab or touch existing rows.
 */
function runSetup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(DATA_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(DATA_SHEET_NAME);
    sheet.appendRow(['Timestamp', 'Page', 'Section', 'Name', 'Label', 'Cover Photo', 'Album Photos', 'Story']);
    sheet.setFrozenRows(1);
  }
}
