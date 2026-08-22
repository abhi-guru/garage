/**
 * gdrive-storage.js — Per-user Google Drive storage layer
 *
 * Replaces Apps Script API calls for users who sign in with Google and
 * choose "Create my garage". Each user's data lives in their own Google Drive:
 *
 *   My Drive/
 *     └── Garage Organizer/
 *           ├── garage_db  (Google Sheet — boxes, items, photos, config)
 *           └── Photos/
 *                 └── BOX_01/ … (image files)
 *
 * Public surface: GDriveDB.* — mirrors the Apps Script action names so
 * index.html can route to either backend with zero UI changes.
 *
 * OAuth scopes required:
 *   https://www.googleapis.com/auth/spreadsheets
 *   https://www.googleapis.com/auth/drive.file
 */

var GDriveDB = (function () {

  /* ── internals ──────────────────────────────────────────────── */
  var _token       = null;
  var _sheetId     = null;
  var _folderId    = null;
  var _adminEmail  = null;
  var _sheetMeta   = null;   // cached sheet ID → numeric sheetId map

  var SHEETS = 'https://sheets.googleapis.com/v4/spreadsheets';
  var DRIVE  = 'https://www.googleapis.com/drive/v3/files';
  var UPLOAD = 'https://www.googleapis.com/upload/drive/v3/files';

  /* Column schemas for each tab (row 1 = header) */
  var SCHEMA = {
    Boxes:      ['name','color','location','tags','coverImage','createdAt'],
    Items:      ['id','box','name','desc','category','qty','dateAdded','lentTo','dueDate','lentContact','lentNotes'],
    Photos:     ['id','box','itemId','itemName','url','driveFileId','dateAdded'],
    GarageZone: ['id','name','desc','category','location','qty','dateAdded'],
    Config:     ['key','value']
  };

  function auth() { return { 'Authorization': 'Bearer ' + _token }; }
  function json() { return { 'Authorization': 'Bearer ' + _token, 'Content-Type': 'application/json' }; }

  /* ── low-level Sheets helpers ───────────────────────────────── */
  async function sheetGet(range) {
    var r = await fetch(SHEETS + '/' + _sheetId + '/values/' + encodeURIComponent(range), { headers: auth() });
    var d = await r.json();
    if (!r.ok) throw new Error(d.error && d.error.message || 'Sheets read failed');
    return d.values || [];
  }

  async function sheetAppend(tab, row) {
    var r = await fetch(
      SHEETS + '/' + _sheetId + '/values/' + encodeURIComponent(tab + '!A1') +
      ':append?valueInputOption=RAW&insertDataOption=INSERT_ROWS',
      { method: 'POST', headers: json(), body: JSON.stringify({ values: [row] }) }
    );
    var d = await r.json();
    if (!r.ok) throw new Error(d.error && d.error.message || 'Sheets append failed');
    return d;
  }

  async function sheetUpdate(range, values) {
    var r = await fetch(
      SHEETS + '/' + _sheetId + '/values/' + encodeURIComponent(range) + '?valueInputOption=RAW',
      { method: 'PUT', headers: json(), body: JSON.stringify({ values: values }) }
    );
    var d = await r.json();
    if (!r.ok) throw new Error(d.error && d.error.message || 'Sheets update failed');
    return d;
  }

  async function sheetDeleteRow(tab, rowIndex /* 0-based, excluding header */) {
    var numericId = await _getNumericSheetId(tab);
    var r = await fetch(SHEETS + '/' + _sheetId + ':batchUpdate', {
      method: 'POST', headers: json(),
      body: JSON.stringify({ requests: [{ deleteDimension: {
        range: { sheetId: numericId, dimension: 'ROWS',
                 startIndex: rowIndex + 1,   // +1 skips header
                 endIndex:   rowIndex + 2 }
      }}]})
    });
    var d = await r.json();
    if (!r.ok) throw new Error(d.error && d.error.message || 'Sheets delete failed');
    return d;
  }

  async function _getNumericSheetId(tabName) {
    if (!_sheetMeta) {
      var r  = await fetch(SHEETS + '/' + _sheetId + '?fields=sheets.properties', { headers: auth() });
      var d  = await r.json();
      _sheetMeta = {};
      (d.sheets || []).forEach(function (s) {
        _sheetMeta[s.properties.title] = s.properties.sheetId;
      });
    }
    return _sheetMeta[tabName];
  }

  /* ── Drive helpers ──────────────────────────────────────────── */
  async function driveFind(q) {
    var r = await fetch(DRIVE + '?q=' + encodeURIComponent(q) + '&fields=files(id,name)', { headers: auth() });
    var d = await r.json();
    if (!r.ok) throw new Error(d.error && d.error.message || 'Drive list failed');
    return d.files || [];
  }

  async function driveCreate(meta) {
    var r = await fetch(DRIVE, { method: 'POST', headers: json(), body: JSON.stringify(meta) });
    var d = await r.json();
    if (!r.ok) throw new Error(d.error && d.error.message || 'Drive create failed');
    return d;
  }

  async function drivePatch(fileId, meta) {
    var r = await fetch(DRIVE + '/' + fileId + '?fields=id', { method: 'PATCH', headers: json(), body: JSON.stringify(meta) });
    var d = await r.json();
    if (!r.ok) throw new Error(d.error && d.error.message || 'Drive patch failed');
    return d;
  }

  async function drivePermit(fileId) {
    /* make file publicly readable so <img> tags can load it */
    await fetch('https://www.googleapis.com/drive/v3/files/' + fileId + '/permissions', {
      method: 'POST', headers: json(),
      body: JSON.stringify({ role: 'reader', type: 'anyone' })
    });
  }

  /* ── folder / spreadsheet bootstrap ────────────────────────── */
  async function _findOrCreateFolder() {
    var q = "name='Garage Organizer' and mimeType='application/vnd.google-apps.folder' and trashed=false";
    var files = await driveFind(q);
    if (files.length) return files[0].id;
    var f = await driveCreate({ name: 'Garage Organizer', mimeType: 'application/vnd.google-apps.folder' });
    return f.id;
  }

  async function _findOrCreateSheet(folderId) {
    var q = "name='garage_db' and mimeType='application/vnd.google-apps.spreadsheet'" +
            " and '" + folderId + "' in parents and trashed=false";
    var files = await driveFind(q);
    if (files.length) return files[0].id;
    return _createSheet(folderId);
  }

  async function _createSheet(folderId) {
    /* 1. Create spreadsheet with all tabs */
    var r = await fetch(SHEETS, {
      method: 'POST', headers: json(),
      body: JSON.stringify({
        properties: { title: 'garage_db' },
        sheets: Object.keys(SCHEMA).map(function (name, i) {
          return { properties: { title: name, index: i } };
        })
      })
    });
    var ss = await r.json();
    var ssId = ss.spreadsheetId;

    /* 2. Move into Garage Organizer folder */
    await fetch(DRIVE + '/' + ssId +
      '?addParents=' + folderId + '&removeParents=root&fields=id',
      { method: 'PATCH', headers: auth() }
    );

    /* 3. Write headers */
    var data = Object.entries(SCHEMA).map(function (e) {
      return { range: e[0] + '!A1', values: [e[1]] };
    });
    await fetch(SHEETS + '/' + ssId + '/values:batchUpdate', {
      method: 'POST', headers: json(),
      body: JSON.stringify({ valueInputOption: 'RAW', data: data })
    });

    return ssId;
  }

  /* ── sub-folder for box photos ──────────────────────────────── */
  async function _photoFolder(boxName) {
    var safe = boxName.replace(/[^a-zA-Z0-9_-]/g, '_');
    var q = "name='Photos_" + safe + "' and '" + _folderId + "' in parents" +
            " and mimeType='application/vnd.google-apps.folder' and trashed=false";
    var files = await driveFind(q);
    if (files.length) return files[0].id;
    var f = await driveCreate({
      name: 'Photos_' + safe,
      mimeType: 'application/vnd.google-apps.folder',
      parents: [_folderId]
    });
    return f.id;
  }

  /* ── base64 → Blob ──────────────────────────────────────────── */
  function _b64Blob(b64, mime) {
    var clean = b64.replace(/^data:[^;]+;base64,/, '');
    var bytes = atob(clean);
    var ab = new ArrayBuffer(bytes.length);
    var ia = new Uint8Array(ab);
    for (var i = 0; i < bytes.length; i++) ia[i] = bytes.charCodeAt(i);
    return new Blob([ab], { type: mime || 'image/jpeg' });
  }

  /* ── uid generator ──────────────────────────────────────────── */
  function _uid(prefix) {
    return (prefix || 'id') + '_' + Date.now() + '_' + Math.random().toString(36).slice(2, 6);
  }

  /* ════════════════════════════════════════════════════════════════
     PUBLIC API
     ════════════════════════════════════════════════════════════════ */
  return {

    /* ── INIT ──────────────────────────────────────────────────── */

    /**
     * Called after the user authorises Drive scopes.
     * Finds or creates the Drive folder + spreadsheet.
     */
    async init(accessToken, email) {
      _token      = accessToken;
      _adminEmail = email;
      _sheetMeta  = null;

      _folderId = await _findOrCreateFolder();
      _sheetId  = await _findOrCreateSheet(_folderId);

      /* persist for session restore */
      localStorage.setItem('gdrive_token',     _token);
      localStorage.setItem('gdrive_token_exp', String(Date.now() + 3500 * 1000));
      localStorage.setItem('gdrive_sheet_id',  _sheetId);
      localStorage.setItem('gdrive_folder_id', _folderId);
      localStorage.setItem('gdrive_email',     email || '');

      return { spreadsheetId: _sheetId, folderId: _folderId };
    },

    /** Restore from localStorage (between page loads) */
    restore() {
      _token    = localStorage.getItem('gdrive_token') || '';
      _sheetId  = localStorage.getItem('gdrive_sheet_id') || '';
      _folderId = localStorage.getItem('gdrive_folder_id') || '';
      _adminEmail = localStorage.getItem('gdrive_email') || '';
      var exp   = parseInt(localStorage.getItem('gdrive_token_exp') || '0', 10);
      return !!(_token && _sheetId && Date.now() < exp);
    },

    clear() {
      ['gdrive_token','gdrive_token_exp','gdrive_sheet_id','gdrive_folder_id','gdrive_email']
        .forEach(function (k) { localStorage.removeItem(k); });
      _token = _sheetId = _folderId = _adminEmail = null;
    },

    /* ── ACCESS / USERS ────────────────────────────────────────── */

    /** In Drive mode the signed-in user is always admin of their garage. */
    async checkAccess(email) {
      return { access: 'admin' };
    },

    /* ── BOXES ─────────────────────────────────────────────────── */

    async getAllBoxes() {
      var rows = await sheetGet('Boxes!A2:F2000');
      var boxes = rows.filter(function (r) { return r[0]; }).map(function (r) {
        return {
          name: r[0] || '', color: r[1] || '#f97316',
          location: r[2] || '',
          tags: (r[3] || '').split(',').filter(Boolean),
          coverImage: r[4] || '', createdAt: r[5] || ''
        };
      });
      return { boxes: boxes };
    },

    async addBox(params) {
      var name = params.name || params.box || '';
      await sheetAppend('Boxes', [
        name, params.color || '#f97316', params.location || '',
        (params.tags || []).join(','), '', new Date().toISOString()
      ]);
      return { success: true };
    },

    async deleteBox(params) {
      var rows = await sheetGet('Boxes!A2:A2000');
      for (var i = 0; i < rows.length; i++) {
        if (rows[i][0] === params.box) {
          await sheetDeleteRow('Boxes', i);
          return { success: true };
        }
      }
      return { success: false, error: 'Box not found' };
    },

    async renameBox(params) {
      var oldName = params.box, newName = params.newName;
      var rows = await sheetGet('Boxes!A2:F2000');
      for (var i = 0; i < rows.length; i++) {
        if (rows[i][0] === oldName) {
          var r = rows[i];
          await sheetUpdate('Boxes!A' + (i + 2), [[newName, r[1], r[2], r[3], r[4], r[5]]]);
          break;
        }
      }
      /* update items */
      var items = await sheetGet('Items!A2:K2000');
      for (var j = 0; j < items.length; j++) {
        if (items[j][1] === oldName) {
          items[j][1] = newName;
          await sheetUpdate('Items!A' + (j + 2), [items[j]]);
        }
      }
      return { success: true, newName: newName };
    },

    async setBoxColor(params) {
      var rows = await sheetGet('Boxes!A2:F2000');
      for (var i = 0; i < rows.length; i++) {
        if (rows[i][0] === params.box) {
          var r = rows[i];
          await sheetUpdate('Boxes!A' + (i + 2), [[r[0], params.color, r[2], r[3], r[4], r[5]]]);
          return { success: true };
        }
      }
      return { success: false };
    },

    async setBoxImage(params) {
      /* Store cover photo: upload to Drive, save URL in Boxes tab */
      var folderId = await _photoFolder(params.box);
      var blob     = _b64Blob(params.image, params.mimeType || 'image/jpeg');
      var fname    = 'cover_' + params.box + '.jpg';

      var form = new FormData();
      form.append('metadata', new Blob([JSON.stringify({ name: fname, parents: [folderId] })], { type: 'application/json' }));
      form.append('file', blob, fname);

      var r = await fetch(UPLOAD + '?uploadType=multipart&fields=id', {
        method: 'POST', headers: { 'Authorization': 'Bearer ' + _token }, body: form
      });
      var f = await r.json();
      await drivePermit(f.id);
      var url = 'https://drive.google.com/uc?export=view&id=' + f.id;

      /* write cover URL into Boxes tab */
      var rows = await sheetGet('Boxes!A2:F2000');
      for (var i = 0; i < rows.length; i++) {
        if (rows[i][0] === params.box) {
          var row = rows[i];
          await sheetUpdate('Boxes!A' + (i + 2), [[row[0], row[1], row[2], row[3], url, row[5]]]);
          return { success: true, url: url };
        }
      }
      return { success: false, error: 'Box not found' };
    },

    async getBoxImages(params) {
      var rows = await sheetGet('Boxes!A2:F2000');
      var images = {};
      rows.forEach(function (r) { if (r[0] && r[4]) images[r[0]] = r[4]; });
      return { images: images };
    },

    /* ── ITEMS ─────────────────────────────────────────────────── */

    async getBoxContents(params) {
      var box  = params.box;
      var rows = await sheetGet('Items!A2:K2000');
      var items = rows.filter(function (r) { return r[1] === box; }).map(function (r) {
        return {
          id: r[0], box: r[1], name: r[2], desc: r[3] || '',
          category: r[4] || '', qty: r[5] || '1',
          dateAdded: r[6] || '', lentTo: r[7] || '',
          dueDate: r[8] || '', lentContact: r[9] || '', lentNotes: r[10] || ''
        };
      });
      return { items: items };
    },

    async addItem(params) {
      var id = _uid('item');
      await sheetAppend('Items', [
        id, params.box, params.name, params.desc || '',
        params.category || '', params.qty || '1',
        new Date().toISOString().slice(0, 10), '', '', '', ''
      ]);
      return { success: true, id: id };
    },

    async editItem(params) {
      var rows = await sheetGet('Items!A2:K2000');
      for (var i = 0; i < rows.length; i++) {
        if (rows[i][0] === params.id) {
          var r = rows[i];
          await sheetUpdate('Items!A' + (i + 2), [[
            r[0], r[1],
            params.name        !== undefined ? params.name        : r[2],
            params.desc        !== undefined ? params.desc        : r[3],
            params.category    !== undefined ? params.category    : r[4],
            params.qty         !== undefined ? params.qty         : r[5],
            r[6],
            params.lentTo      !== undefined ? params.lentTo      : r[7],
            params.dueDate     !== undefined ? params.dueDate     : r[8],
            params.lentContact !== undefined ? params.lentContact : r[9],
            params.lentNotes   !== undefined ? params.lentNotes   : r[10]
          ]]);
          return { success: true };
        }
      }
      return { success: false, error: 'Item not found' };
    },

    async deleteItem(params) {
      var rows = await sheetGet('Items!A2:A2000');
      for (var i = 0; i < rows.length; i++) {
        if (rows[i][0] === params.id) {
          await sheetDeleteRow('Items', i);
          return { success: true };
        }
      }
      return { success: false };
    },

    async markReturned(params) {
      return this.editItem({ id: params.id, lentTo: '', dueDate: '', lentContact: '', lentNotes: '' });
    },

    /* ── SEARCH ─────────────────────────────────────────────────── */

    async getAllItems(params) {
      var rows = await sheetGet('Items!A2:K2000');
      var items = rows.filter(function (r) { return r[0]; }).map(function (r) {
        return {
          id: r[0], box: r[1], name: r[2], desc: r[3] || '',
          category: r[4] || '', qty: r[5] || '1', dateAdded: r[6] || '',
          lentTo: r[7] || '', dueDate: r[8] || ''
        };
      });
      return { items: items };
    },

    async getLentItems(params) {
      var rows = await sheetGet('Items!A2:K2000');
      var lent = rows.filter(function (r) { return r[7]; }).map(function (r) {
        return {
          id: r[0], box: r[1], name: r[2],
          lentTo: r[7], dueDate: r[8] || '', lentContact: r[9] || '', lentNotes: r[10] || ''
        };
      });
      return { items: lent };
    },

    /* ── PHOTOS ─────────────────────────────────────────────────── */

    async uploadPhoto(params) {
      var folderId = await _photoFolder(params.box);
      var blob  = _b64Blob(params.image || params.base64, params.mimeType || 'image/jpeg');
      var fname = _uid('photo') + '.jpg';

      var form = new FormData();
      form.append('metadata', new Blob([JSON.stringify({ name: fname, parents: [folderId] })], { type: 'application/json' }));
      form.append('file', blob, fname);

      var r = await fetch(UPLOAD + '?uploadType=multipart&fields=id', {
        method: 'POST', headers: { 'Authorization': 'Bearer ' + _token }, body: form
      });
      var f = await r.json();
      await drivePermit(f.id);
      var url = 'https://drive.google.com/uc?export=view&id=' + f.id;

      var photoId = _uid('ph');
      await sheetAppend('Photos', [
        photoId, params.box, params.itemId || '', params.itemName || '',
        url, f.id, new Date().toISOString().slice(0, 10)
      ]);
      return { success: true, url: url, id: photoId };
    },

    async getPhotos(params) {
      var rows = await sheetGet('Photos!A2:G2000');
      var photos = rows
        .filter(function (r) {
          return r[1] === params.box && (!params.itemId || r[2] === params.itemId);
        })
        .map(function (r) {
          return { id: r[0], box: r[1], itemId: r[2], itemName: r[3], url: r[4], driveFileId: r[5], dateAdded: r[6] };
        });
      return { photos: photos };
    },

    async deletePhoto(params) {
      var rows = await sheetGet('Photos!A2:F2000');
      for (var i = 0; i < rows.length; i++) {
        if (rows[i][0] === params.id) {
          await sheetDeleteRow('Photos', i);
          /* optionally delete the Drive file too */
          if (rows[i][5]) {
            fetch(DRIVE + '/' + rows[i][5], { method: 'DELETE', headers: auth() }).catch(function(){});
          }
          return { success: true };
        }
      }
      return { success: false };
    },

    /* ── GARAGE ZONE ─────────────────────────────────────────────── */

    async getGarageZone(params) {
      var rows = await sheetGet('GarageZone!A2:G2000');
      var items = rows.filter(function (r) { return r[0]; }).map(function (r) {
        return { id: r[0], name: r[1], desc: r[2] || '', category: r[3] || '',
                 location: r[4] || '', qty: r[5] || '1', dateAdded: r[6] || '' };
      });
      return { items: items };
    },

    async addGarageZoneItem(params) {
      var id = _uid('gz');
      await sheetAppend('GarageZone', [
        id, params.name, params.desc || '', params.category || '',
        params.location || '', params.qty || '1',
        new Date().toISOString().slice(0, 10)
      ]);
      return { success: true, id: id };
    },

    async editGarageZoneItem(params) {
      var rows = await sheetGet('GarageZone!A2:G2000');
      for (var i = 0; i < rows.length; i++) {
        if (rows[i][0] === params.id) {
          var r = rows[i];
          await sheetUpdate('GarageZone!A' + (i + 2), [[
            r[0],
            params.name     !== undefined ? params.name     : r[1],
            params.desc     !== undefined ? params.desc     : r[2],
            params.category !== undefined ? params.category : r[3],
            params.location !== undefined ? params.location : r[4],
            params.qty      !== undefined ? params.qty      : r[5],
            r[6]
          ]]);
          return { success: true };
        }
      }
      return { success: false };
    },

    async deleteGarageZoneItem(params) {
      var rows = await sheetGet('GarageZone!A2:A2000');
      for (var i = 0; i < rows.length; i++) {
        if (rows[i][0] === params.id) {
          await sheetDeleteRow('GarageZone', i);
          return { success: true };
        }
      }
      return { success: false };
    },

    /* ── CONFIG ──────────────────────────────────────────────────── */

    async getConfig(key) {
      var rows = await sheetGet('Config!A2:B100');
      var row = rows.find(function (r) { return r[0] === key; });
      return row ? row[1] : null;
    },

    async setConfig(key, value) {
      var rows = await sheetGet('Config!A2:B100');
      for (var i = 0; i < rows.length; i++) {
        if (rows[i][0] === key) {
          await sheetUpdate('Config!A' + (i + 2), [[key, value]]);
          return { success: true };
        }
      }
      await sheetAppend('Config', [key, value]);
      return { success: true };
    },

    /* ── TAGS ────────────────────────────────────────────────────── */

    async getBoxTags(params) {
      var rows = await sheetGet('Boxes!A2:D2000');
      var tags = {};
      rows.forEach(function (r) {
        if (r[0] && r[3]) tags[r[0]] = r[3].split(',').filter(Boolean);
      });
      return { tags: tags };
    },

    async setBoxTags(params) {
      var rows = await sheetGet('Boxes!A2:F2000');
      for (var i = 0; i < rows.length; i++) {
        if (rows[i][0] === params.box) {
          var r = rows[i];
          await sheetUpdate('Boxes!A' + (i + 2), [[r[0], r[1], r[2], (params.tags || []).join(','), r[4], r[5]]]);
          return { success: true };
        }
      }
      return { success: false };
    },

    /* ── SEARCH ─────────────────────────────────────────────────── */
    async searchAllBoxes(params) {
      var q = (params.q || '').toLowerCase().trim();
      if (!q) return { results: [] };
      var rows = await sheetGet('Items!A2:K2000');
      var results = rows
        .filter(function(r) {
          return r[0] && (
            (r[2] || '').toLowerCase().indexOf(q) >= 0 ||
            (r[3] || '').toLowerCase().indexOf(q) >= 0 ||
            (r[4] || '').toLowerCase().indexOf(q) >= 0 ||
            (r[1] || '').toLowerCase().indexOf(q) >= 0
          );
        })
        .map(function(r) {
          return {
            id: r[0], box: r[1], name: r[2], desc: r[3] || '',
            category: r[4] || '', qty: r[5] || '1'
          };
        });
      return { results: results };
    },

    /* ── PHOTO COUNTS ───────────────────────────────────────────── */
    async getBoxPhotoCounts(params) { return { counts: {} }; },

    /* ── ACTIVITY (Drive mode builds from local data) ───────────────── */
    async getRecentActivity(params) { return { activity: [] }; },
    async getActivity(params)       { return { activity: [] }; },

    /* ── ACTION NAME ALIASES (UI uses Apps Script names) ────────── */
    // deleteItem alias
    async removeItem(params)      { return this.deleteItem(params); },
    // markReturned alias
    async returnItem(params)      { return this.markReturned(params); },
    // getPhotos alias
    async getItemPhotos(params)   { return this.getPhotos(params); },
    // getGarageZone alias
    async getGarageItems(params)  { return this.getGarageZone(params); },
    // deleteGarageZoneItem alias
    async removeGarageItem(params){ return this.deleteGarageZoneItem(params); },

    // lendItem — update item's lending fields
    async lendItem(params) {
      return this.editItem({
        id: params.id, lentTo: params.lentTo || '',
        dueDate: params.dueDate || '', lentContact: params.lentContact || '',
        lentNotes: params.lentNotes || ''
      });
    },

    // moveItem — change an item's box
    async moveItem(params) {
      var rows = await sheetGet('Items!A2:K2000');
      for (var i = 0; i < rows.length; i++) {
        if (rows[i][0] === params.id) {
          var r = rows[i]; r[1] = params.newBox;
          await sheetUpdate('Items!A' + (i + 2), [r]);
          return { success: true };
        }
      }
      return { success: false, error: 'Item not found' };
    },

    /* ── ADMIN / PERMISSIONS (no-op in Drive mode) ───────────────── */
    async getPendingRequests()     { return { requests: [] }; },
    async approveRequest()         { return { success: true }; },
    async revokeAccess()           { return { success: true }; },
    async getUsers()               { return { users: [] }; },
    async requestAccess()          { return { success: true }; },
    async getMyBoxPerms()          { return { perms: {} }; },
    async submitAccessRequest()    { return { success: true }; },
    async approveUser()            { return { success: true }; },
    async denyUser()               { return { success: true }; },
    async sendLendReminder()       { return { success: true }; },
    async getApprovedUsers()       { return { users: [] }; },
    async getUserBoxPerms()        { return { perms: {} }; },
    async assignBoxCoOwner()       { return { success: true }; },
    async removeBoxCoOwner()       { return { success: true }; },
    async removeUser()             { return { success: true }; }

  }; // end return
})();

/* ════════════════════════════════════════════════════════════════════
   OneDrive storage layer stub (Phase 2 — Outlook / Microsoft sign-in)
   ════════════════════════════════════════════════════════════════════
   When a user signs in with Microsoft (MSAL.js), OneDriveDB mirrors
   the same public surface as GDriveDB but calls Microsoft Graph API
   and stores data in /Apps/GarageOrganizer/ in their OneDrive.

   Status: PLANNED — scaffold only.
*/
var OneDriveDB = (function () {
  var _token  = null;
  var GRAPH   = 'https://graph.microsoft.com/v1.0';
  var APP_PATH = '/me/drive/special/approot:/GarageOrganizer';

  return {
    init: async function (accessToken, email) {
      _token = accessToken;
      localStorage.setItem('onedrive_token',     _token);
      localStorage.setItem('onedrive_token_exp', String(Date.now() + 3500 * 1000));
      localStorage.setItem('onedrive_email',     email || '');
      // Phase 2: create /Apps/GarageOrganizer/garage_db.xlsx if not present
      console.log('[OneDriveDB] Phase 2 — not yet implemented');
      return { success: false, message: 'OneDrive support coming soon' };
    },
    restore: function () {
      _token = localStorage.getItem('onedrive_token') || '';
      var exp = parseInt(localStorage.getItem('onedrive_token_exp') || '0', 10);
      return !!(_token && Date.now() < exp);
    },
    clear: function () {
      ['onedrive_token','onedrive_token_exp','onedrive_email'].forEach(function (k) { localStorage.removeItem(k); });
      _token = null;
    },
    checkAccess: async function () { return { access: 'admin' }; }
  };
})();
