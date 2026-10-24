const CFG = {
  MASTER: 'BM Master',
  AUDIT: 'BM Audit Master',
  LINKS: 'Technician Links',
  ADMIN: 'Admin',
  PINKEY: 'BM_ADMIN_PIN'
};


function doGet(e) {
  const template = HtmlService.createTemplateFromFile('Index');

  template.token = e && e.parameter && e.parameter.token
    ? String(e.parameter.token).trim()
    : '';

  return template
    .evaluate()
    .setTitle('BM Audit App')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}


/* ================================
   INITIAL SETUP
================================ */

function setupBMApp() {
  // NON-DESTRUCTIVE setup: never clear existing audit/master/link/settings data.
  const ss = SpreadsheetApp.getActive();
  const master = ss.getSheetByName(CFG.MASTER);
  if (!master) throw new Error('BM Master sheet not found. Existing master data was not changed.');

  const auditHeaders = [
    'Audit Date & Time','Audit Date','Technician Name','Site ID','NE BM','BM Make',
    'Physical BM','Zabbix Visible BM','Faulty BM at Site','BM Sent to WH','BM Theft',
    'BM in Site Down','BM in Degrow','SMPS Issue','SMPS Upgradation','BM OEM','SMPS OEM',
    'ISP Issue','Workable BM Total','Missing BM','Validation Status'
  ];
  const linksHeaders = ['Technician Name','Token','Technician Link','Active'];

  let audit = ss.getSheetByName(CFG.AUDIT);
  if (!audit) audit = ss.insertSheet(CFG.AUDIT);
  if (audit.getLastRow() === 0) audit.appendRow(auditHeaders);
  audit.setFrozenRows(1);

  let links = ss.getSheetByName(CFG.LINKS);
  if (!links) links = ss.insertSheet(CFG.LINKS);
  if (links.getLastRow() === 0) links.appendRow(linksHeaders);
  links.setFrozenRows(1);

  let admin = ss.getSheetByName(CFG.ADMIN);
  if (!admin) admin = ss.insertSheet(CFG.ADMIN);
  if (admin.getLastRow() === 0) admin.getRange(1, 1, 3, 2).setValues([
    ['Setting','Value'], ['Admin PIN',''], ['Web App URL','']
  ]);
  else {
    if (!admin.getRange('A1').getValue()) admin.getRange('A1:B1').setValues([['Setting','Value']]);
    if (!admin.getRange('A2').getValue()) admin.getRange('A2').setValue('Admin PIN');
    if (!admin.getRange('A3').getValue()) admin.getRange('A3').setValue('Web App URL');
  }

  let pin = PropertiesService.getScriptProperties().getProperty(CFG.PINKEY);
  if (!pin) {
    pin = String(Math.floor(100000 + Math.random() * 900000));
    PropertiesService.getScriptProperties().setProperty(CFG.PINKEY, pin);
  }
  admin.getRange('B2').setValue(pin);

  let webUrl = ScriptApp.getService().getUrl() || '';
  webUrl = webUrl.replace(/\/dev(?:\/)?$/, '/exec');
  if (webUrl) admin.getRange('B3').setValue(webUrl);

  const technicians = [...new Set(getMasterRows_().map(x => x.tech).filter(Boolean))].sort();
  const existing = links.getLastRow() > 1
    ? links.getRange(2, 1, links.getLastRow() - 1, Math.max(4, links.getLastColumn())).getValues()
    : [];
  const rowByTech = {};
  existing.forEach((r, i) => {
    const name = String(r[0] || '').trim();
    if (name && !rowByTech[name]) rowByTech[name] = i + 2;
  });

  technicians.forEach(tech => {
    const rowNo = rowByTech[tech];
    const existingToken = rowNo ? String(links.getRange(rowNo, 2).getValue() || '').trim() : '';
    const token = existingToken || Utilities.getUuid();
    const link = webUrl ? webUrl + '?token=' + encodeURIComponent(token) : '';
    if (rowNo) {
      // Update only this technician's link row; preserve all other rows and tokens.
      links.getRange(rowNo, 1, 1, 4).setValues([[tech, token, link, 'YES']]);
    } else {
      links.appendRow([tech, token, link, 'YES']);
    }
  });

  [master, audit, links, admin].forEach(sheet => {
    if (sheet.getLastColumn() > 0) {
      sheet.getRange(1, 1, 1, sheet.getLastColumn()).setFontWeight('bold');
      sheet.autoResizeColumns(1, sheet.getLastColumn());
    }
  });

  return { sites: getMasterRows_().length, technicians: technicians.length, pin: pin, url: webUrl };
}

/* Backup all previous rows before any latest-record overwrite or deduplication. */
function backupAuditRows_(rows, reason) {
  if (!rows || !rows.length) return;
  const ss = SpreadsheetApp.getActive();
  let backup = ss.getSheetByName('BM Audit Backup');
  const audit = ss.getSheetByName(CFG.AUDIT);
  const headers = audit.getLastColumn() ? audit.getRange(1, 1, 1, audit.getLastColumn()).getValues()[0] : [];
  if (!backup) backup = ss.insertSheet('BM Audit Backup');
  if (backup.getLastRow() === 0) backup.appendRow(['Backup Timestamp','Backup Reason'].concat(headers));
  const stamp = new Date();
  const out = rows.map(item => [stamp, reason].concat(item.values));
  backup.getRange(backup.getLastRow() + 1, 1, out.length, out[0].length).setValues(out);
}

/* Keep exactly one latest row per Site ID. Existing rows are backed up before edits. */
function saveLatestAuditForSite_(sheet, siteId, auditRow, reason) {
  const last = sheet.getLastRow();
  const matches = [];
  if (last >= 2) {
    const ids = sheet.getRange(2, 4, last - 1, 1).getDisplayValues();
    ids.forEach((r, i) => { if (String(r[0] || '').trim() === String(siteId).trim()) matches.push(i + 2); });
  }
  if (!matches.length) {
    sheet.appendRow(auditRow);
    return;
  }
  const oldRows = matches.map(rowNo => ({ rowNo: rowNo, values: sheet.getRange(rowNo, 1, 1, sheet.getLastColumn()).getValues()[0] }));
  backupAuditRows_(oldRows, reason || 'Latest record updated');
  const keepRow = matches[matches.length - 1];
  // Remove duplicate legacy rows only after each one has been copied to BM Audit Backup.
  matches.slice(0, -1).reverse().forEach(rowNo => sheet.deleteRow(rowNo));
  const adjustedKeepRow = keepRow - (matches.length - 1);
  sheet.getRange(adjustedKeepRow, 1, 1, auditRow.length).setValues([auditRow]);
}

/* ================================
   READ BM MASTER
================================ */

function getMasterRows_() {

  const sheet =
    SpreadsheetApp
      .getActive()
      .getSheetByName(CFG.MASTER);

  if (!sheet) {
    throw new Error('BM Master sheet not found.');
  }

  // Display values preserve formatted Site IDs (including leading zeros).
  const values =
    sheet
      .getDataRange()
      .getDisplayValues();

  if (values.length === 0) {
    return [];
  }

  let headerRowIndex = -1;

  for (let i = 0; i < values.length; i++) {

    const rowText =
      values[i]
        .map(x => String(x).trim())
        .join('|');

    if (
      rowText.includes('SAP ID') &&
      rowText.includes('BM MAKE_') &&
      rowText.includes('Tech. Name')
    ) {

      headerRowIndex = i;
      break;

    }

  }

  if (headerRowIndex === -1) {

    throw new Error(
      'BM Master header row not found. Please check SAP ID / BM MAKE_ / Tech. Name headers.'
    );

  }

  const headers =
    values[headerRowIndex]
      .map(x => String(x).trim());

  const index = {};

  headers.forEach(
    (header, i) => {
      index[header] = i;
    }
  );

  const requiredHeaders = [
    'SAP ID',
    'BM MAKE_',
    'Tech. Name',
    'BM As per NE'
  ];

  requiredHeaders.forEach(
    function(header) {

      if (index[header] === undefined) {

        throw new Error(
          'Missing BM Master header: ' +
          header
        );

      }

    }
  );

  return values

    .slice(headerRowIndex + 1)

    .filter(
      row =>
        row[index['SAP ID']]
    )

    .map(
      function(row) {

        return {

          site:
            String(
              row[index['SAP ID']]
            ).trim(),

          make:
            String(
              row[index['BM MAKE_']] || ''
            ).trim(),

          tech:
            String(
              row[index['Tech. Name']] || ''
            ).trim(),

          ne:
            Number(String(row[index['BM As per NE']] || '0').replace(/,/g, '')) || 0,

          mp:
            index['MP'] !== undefined
              ? String(
                  row[index['MP']] || ''
                ).trim()
              : '',

          siteType:
            index['SiteType'] !== undefined
              ? String(row[index['SiteType']] || '').trim()
              : '',

          cluster:
            (() => {
              const k = ['JC/Cluster', 'JC', 'Cluster', 'JC Name'].find(h => index[h] !== undefined);
              return k ? String(row[index[k]] || '').trim() : '';
            })()

        };

      }
    );

}

/* ================================
   TECHNICIAN TOKEN
================================ */

function techByToken_(token) {

  const sheet =
    SpreadsheetApp
      .getActive()
      .getSheetByName(CFG.LINKS);


  const values =
    sheet
      .getDataRange()
      .getValues();


  const row =
    values
      .slice(1)
      .find(r =>

        String(r[1]).trim()
        ===
        String(token || '').trim()

        &&

        String(r[3])
          .toUpperCase()
        === 'YES'

      );


  if (!row) {

    throw new Error(
      'Invalid or inactive technician link.'
    );

  }


  return String(row[0]).trim();
}


/* ================================
   TECHNICIAN APP DATA
================================ */

function getAppData(token) {

  const technician =
    techByToken_(token);


  const sites =
    getMasterRows_()
      .filter(
        x => x.tech === technician
      );


  return {

    technician: technician,

    sites: sites,

    recent:
      getRecentAudits_(technician)

  };

}


/* ================================
   RECENT AUDITS
================================ */

function getRecentAudits_(technician) {

  const sheet =
    SpreadsheetApp
      .getActive()
      .getSheetByName(CFG.AUDIT);


  if (
    !sheet ||
    sheet.getLastRow() < 2
  ) {

    return [];

  }


  const values =
    sheet
      .getRange(
        2,
        1,
        sheet.getLastRow() - 1,
        21
      )
      .getValues();


  return values
  .filter(r =>
    r[2] === technician
    &&
    r[0] instanceof Date
  )
  .reverse()

  .map(r => {
    return {
  time: r[0] instanceof Date
    ? Utilities.formatDate(
        r[0],
        Session.getScriptTimeZone(),
        "yyyy-MM-dd HH:mm:ss"
      )
    : String(r[0] || ''),

 site: String(r[3] || ''),

physical: Number(r[6] || 0),

visible: Number(r[7] || 0),

nonWorkable:
  Number(r[8] || 0) +
  Number(r[9] || 0) +
  Number(r[10] || 0) +
  Number(r[11] || 0) +
  Number(r[12] || 0),

workable: Number(r[18] || 0)
};

});
}



/* ================================
   SUBMIT BM AUDIT
================================ */

function submitAudit(token, p) {

  const technician =
    techByToken_(token);


  const master =
    getMasterRows_();


  const record =
    master.find(x =>

      x.site ===
      String(p.site || '').trim()

      &&

      x.tech === technician

    );


  if (!record) {

    throw new Error(
      'Selected site is not assigned to this technician.'
    );

  }


  const keys = [

    'physical',
    'visible',
    'faulty',
    'wh',
    'theft',
    'siteDown',
    'degrow',

    'smpsIssue',
    'smpsUpgrade',
    'bmOem',
    'smpsOem',
    'ispIssue'

  ];


  const n = {};


  keys.forEach(key => {

    n[key] =
      Number(
        p[key] || 0
      );

  });


  /* ---------- 0 TO 10 VALIDATION ---------- */

  if (

    keys.some(key =>

      !Number.isInteger(n[key])
      ||
      n[key] < 0
      ||
      n[key] > 10

    )

  ) {

    throw new Error(
      'All BM counts must be 0 to 10.'
    );

  }


  /* ---------- WORKABLE ---------- */

  const workable =

    n.smpsIssue
    +
    n.smpsUpgrade
    +
    n.bmOem
    +
    n.smpsOem
    +
    n.ispIssue;


  /* ---------- STATUS TOTAL ---------- */

  const statusTotal =
  n.faulty
  + n.wh
  + n.theft
  + n.siteDown
  + n.degrow
  + n.visible;



const missing = Math.max(
  0,
  Number(n.physical || 0)
  - n.faulty
  - n.wh
  - n.theft
  - n.siteDown
  - n.degrow
  - workable
  - n.visible
);

// Visible BM is included in the BM reconciliation.
// Missing BM is calculated after deducting Visible BM.

if (
  statusTotal + workable + missing !==
  Number(n.physical || 0)
) {
 throw new Error(
  'Faulty + WH + Theft + Site Down + Degrow + Workable + Visible + Missing must equal Physical BM.'
);
}


  /* ---------- REMARK ---------- */

  


  /* ---------- SAVE ---------- */

  const now =
    new Date();


  const sheet =
    SpreadsheetApp
      .getActive()
      .getSheetByName(CFG.AUDIT);


  
  const siteId = String(record.site || '').trim();
  const auditRow = [
    now,
    Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
    technician, record.site, record.ne, record.make,
    n.physical, n.visible, n.faulty, n.wh, n.theft, n.siteDown, n.degrow,
    n.smpsIssue, n.smpsUpgrade, n.bmOem, n.smpsOem, n.ispIssue,
    workable, missing, 'OK'
  ];
  saveLatestAuditForSite_(sheet, siteId, auditRow, 'Technician submitted a new audit for this Site ID');

  return {

    message:
      'BM Audit submitted successfully.',

    recent:
      getRecentAudits_(technician)

  };

}


/* ================================
   ADMIN LOGIN
================================ */

function adminLogin(pin) {

  const savedPin =
    PropertiesService
      .getScriptProperties()
      .getProperty(
        CFG.PINKEY
      );


  return (

    String(pin || '').trim()
    ===
    String(savedPin || '').trim()

  );

}


/* ================================
   ADMIN DATA
================================ */

function getAdminData(pin) {
  if (!adminLogin(pin)) throw new Error('Invalid Admin PIN.');
  const master = getMasterRows_();
  const auditSheet = SpreadsheetApp.getActive().getSheetByName(CFG.AUDIT);
  const audits = [];
  if (auditSheet && auditSheet.getLastRow() > 1) {
    const values = auditSheet.getRange(2, 1, auditSheet.getLastRow() - 1, 21).getValues();
    values.forEach(r => audits.push({
      time: r[0] instanceof Date ? Utilities.formatDate(r[0], Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss') : String(r[0] || ''),
      date: r[1] instanceof Date ? Utilities.formatDate(r[1], Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(r[1] || ''),
      technician: String(r[2] || ''), site: String(r[3] || ''), ne: Number(r[4] || 0), make: String(r[5] || ''),
      physical: Number(r[6] || 0), visible: Number(r[7] || 0), faulty: Number(r[8] || 0), wh: Number(r[9] || 0),
      theft: Number(r[10] || 0), siteDown: Number(r[11] || 0), degrow: Number(r[12] || 0), smpsIssue: Number(r[13] || 0),
      smpsUpgrade: Number(r[14] || 0), bmOEM: Number(r[15] || 0), smpsOEM: Number(r[16] || 0), ispIssue: Number(r[17] || 0),
      workable: Number(r[18] || 0), missing: Number(r[19] || 0), status: String(r[20] || '')
    }));
  }
  // Never send technician token links to the browser; they remain in the spreadsheet only.
  return { technicians: [...new Set(master.map(x => x.tech).filter(Boolean))].sort(), sites: master, audits: audits };
}

/* ================================
   CHANGE TECHNICIAN
================================ */

function updateSiteTechnician(
  pin,
  site,
  newTechnician
) {

  if (
    !adminLogin(pin)
  ) {

    throw new Error(
      'Invalid Admin PIN.'
    );

  }


  const sheet =
    SpreadsheetApp
      .getActive()
      .getSheetByName(
        CFG.MASTER
      );


  const values = sheet.getDataRange().getDisplayValues();
  let headerRowIndex = -1;
  for (let r = 0; r < values.length; r++) {
    const rowText = values[r].map(x => String(x).trim()).join('|');
    if (rowText.includes('SAP ID') && rowText.includes('Tech. Name')) { headerRowIndex = r; break; }
  }
  if (headerRowIndex < 0) throw new Error('BM Master headers missing.');
  const headers = values[headerRowIndex].map(x => String(x).trim());
  const siteIndex = headers.indexOf('SAP ID');
  const techIndex = headers.indexOf('Tech. Name');


  if (
    siteIndex < 0
    ||
    techIndex < 0
  ) {

    throw new Error(
      'Master headers missing.'
    );

  }


  let changed = 0;


  for (
    let i = headerRowIndex + 1;
    i < values.length;
    i++
  ) {

    if (

      String(
        values[i][siteIndex]
      ).trim()
      ===
      String(site).trim()

    ) {

      sheet
        .getRange(
          i + 1,
          techIndex + 1
        )
        .setValue(
          newTechnician
        );

      changed++;

    }

  }


  if (!changed) {

    throw new Error(
      'Site not found.'
    );

  }


  setupBMApp();


  return {

    message:
      site +
      ' assigned to ' +
      newTechnician

  };

}
function updateAuditTechnician(pin, site, newTechnician) {
  if (!adminLogin(pin)) throw new Error('Invalid Admin PIN.');


  site = String(site || '').trim();
  newTechnician = String(newTechnician || '').trim();

  if (!site) {
    throw new Error('Site ID is required.');
  }

  if (!newTechnician) {
    throw new Error('Technician Name is required.');
  }

  const sheet =
    SpreadsheetApp
      .getActive()
      .getSheetByName(CFG.AUDIT);

  if (!sheet || sheet.getLastRow() < 2) {
    throw new Error('No audit data found.');
  }

  const values =
    sheet
      .getRange(
        2,
        1,
        sheet.getLastRow() - 1,
        21
      )
      .getValues();

  let foundRow = 0;

  for (let i = 0; i < values.length; i++) {

    if (String(values[i][3] || '').trim() === site) {
      foundRow = i + 2; // Keep the latest matching row.
    }
  }

  if (!foundRow) {
    throw new Error(
      'Site not found in Audit History: ' + site
    );
  }

  // Backup the existing row before changing the technician.
  backupAuditRows_([{ rowNo: foundRow, values: sheet.getRange(foundRow, 1, 1, sheet.getLastColumn()).getValues()[0] }], 'Admin changed audit technician');
  // Column C = Technician Name
  sheet.getRange(foundRow, 3).setValue(newTechnician);

  return {
    message:
      'Technician updated successfully for site: ' +
      site,

    recent:
      getRecentAudits_(newTechnician)
  };
}
function getAuditBySite(site, pin) {
  if (!adminLogin(pin)) throw new Error('Invalid Admin PIN.');
  site = String(site || '').trim();
  if (!site) throw new Error('Site ID is required.');
  const sheet = SpreadsheetApp.getActive().getSheetByName(CFG.AUDIT);
  if (!sheet || sheet.getLastRow() < 2) throw new Error('No audit data found.');
  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 21).getValues();
  let found = null;
  values.forEach(r => {
    if (String(r[3] || '').trim() === site) {
      found = {
        site: r[3], technician: r[2], physical: Number(r[6] || 0), visible: Number(r[7] || 0),
        faulty: Number(r[8] || 0), wh: Number(r[9] || 0), theft: Number(r[10] || 0),
        siteDown: Number(r[11] || 0), degrow: Number(r[12] || 0), smpsIssue: Number(r[13] || 0),
        smpsUpgrade: Number(r[14] || 0), bmOEM: Number(r[15] || 0), smpsOEM: Number(r[16] || 0),
        ispIssue: Number(r[17] || 0), workable: Number(r[18] || 0), missing: Number(r[19] || 0), status: r[20] || ''
      };
    }
  });
  if (found) return found;
  throw new Error('Site not found in Audit History: ' + site);
}

function updateAuditRecord(p, pin) {
  if (!adminLogin(pin)) throw new Error('Invalid Admin PIN.');

  if (!p) {
    throw new Error('Audit data is required.');
  }

  const site = String(p.site || '').trim();
  const technician = String(p.technician || '').trim();

  if (!site) {
    throw new Error('Site ID is required.');
  }

  if (!technician) {
    throw new Error('Technician Name is required.');
  }

  const fields = [
    'physical',
    'visible',
    'faulty',
    'wh',
    'theft',
    'siteDown',
    'degrow',
    'smpsIssue',
    'smpsUpgrade',
    'bmOEM',
    'smpsOEM',
    'ispIssue'
  ];

  const v = {};

  fields.forEach(function(key) {

    const n = Number(p[key]);

    if (!Number.isFinite(n) || n < 0 || n > 10) {
      throw new Error(
        key + ' must be between 0 and 10.'
      );
    }

    v[key] = n;
  });


  // WORKABLE BM
  const workable =
    v.smpsIssue +
    v.smpsUpgrade +
    v.bmOEM +
    v.smpsOEM +
    v.ispIssue;


  // MISSING BM
  const missing =
    v.physical -
    v.faulty -
    v.wh -
    v.theft -
    v.siteDown -
    v.degrow -
    workable -
    v.visible;


  if (missing < 0 || missing > 10) {
    throw new Error(
      'Calculated Missing BM is invalid: ' + missing
    );
  }


  // PHYSICAL RECONCILIATION
  const reconciliation =
    v.faulty +
    v.wh +
    v.theft +
    v.siteDown +
    v.degrow +
    workable +
    v.visible +
    missing;


  if (reconciliation !== v.physical) {
    throw new Error(
      'Physical BM reconciliation failed. ' +
      'Physical = ' + v.physical +
      ', Calculated = ' + reconciliation
    );
  }


  const sheet =
    SpreadsheetApp
      .getActive()
      .getSheetByName(CFG.AUDIT);


  if (!sheet || sheet.getLastRow() < 2) {
    throw new Error('No audit data found.');
  }


  const values =
    sheet
      .getRange(
        2,
        1,
        sheet.getLastRow() - 1,
        21
      )
      .getValues();


  let foundRow = 0;


  for (let i = 0; i < values.length; i++) {

    if (
      String(values[i][3] || '').trim() === site
    ) {
      // Continue scanning so the newest (last) matching Site ID row wins.
      foundRow = i + 2;
    }
  }


  if (!foundRow) {
    throw new Error(
      'Site not found in Audit History: ' + site
    );
  }


  // Preserve original date/site/master values; backup the old row before updating.
  const original = values[foundRow - 2];
  const updatedRow = [
    original[0], original[1], technician, original[3], original[4], original[5],
    v.physical, v.visible, v.faulty, v.wh, v.theft, v.siteDown, v.degrow,
    v.smpsIssue, v.smpsUpgrade, v.bmOEM, v.smpsOEM, v.ispIssue,
    workable, missing, 'OK'
  ];
  saveLatestAuditForSite_(sheet, site, updatedRow, 'Admin edited the latest audit record');

  return {
    message:
      'Audit updated successfully for site: ' + site,

    technician: technician,

    recent:
      getRecentAudits_(technician)
  };
}