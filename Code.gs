const CONFIG = {

  SPREADSHEET_ID: 'YOUR_SPREADSHEET_ID_HERE',

  CHECKER_SHEET: 'CHECKER',
  PHTEAMS_SHEET: 'PHteams',

  PRIMARY_ENABLED_SHEET: 'Primary enabled per LDAP',
  SECONDARY_ENABLED_SHEET: 'Secondary enabled per LDAP',
  POOL_NAMES_SHEET: 'Pool ID Names',

  DATA_START_ROW: 2
};

/* =========================
   WEB APP ENTRY POINT
========================= */

// This function is required to serve the script as a website URL
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Dashboard')
    .setTitle('Pool Checker Dashboard')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1') // Makes it mobile responsive
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL); // Allows embedding if needed
}


/* =========================
   SHEETS MENU (Optional)
========================= */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Pool Checker')
    .addItem('Open Dashboard in Sheet', 'showDashboard')
    .addToUi();
}

function showDashboard() {
  const html = HtmlService
    .createHtmlOutputFromFile('Dashboard')
    .setWidth(1250)
    .setHeight(850);

  SpreadsheetApp.getUi().showModalDialog(
    html,
    'LDAP Pool ID Checker & Manager'
  );
}


/* =========================
   DATA
========================= */

function getDashboardData(team, agent) {

  const phteams = getSheet_(CONFIG.PHTEAMS_SHEET);

  team = String(team || '').trim();
  agent = String(agent || '').trim();

  const teams = getTeams_(phteams);

  const agents = getAgents_(phteams, team);

  const pools = getPools_(agent);

  const primary =
    pools.filter(x => x.type === 'Primary');

  const secondary =
    pools.filter(x => x.type === 'Secondary');


  return {
    team,
    agent,
    teams,
    agents,
    pools,

    stats: {
      primary: primary.length,
      secondary: secondary.length,
      total: pools.length
    }
  };
}




function getTeams_(sheet) {

  return sheet
    .getRange('D:D')   // PHteams column D only
    .getDisplayValues()
    .flat()
    .map(v => String(v || '').trim())
    .filter(Boolean);

}



function getAgents_(sheet, selectedTeam) {

  selectedTeam = String(selectedTeam || '').trim();

  if (!selectedTeam) {
    return [];
  }

  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return [];
  }

  const data = sheet
    .getRange(2,1,lastRow-1,2)
    .getDisplayValues();


  return data
    .filter(row =>
      String(row[1]).trim() === selectedTeam
    )
    .map(row =>
      String(row[0]).trim()
    )
    .filter(Boolean);

}





function getPools_(ldap) {

  const ss = getSpreadsheet_();

  const primarySheet =
    ss.getSheetByName(CONFIG.PRIMARY_ENABLED_SHEET);

  const secondarySheet =
    ss.getSheetByName(CONFIG.SECONDARY_ENABLED_SHEET);

  const nameSheet =
    ss.getSheetByName(CONFIG.POOL_NAMES_SHEET);


  if (!primarySheet ||
      !secondarySheet ||
      !nameSheet) {

    throw new Error(
      'Missing pool source sheets.'
    );
  }


  /*
    B1 is your LDAP selection column header
    Same as MATCH(B1, row 1)
  */



  if (!ldap) {
    return [];
  }


  // =====================
  // PRIMARY IDS
  // =====================

  const primaryHeaders =
    primarySheet
      .getRange(1,1,1,primarySheet.getLastColumn())
      .getDisplayValues()[0];


  const primaryColumn =
    primaryHeaders.indexOf(ldap);


  if (primaryColumn === -1) {
    return [];
  }


  const primaryIds =
    primarySheet
      .getRange(
        2,
        primaryColumn + 1,
        primarySheet.getLastRow() - 1,
        1
      )
      .getDisplayValues()
      .flat()
      .filter(String);



  // =====================
  // SECONDARY IDS
  // =====================

  const secondaryHeaders =
    secondarySheet
      .getRange(1,1,1,secondarySheet.getLastColumn())
      .getDisplayValues()[0];


  const secondaryColumn =
    secondaryHeaders.indexOf(ldap);


  const secondaryIds =
    secondaryColumn >= 0
      ?
        secondarySheet
        .getRange(
          2,
          secondaryColumn + 1,
          secondarySheet.getLastRow() - 1,
          1
        )
        .getDisplayValues()
        .flat()
        .filter(String)
      :
        [];



  // =====================
  // POOL ID → NAME MAP
  // =====================

  const nameData =
    nameSheet
      .getRange(
        1,
        1,
        nameSheet.getLastRow(),
        2
      )
      .getDisplayValues();


  const nameMap = {};

  nameData.forEach(row => {

    const name = row[0];
    const id = row[1];

    if (id) {
      nameMap[id] = name;
    }

  });



  // =====================
  // BUILD DASHBOARD DATA
  // =====================

  const pools = [];


  primaryIds.forEach(id => {

    pools.push({

      row: null,

      type: 'Primary',

      id: id,

      name:
        nameMap[id] || '',

      category:
        classifyPool(nameMap[id])

    });

  });



  secondaryIds.forEach(id => {

    pools.push({

      row: null,

      type: 'Secondary',

      id: id,

      name:
        nameMap[id] || '',

      category:
        classifyPool(nameMap[id])

    });

  });


  return pools;
}


function classifyPool(name) {
  const n = String(name || '').toLowerCase();

  if (!n) return 'Unnamed / ID Only';

  if (
    n.includes('google play') ||
    /\bplay\b/.test(n)
  ) {
    return 'Google Play';
  }

  if (
    n.includes('google tv') ||
    /\btv\b/.test(n)
  ) {
    return 'Google TV';
  }

  return 'Other';
}

function getColumnValues_(sheet, column, startRow) {
  const maxRows = sheet.getMaxRows();

  if (startRow > maxRows) return [];

  return sheet
    .getRange(startRow, column, maxRows - startRow + 1, 1)
    .getDisplayValues()
    .flat()
    .map(v => String(v || '').trim())
    .filter(Boolean);
}


/* =========================
   TEAM / AGENT
========================= */
function setTeam(team, agent) {

  return getDashboardData(
    team,
    agent
  );

}


function setAgent(team, agent) {

  return getDashboardData(
    team,
    agent
  );

}




function addCustomTeam(team) {
  team = String(team || '').trim();

  if (!team) {
    throw new Error('Please enter a team name.');
  }

  const sheet = getSheet_(CONFIG.PHTEAMS_SHEET);

  const values = sheet
    .getRange(1, 4, sheet.getMaxRows(), 1)
    .getDisplayValues()
    .flat();

  let row = values.findIndex(v => !String(v).trim()) + 1;

  if (row === 0) {
    row = sheet.getLastRow() + 1;
  }

  sheet.getRange(row, 4).setValue(team);

return getDashboardData();

}

function addCustomAgent(agent) {
  agent = String(agent || '').trim();

  if (!agent) {
    throw new Error('Please enter an agent name.');
  }

  const sheet = getSheet_(CONFIG.PHTEAMS_SHEET);
  const startRow = 3;

  const values = sheet
    .getRange(
      startRow,
      6,
      sheet.getMaxRows() - startRow + 1,
      1
    )
    .getDisplayValues()
    .flat();

  let offset = values.findIndex(v => !String(v).trim());

  if (offset === -1) {
    offset = values.length;
  }

  const row = startRow + offset;

sheet.getRange(row, 6).setValue(agent);

return getDashboardData();

}


/* =========================
   POOLS
========================= */

function addPool(type, name, id) {
  type = String(type || '');
  name = String(name || '').trim();
  id = String(id || '').trim();

  if (type !== 'Primary' && type !== 'Secondary') {
    throw new Error('Invalid pool type.');
  }

  if (!name && !id) {
    throw new Error('Pool name or ID is required.');
  }

  const sheet = getSheet_(CONFIG.CHECKER_SHEET);

  if (type === 'Secondary') {
    const lastRow = Math.max(sheet.getLastRow(), 1);

    if (lastRow >= CONFIG.DATA_START_ROW) {
      const values = sheet
        .getRange(
          CONFIG.DATA_START_ROW,
          3,
          lastRow - CONFIG.DATA_START_ROW + 1,
          4
        )
        .getDisplayValues();

      for (let i = 0; i < values.length; i++) {
        const primaryName = values[i][0];
        const primaryId = values[i][1];
        const secondaryName = values[i][2];
        const secondaryId = values[i][3];

        if (
          (primaryName || primaryId) &&
          !secondaryName &&
          !secondaryId
        ) {
          sheet
            .getRange(CONFIG.DATA_START_ROW + i, 5, 1, 2)
            .setValues([[name, id]]);

          return getDashboardData();
        }
      }
    }
  }

  const row = Math.max(
    sheet.getLastRow() + 1,
    CONFIG.DATA_START_ROW
  );

  const column = type === 'Primary' ? 3 : 5;

  sheet
    .getRange(row, column, 1, 2)
    .setValues([[name, id]]);

  return getDashboardData();
}

function editPool(type, row, name, id) {
  type = String(type || '');
  row = Number(row);

  if (!row || row < CONFIG.DATA_START_ROW) {
    throw new Error('Invalid sheet row.');
  }

  const column = type === 'Primary' ? 3 : 5;

  getSheet_(CONFIG.CHECKER_SHEET)
    .getRange(row, column, 1, 2)
    .setValues([
      [
        String(name || '').trim(),
        String(id || '').trim()
      ]
    ]);

  return getDashboardData();
}

function deletePool(type, row) {
  type = String(type || '');
  row = Number(row);

  if (!row || row < CONFIG.DATA_START_ROW) {
    throw new Error('Invalid sheet row.');
  }

  const sheet = getSheet_(CONFIG.CHECKER_SHEET);
  const column = type === 'Primary' ? 3 : 5;

  sheet.getRange(row, column, 1, 2).clearContent();

  return getDashboardData();
}

function movePool(type, row, direction) {
  type = String(type || '');
  row = Number(row);
  direction = Number(direction);

  if (!row || row < CONFIG.DATA_START_ROW) {
    throw new Error('Invalid sheet row.');
  }

  if (direction !== -1 && direction !== 1) {
    throw new Error('Invalid move direction.');
  }

  const sheet = getSheet_(CONFIG.CHECKER_SHEET);
  const column = type === 'Primary' ? 3 : 5;

  const targetRow = row + direction;

  if (targetRow < CONFIG.DATA_START_ROW) {
    return getDashboardData();
  }

  const lastRow = Math.max(sheet.getLastRow(), CONFIG.DATA_START_ROW);

  if (targetRow > lastRow) {
    return getDashboardData();
  }

  const current = sheet
    .getRange(row, column, 1, 2)
    .getValues()[0];

  const target = sheet
    .getRange(targetRow, column, 1, 2)
    .getValues()[0];

  sheet
    .getRange(row, column, 1, 2)
    .setValues([target]);

  sheet
    .getRange(targetRow, column, 1, 2)
    .setValues([current]);

  return getDashboardData();
}


/* =========================
   DROPDOWNS
========================= */

function installDropdowns() {
  const ss = CONFIG.SPREADSHEET_ID 
    ? SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID) 
    : SpreadsheetApp.getActiveSpreadsheet();
    
  const checker = ss.getSheetByName(CONFIG.CHECKER_SHEET);
  const phteams = ss.getSheetByName(CONFIG.PHTEAMS_SHEET);

  if (!checker || !phteams) {
    throw new Error('CHECKER or PHteams sheet was not found.');
  }

  const teamRule = SpreadsheetApp
    .newDataValidation()
    .requireValueInRange(
      phteams.getRange('D:D'),
      true
    )
    .setAllowInvalid(true)
    .build();

  const agentRule = SpreadsheetApp
    .newDataValidation()
    .requireValueInRange(
      phteams.getRange('F3:F'),
      true
    )
    .setAllowInvalid(true)
    .build();

  checker.getRange('A1').setDataValidation(teamRule);
  checker.getRange('B1').setDataValidation(agentRule);

  SpreadsheetApp.getUi().alert(
    'Dropdowns installed on CHECKER!A1 and CHECKER!B1.'
  );
}


/* =========================
   HELPERS
========================= */

function getSheet_(name) {

  const ss = getSpreadsheet_();

  const sheet = ss.getSheetByName(name);

  if (!sheet) {
    throw new Error(
      'Sheet "' + name + '" was not found.'
    );
  }

  return sheet;
}

function getSpreadsheet_() {

  if (
    CONFIG.SPREADSHEET_ID &&
    CONFIG.SPREADSHEET_ID !== 'YOUR_SPREADSHEET_ID_HERE'
  ) {
    return SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  }

  return SpreadsheetApp.getActiveSpreadsheet();
}

function getTeamPoolSummary() {

  const sheet = getSheet_('Counts');

  const startRow = 3;
  const lastRow = sheet.getLastRow();

  if (lastRow < startRow) {
    return [];
  }


  const data = sheet
    .getRange(startRow, 1, lastRow - startRow + 1, 5)
    .getDisplayValues();


  const result = data
    .filter(row =>
      row[0] || row[1]
    )
    .map(row => {

      return {

        agent: row[0],
        team: row[1],

        primary: Number(row[2]) || 0,

        secondary: Number(row[3]) || 0,

        total: Number(row[4]) || 0

      };

    });


  // Sort Team Name then Agent Name
  result.sort((a,b)=>{

    const teamSort =
      a.team.localeCompare(b.team);

    if (teamSort !== 0) {
      return teamSort;
    }

    return a.agent.localeCompare(b.agent);

  });


  return result;

}
