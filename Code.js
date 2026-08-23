/*************************************************
 * CUSTOMERS HUB
 * Sales Territory / Customer Mapping
 *************************************************/

const AREAS_SHEET = 'Areas';
const CUSTOMERS_SHEET = 'Customers';
const PEOPLE_SHEET = 'People';
const API_KEY_SHEET = 'API Key';


/*************************************************
 * WEB APP
 *************************************************/

function doGet() {

  return HtmlService
    .createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Customers Hub')
    .setXFrameOptionsMode(
      HtmlService.XFrameOptionsMode.ALLOWALL
    );

}


/*************************************************
 * INCLUDE HTML
 *************************************************/

function include(filename) {

  return HtmlService
    .createHtmlOutputFromFile(filename)
    .getContent();

}


/*************************************************
 * APP CONFIG
 *************************************************/

function getAppConfig() {

  return {
    mapApiKey: getMapApiKey()
  };

}


/*************************************************
 * GOOGLE MAPS API KEY
 *************************************************/

function getMapApiKey() {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const sheets =
    ss.getSheets();

  let sheet = null;


  for (let i = 0; i < sheets.length; i++) {

    const name =
      sheets[i]
        .getName()
        .trim()
        .toLowerCase();

    if (name === 'api key') {

      sheet = sheets[i];

      break;

    }

  }


  if (!sheet) {

    const availableSheets =
      sheets
        .map(s => s.getName())
        .join(', ');

    throw new Error(
      'API Key sheet not found. Available sheets: ' +
      availableSheets
    );

  }


  const values =
    sheet
      .getDataRange()
      .getDisplayValues();


  for (let r = 0; r < values.length; r++) {

    for (let c = 0; c < values[r].length; c++) {

      const value =
        String(values[r][c] || '').trim();


      if (
        value.startsWith('AIza') &&
        value.length > 20
      ) {

        return value;

      }

    }

  }


  throw new Error(
    'Google Maps API key was not found in the API Key sheet.'
  );

}


/*************************************************
 * GET INDUSTRIAL HUBS
 *************************************************/

function getIndustrialHubs() {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const sheet =
    ss.getSheetByName(AREAS_SHEET);


  if (!sheet) {

    throw new Error(
      'Areas sheet not found.'
    );

  }


  const data =
    sheet
      .getDataRange()
      .getDisplayValues();


  if (data.length < 2) {

    return [];

  }


  const headers =
    data[0].map(normalizeHeader);


  const hubIndex =
    findHeader(
      headers,
      [
        'industrial hub',
        'industrialhub'
      ]
    );


  const cityIndex =
    findHeader(
      headers,
      ['city']
    );


  const latitudeIndex =
    findHeader(
      headers,
      [
        'latitude',
        'lat'
      ]
    );


  const longitudeIndex =
    findHeader(
      headers,
      [
        'longitude',
        'long',
        'lng'
      ]
    );


  if (
    hubIndex === -1 ||
    cityIndex === -1 ||
    latitudeIndex === -1 ||
    longitudeIndex === -1
  ) {

    throw new Error(
      'Areas sheet must contain Industrial Hub, City, Latitude and Longitude columns.'
    );

  }


  const hubs = {};


  for (let i = 1; i < data.length; i++) {

    const row = data[i];


    const hub =
      String(
        row[hubIndex] || ''
      ).trim();


    const city =
      String(
        row[cityIndex] || ''
      ).trim();


    const latitude =
      parseFloat(
        row[latitudeIndex]
      );


    const longitude =
      parseFloat(
        row[longitudeIndex]
      );


    if (!hub) {
      continue;
    }


    if (
      isNaN(latitude) ||
      isNaN(longitude)
    ) {
      continue;
    }


    const key =
      hub.toLowerCase();


    if (!hubs[key]) {

      hubs[key] = {

        hub: hub,

        city: city,

        latitude: latitude,

        longitude: longitude

      };

    }

  }


  return Object
    .values(hubs)
    .sort(
      (a, b) =>
        a.hub.localeCompare(b.hub)
    );

}


/*************************************************
 * GET CUSTOMERS BY INDUSTRIAL HUB
 *************************************************/

function getCustomersByHub(hubName) {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const sheet =
    ss.getSheetByName(CUSTOMERS_SHEET);


  if (!sheet) {

    throw new Error(
      'Customers sheet not found.'
    );

  }


  const data =
    sheet
      .getDataRange()
      .getDisplayValues();


  if (data.length < 2) {

    return [];

  }


  const headers =
    data[0].map(normalizeHeader);


  const indexes = {

    companyId:
      findHeader(
        headers,
        [
          'company id',
          'companyid'
        ]
      ),

    customerName:
      findHeader(
        headers,
        [
          'customer name',
          'customername'
        ]
      ),

    industrialArea:
      findHeader(
        headers,
        [
          'industrial area',
          'industrialarea'
        ]
      ),

    industrialHub:
      findHeader(
        headers,
        [
          'industrial hub',
          'industrialhub'
        ]
      ),

    city:
      findHeader(
        headers,
        ['city']
      ),

    type:
      findHeader(
        headers,
        ['type']
      ),

    status:
      findHeader(
        headers,
        ['status']
      ),

    potential:
      findHeader(
        headers,
        ['potential']
      ),

    website:
      findHeader(
        headers,
        ['website']
      ),

    notes:
      findHeader(
        headers,
        ['notes']
      )

  };


  if (
    indexes.companyId === -1 ||
    indexes.customerName === -1 ||
    indexes.industrialHub === -1
  ) {

    throw new Error(
      'Customers sheet is missing required columns.'
    );

  }


  const wantedHub =
    String(
      hubName || ''
    )
      .trim()
      .toLowerCase();


  const customers = [];


  for (let i = 1; i < data.length; i++) {

    const row = data[i];


    const rowHub =
      String(
        row[indexes.industrialHub] || ''
      )
        .trim()
        .toLowerCase();


    if (rowHub !== wantedHub) {

      continue;

    }


    customers.push({

      companyId:
        getCell(
          row,
          indexes.companyId
        ),

      customerName:
        getCell(
          row,
          indexes.customerName
        ),

      industrialArea:
        getCell(
          row,
          indexes.industrialArea
        ),

      industrialHub:
        getCell(
          row,
          indexes.industrialHub
        ),

      city:
        getCell(
          row,
          indexes.city
        ),

      type:
        getCell(
          row,
          indexes.type
        ),

      status:
        getCell(
          row,
          indexes.status
        ),

      potential:
        getCell(
          row,
          indexes.potential
        ),

      website:
        getCell(
          row,
          indexes.website
        ),

      notes:
        getCell(
          row,
          indexes.notes
        )

    });

  }


  return customers.sort(
    (a, b) =>
      a.customerName.localeCompare(
        b.customerName
      )
  );

}


/*************************************************
 * GET PEOPLE BY COMPANY ID
 *************************************************/

function getPeopleByCompanyId(companyId) {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const sheet =
    ss.getSheetByName(PEOPLE_SHEET);


  if (!sheet) {

    throw new Error(
      'People sheet not found.'
    );

  }


  const data =
    sheet
      .getDataRange()
      .getDisplayValues();


  if (data.length < 2) {

    return [];

  }


  const headers =
    data[0].map(normalizeHeader);


  const indexes = {

    peopleId:
      findHeader(
        headers,
        [
          'people id',
          'peopleid'
        ]
      ),

    companyId:
      findHeader(
        headers,
        [
          'company id',
          'companyid'
        ]
      ),

    company:
      findHeader(
        headers,
        ['company']
      ),

    name:
      findHeader(
        headers,
        ['name']
      ),

    designation:
      findHeader(
        headers,
        ['designation']
      ),

    phone:
      findHeader(
        headers,
        ['phone']
      ),

    email:
      findHeader(
        headers,
        ['email']
      ),

    photo:
      findHeader(
        headers,
        ['photo']
      )

  };


  if (
    indexes.companyId === -1 ||
    indexes.name === -1
  ) {

    throw new Error(
      'People sheet is missing required columns.'
    );

  }


  const wantedId =
    String(
      companyId || ''
    )
      .trim()
      .toLowerCase();


  const people = [];


  for (let i = 1; i < data.length; i++) {

    const row = data[i];


    const rowCompanyId =
      String(
        row[indexes.companyId] || ''
      )
        .trim()
        .toLowerCase();


    if (
      rowCompanyId !== wantedId
    ) {

      continue;

    }


    const photoName =
      getCell(
        row,
        indexes.photo
      );


    people.push({

      peopleId:
        getCell(
          row,
          indexes.peopleId
        ),

      companyId:
        getCell(
          row,
          indexes.companyId
        ),

      company:
        getCell(
          row,
          indexes.company
        ),

      name:
        getCell(
          row,
          indexes.name
        ),

      designation:
        getCell(
          row,
          indexes.designation
        ),

      phone:
        getCell(
          row,
          indexes.phone
        ),

      email:
        getCell(
          row,
          indexes.email
        ),

      photo:
        photoName,

      visitingCardUrl:
        getDriveFileUrl(
          photoName
        )

    });

  }


  return people;

}


/*************************************************
 * GET DRIVE FILE URL
 *************************************************/

function getDriveFileUrl(fileName) {

  if (!fileName) {

    return '';

  }


  try {

    const files =
      DriveApp.getFilesByName(
        fileName
      );


    if (files.hasNext()) {

      const file =
        files.next();


      return file.getUrl();

    }

  } catch (error) {

    return '';

  }


  return '';

}


/*************************************************
 * CELL VALUE
 *************************************************/

function getCell(row, index) {

  if (
    index === -1 ||
    index >= row.length
  ) {

    return '';

  }


  return String(
    row[index] || ''
  ).trim();

}


/*************************************************
 * NORMALIZE HEADER
 *************************************************/

function normalizeHeader(value) {

  return String(value || '')
    .trim()
    .toLowerCase()
    .replace(/\s+/g, ' ');

}


/*************************************************
 * FIND HEADER
 *************************************************/

function findHeader(headers, names) {

  for (
    let i = 0;
    i < headers.length;
    i++
  ) {

    if (
      names.includes(
        headers[i]
      )
    ) {

      return i;

    }

  }


  return -1;

}

function getSpreadsheetUrl() {

  return SpreadsheetApp
    .getActiveSpreadsheet()
    .getUrl();

}

function updateCustomerRecord(companyId, data) {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const sheet =
    ss.getSheetByName('Customers');

  if (!sheet) {
    throw new Error(
      'Customers sheet not found.'
    );
  }


  const values =
    sheet.getDataRange().getValues();

  const headers =
    values[0].map(function(h) {

      return String(h)
        .trim()
        .toLowerCase();

    });


  const idColumn =
    headers.indexOf('company id');


  if (idColumn === -1) {

    throw new Error(
      'Company ID column not found.'
    );

  }


  let row = -1;


  for (
    let i = 1;
    i < values.length;
    i++
  ) {

    if (
      String(
        values[i][idColumn]
      ).trim() ===
      String(companyId).trim()
    ) {

      row = i + 1;

      break;

    }

  }


  if (row === -1) {

    throw new Error(
      'Customer not found.'
    );

  }


  const fields = {

    'customer name':
      data.customerName,

    'industrial area':
      data.industrialArea,

    'industrial hub':
      data.industrialHub,

    'city':
      data.city,

    'type':
      data.type,

    'status':
      data.status,

    'potential':
      data.potential,

    'website':
      data.website,

    'notes':
      data.notes

  };


  Object.keys(fields)
    .forEach(function(header) {

      const column =
        headers.indexOf(header);


      if (column !== -1) {

        sheet
          .getRange(
            row,
            column + 1
          )
          .setValue(
            fields[header]
          );

      }

    });


  return true;

}

function getCustomerById(companyId) {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const sheet =
    ss.getSheetByName('Customers');


  const values =
    sheet.getDataRange().getValues();


  const headers =
    values[0];


  const idColumn =
    headers.findIndex(function(h) {

      return String(h)
        .trim()
        .toLowerCase() ===
        'company id';

    });


  for (
    let i = 1;
    i < values.length;
    i++
  ) {

    if (
      String(
        values[i][idColumn]
      ).trim() ===
      String(companyId).trim()
    ) {

      const obj = {};

      headers.forEach(
        function(header, index) {

          obj[
            String(header)
              .trim()
          ] =
            values[i][index];

        }
      );

      return obj;

    }

  }


  return null;

}

function savePersonRecord(
  peopleId,
  companyId,
  data
) {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const sheet =
    ss.getSheetByName('People');

  if (!sheet) {

    throw new Error(
      'People sheet not found.'
    );

  }


  const values =
    sheet.getDataRange().getValues();


  const headers =
    values[0].map(function(h) {

      return String(h)
        .trim()
        .toLowerCase();

    });


  const idColumn =
    headers.indexOf(
      'people id'
    );


  if (idColumn === -1) {

    throw new Error(
      'People ID column not found.'
    );

  }


  let row = -1;


  /*
   * EDIT EXISTING PERSON
   */

  if (peopleId) {

    for (
      let i = 1;
      i < values.length;
      i++
    ) {

      if (
        String(
          values[i][idColumn]
        ).trim() ===
        String(peopleId).trim()
      ) {

        row = i + 1;

        break;

      }

    }

  }


  /*
   * ADD NEW PERSON
   */

  if (row === -1) {

    row =
      sheet.getLastRow() + 1;


    peopleId =
      generateNextPeopleId(
        values,
        idColumn
      );


    sheet
      .getRange(
        row,
        idColumn + 1
      )
      .setValue(
        peopleId
      );


    const companyColumn =
      headers.indexOf(
        'company id'
      );


    if (companyColumn !== -1) {

      sheet
        .getRange(
          row,
          companyColumn + 1
        )
        .setValue(
          companyId
        );

    }

  }


  const fields = {

    'name':
      data.name,

    'designation':
      data.designation,

    'phone':
      data.phone,

    'email':
      data.email,

    'photo':
      data.photo

  };


  Object.keys(fields)
    .forEach(function(header) {

      const column =
        headers.indexOf(header);


      if (column !== -1) {

        sheet
          .getRange(
            row,
            column + 1
          )
          .setValue(
            fields[header]
          );

      }

    });


  return true;

}

function generateNextPeopleId(
  values,
  idColumn
) {

  let max = 0;


  for (
    let i = 1;
    i < values.length;
    i++
  ) {

    const value =
      String(
        values[i][idColumn] || ''
      ).trim();


    const match =
      value.match(
        /^P(\d+)$/i
      );


    if (match) {

      max =
        Math.max(
          max,
          Number(match[1])
        );

    }

  }


  return 'P' +
    String(
      max + 1
    ).padStart(4, '0');

}

function getPersonById(peopleId) {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const sheet =
    ss.getSheetByName('People');


  const values =
    sheet.getDataRange().getValues();


  const headers =
    values[0];


  const idColumn =
    headers.findIndex(function(h) {

      return String(h)
        .trim()
        .toLowerCase() ===
        'people id';

    });


  for (
    let i = 1;
    i < values.length;
    i++
  ) {

    if (
      String(
        values[i][idColumn]
      ).trim() ===
      String(peopleId).trim()
    ) {

      const obj = {};

      headers.forEach(
        function(header, index) {

          obj[
            String(header).trim()
          ] =
            values[i][index];

        }
      );

      return obj;

    }

  }


  return null;

}

/*************************************************
 * GET SETTINGS DROPDOWN VALUES
 *************************************************/

function getSettingsOptions() {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const sheet =
    ss.getSheetByName('Settings');

  if (!sheet) {
    throw new Error(
      'Settings sheet not found.'
    );
  }

  const values =
    sheet.getDataRange().getValues();

  if (!values.length) {
    return {};
  }

  const result = {};

  /*
   * Settings sheet has the following headers:
   *
   * Industrial Area
   * City
   * Industrial Hub
   * Designation
   * Type
   * Status
   */

  values[0].forEach(
    function(header, columnIndex) {

      const key =
        String(header)
          .trim();

      if (!key) {
        return;
      }

      const list = [];

      for (
        let row = 1;
        row < values.length;
        row++
      ) {

        const value =
          String(
            values[row][columnIndex] || ''
          ).trim();

        if (
          value &&
          list.indexOf(value) === -1
        ) {

          list.push(value);

        }

      }

      result[key] = list;

    }
  );

  return result;
}