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
  getGitHubImageUrl(
    photoName
  )

    });

  }


  return people;

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

/*************************************************
 * SAVE PERSON
 * GOOGLE SHEET + GITHUB VISITING CARD
 *************************************************/

function savePersonRecord(
  peopleId,
  companyId,
  data
) {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const sheet =
    ss.getSheetByName(PEOPLE_SHEET);

  if (!sheet) {

    throw new Error(
      'People sheet not found.'
    );

  }


  const values =
    sheet
      .getDataRange()
      .getValues();


  if (!values.length) {

    throw new Error(
      'People sheet is empty.'
    );

  }


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
        String(
          peopleId
        ).trim()
      ) {

        row =
          i + 1;

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

  }


  /*
   * GET COMPANY NAME
   */

  const companyName =
    getCompanyNameById(
      companyId
    );


  /*
   * GET EXISTING PHOTO
   */

  const photoColumn =
    headers.indexOf(
      'photo'
    );


  let oldPhoto =
    '';


  if (
    photoColumn !== -1 &&
    row <= sheet.getLastRow()
  ) {

    oldPhoto =
      String(
        sheet
          .getRange(
            row,
            photoColumn + 1
          )
          .getValue() ||
        ''
      ).trim();

  }


  /*
   * HANDLE NEW VISITING CARD
   */

  let finalPhoto =
    oldPhoto;


  if (
    data.photoFile &&
    data.photoFile.base64
  ) {

    finalPhoto =
      uploadVisitingCardToGitHub(
        data.name,
        data.photoFile,
        oldPhoto
      );

  }


  /*
   * SAVE SHEET FIELDS
   */

  const fields = {

    'company id':
      companyId,

    'company':
      companyName,

    'name':
      data.name,

    'designation':
      data.designation,

    'phone':
      data.phone,

    'email':
      data.email,

    'photo':
      finalPhoto

  };


  Object.keys(fields)
    .forEach(function(header) {

      const column =
        headers.indexOf(
          header
        );


      if (
        column !== -1
      ) {

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

/*************************************************
 * GET COMPANY NAME BY COMPANY ID
 *************************************************/

function getCompanyNameById(
  companyId
) {

  const ss =
    SpreadsheetApp
      .getActiveSpreadsheet();


  const sheet =
    ss.getSheetByName(
      CUSTOMERS_SHEET
    );


  if (!sheet) {

    throw new Error(
      'Customers sheet not found.'
    );

  }


  const data =
    sheet
      .getDataRange()
      .getDisplayValues();


  if (
    data.length < 2
  ) {

    return '';

  }


  const headers =
    data[0].map(
      normalizeHeader
    );


  const idIndex =
    findHeader(
      headers,
      [
        'company id',
        'companyid'
      ]
    );


  const nameIndex =
    findHeader(
      headers,
      [
        'customer name',
        'customername'
      ]
    );


  if (
    idIndex === -1 ||
    nameIndex === -1
  ) {

    throw new Error(
      'Customers sheet must contain Company ID and Customer Name.'
    );

  }


  const wantedId =
    String(
      companyId || ''
    )
      .trim()
      .toLowerCase();


  for (
    let i = 1;
    i < data.length;
    i++
  ) {

    const rowId =
      String(
        data[i][idIndex] || ''
      )
      .trim()
      .toLowerCase();


    if (
      rowId === wantedId
    ) {

      return String(
        data[i][nameIndex] || ''
      ).trim();

    }

  }


  return '';

}

/*************************************************
 * SAVE VISITING CARD IMAGE
 *************************************************/

function saveVisitingCardImage(
  peopleId,
  fileData,
  oldPhotoName
) {

  if (
    !fileData ||
    !fileData.base64
  ) {

    return oldPhotoName || '';

  }


  /*
   * Get / create dedicated folder.
   */

  const folder =
    getVisitingCardFolder();


  /*
   * Remove old file when replacing.
   */

  if (oldPhotoName) {

    try {

      const oldFiles =
        folder.getFilesByName(
          oldPhotoName
        );


      while (
        oldFiles.hasNext()
      ) {

        const oldFile =
          oldFiles.next();


        oldFile.setTrashed(
          true
        );

      }

    } catch (error) {

      console.log(
        'Old visiting card could not be removed:',
        error
      );

    }

  }


  /*
   * Determine extension.
   */

  const originalName =
    String(
      fileData.fileName ||
      'VisitingCard'
    );


  let extension =
    'jpg';


  const match =
    originalName.match(
      /\.([a-zA-Z0-9]+)$/
    );


  if (match) {

    extension =
      match[1].toLowerCase();

  }


  /*
   * Clean People ID.
   */

  const safePeopleId =
    String(
      peopleId
    )
      .replace(
        /[^a-zA-Z0-9_-]/g,
        ''
      );


  /*
   * Generated filename.
   *
   * Example:
   * P0001_VisitingCard.jpg
   */

  const fileName =
    safePeopleId +
    '_VisitingCard.' +
    extension;


  /*
   * Convert Base64 to blob.
   */

  const bytes =
    Utilities
      .base64Decode(
        fileData.base64
      );


  const blob =
    Utilities
      .newBlob(
        bytes,
        fileData.mimeType ||
          'image/jpeg',
        fileName
      );


  /*
   * Create Drive file.
   */

  const file =
    folder.createFile(
      blob
    );


  /*
   * Return ONLY filename.
   * This goes into People > Photo.
   */

  return file.getName();

}

/*************************************************
 * VISITING CARD FOLDER
 *************************************************/

function getVisitingCardFolder() {

  const folderName =
    'Customers Hub - Visiting Cards';


  const folders =
    DriveApp.getFoldersByName(
      folderName
    );


  if (
    folders.hasNext()
  ) {

    return folders.next();

  }


  return DriveApp.createFolder(
    folderName
  );

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

/*************************************************
 * GITHUB TOKEN
 * API Key sheet A3
 *************************************************/

function getGitHubToken() {

  const ss =
    SpreadsheetApp
      .getActiveSpreadsheet();


  const sheet =
    ss.getSheetByName(
      API_KEY_SHEET
    );


  if (!sheet) {

    throw new Error(
      'API Key sheet not found.'
    );

  }


  const token =
    String(
      sheet
        .getRange('A3')
        .getDisplayValue() ||
      ''
    ).trim();


  if (!token) {

    throw new Error(
      'GitHub token not found in API Key!A3.'
    );

  }


  return token;

}

/*************************************************
 * GITHUB SETTINGS
 *************************************************/

const GITHUB_OWNER =
  'rajinikanthhc';

const GITHUB_REPO =
  'images';

const GITHUB_BRANCH =
  'main';

const GITHUB_FOLDER =
  'visiting-cards';


/*************************************************
 * UPLOAD / UPDATE VISITING CARD
 *************************************************/

function uploadVisitingCardToGitHub(
  contactName,
  fileData,
  oldPhotoName
) {

  if (
    !contactName
  ) {

    throw new Error(
      'Contact name is required before uploading the visiting card.'
    );

  }


  if (
    !fileData ||
    !fileData.base64
  ) {

    return oldPhotoName || '';

  }


  const token =
    getGitHubToken();


  /*
   * Clean contact name.
   */

  const cleanName =
    String(
      contactName
    )
      .trim()
      .replace(
        /[\\/:*?"<>|]/g,
        ''
      )
      .replace(
        /\s+/g,
        ' '
      );


  if (!cleanName) {

    throw new Error(
      'Invalid contact name.'
    );

  }


  /*
   * Get extension from uploaded file.
   */

  let extension =
    'png';


  const originalName =
    String(
      fileData.fileName ||
      ''
    );


  const match =
    originalName.match(
      /\.([a-zA-Z0-9]+)$/
    );


  if (match) {

    extension =
      match[1]
        .toLowerCase();

  }


  /*
   * Keep only common image extensions.
   */

  const allowed =
    [
      'jpg',
      'jpeg',
      'png',
      'webp'
    ];


  if (
    allowed.indexOf(
      extension
    ) === -1
  ) {

    extension =
      'png';

  }


  /*
   * Final filename.
   *
   * Example:
   * Rajinikanth H C.png
   */

  const newFileName =
    cleanName +
    '.' +
    extension;


  /*
   * If an old filename exists
   * and it is different, remove it.
   */

  if (
    oldPhotoName &&
    oldPhotoName !==
      newFileName
  ) {

    deleteGitHubFile(
      oldPhotoName,
      token
    );

  }


  /*
   * Upload / update new file.
   */

  const existing =
    getGitHubFile(
      newFileName,
      token
    );


  const url =
    getGitHubContentsUrl(
      newFileName
    );


  const payload = {

    message:
      existing
        ? 'Update visiting card - ' +
          newFileName
        : 'Add visiting card - ' +
          newFileName,

    content:
      fileData.base64,

    branch:
      GITHUB_BRANCH

  };


  if (
    existing &&
    existing.sha
  ) {

    payload.sha =
      existing.sha;

  }


  const response =
    UrlFetchApp.fetch(
      url,
      {

        method:
          'put',

        contentType:
          'application/json',

        headers: {

          Authorization:
            'Bearer ' +
            token,

          Accept:
            'application/vnd.github+json',

          'X-GitHub-Api-Version':
            '2022-11-28'

        },

        payload:
          JSON.stringify(
            payload
          ),

        muteHttpExceptions:
          true

      }
    );


  const code =
    response.getResponseCode();


  const body =
    response.getContentText();


  if (
    code !== 200 &&
    code !== 201
  ) {

    throw new Error(
      'GitHub upload failed (' +
      code +
      '): ' +
      body
    );

  }


  return newFileName;

}

/*************************************************
 * GITHUB CONTENTS URL
 *************************************************/

function getGitHubContentsUrl(
  fileName
) {

  return (
    'https://api.github.com/repos/' +
    GITHUB_OWNER +
    '/' +
    GITHUB_REPO +
    '/contents/' +
    GITHUB_FOLDER +
    '/' +
    encodeURIComponent(
      fileName
    )
  );

}


/*************************************************
 * GET EXISTING GITHUB FILE
 *************************************************/

function getGitHubFile(
  fileName,
  token
) {

  const url =
    getGitHubContentsUrl(
      fileName
    );


  const response =
    UrlFetchApp.fetch(
      url,
      {

        method:
          'get',

        headers: {

          Authorization:
            'Bearer ' +
            token,

          Accept:
            'application/vnd.github+json',

          'X-GitHub-Api-Version':
            '2022-11-28'

        },

        muteHttpExceptions:
          true

      }
    );


  const code =
    response.getResponseCode();


  if (
    code === 404
  ) {

    return null;

  }


  if (
    code !== 200
  ) {

    throw new Error(
      'Unable to check GitHub file: ' +
      response.getContentText()
    );

  }


  return JSON.parse(
    response.getContentText()
  );

}


/*************************************************
 * DELETE OLD GITHUB FILE
 *************************************************/

function deleteGitHubFile(
  fileName,
  token
) {

  if (!fileName) {

    return;

  }


  const existing =
    getGitHubFile(
      fileName,
      token
    );


  if (
    !existing ||
    !existing.sha
  ) {

    return;

  }


  const url =
    getGitHubContentsUrl(
      fileName
    );


  const payload = {

    message:
      'Remove old visiting card - ' +
      fileName,

    sha:
      existing.sha,

    branch:
      GITHUB_BRANCH

  };


  const response =
    UrlFetchApp.fetch(
      url,
      {

        method:
          'delete',

        contentType:
          'application/json',

        headers: {

          Authorization:
            'Bearer ' +
            token,

          Accept:
            'application/vnd.github+json',

          'X-GitHub-Api-Version':
            '2022-11-28'

        },

        payload:
          JSON.stringify(
            payload
          ),

        muteHttpExceptions:
          true

      }
    );


  const code =
    response.getResponseCode();


  if (
    code !== 200
  ) {

    throw new Error(
      'Unable to delete old GitHub visiting card: ' +
      response.getContentText()
    );

  }

}


/*************************************************
 * GITHUB RAW IMAGE URL
 *************************************************/

function getGitHubImageUrl(
  fileName
) {

  if (!fileName) {

    return '';

  }


  return (
    'https://raw.githubusercontent.com/' +
    GITHUB_OWNER +
    '/' +
    GITHUB_REPO +
    '/' +
    GITHUB_BRANCH +
    '/' +
    GITHUB_FOLDER +
    '/' +
    encodeURIComponent(
      fileName
    )
  );

}

/*************************************************
 * GET LOWEST AVAILABLE COMPANY ID
 *************************************************/

function getNextAvailableCompanyId() {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const sheet =
    ss.getSheetByName(
      CUSTOMERS_SHEET
    );


  if (!sheet) {

    throw new Error(
      'Customers sheet not found.'
    );

  }


  const values =
    sheet
      .getDataRange()
      .getValues();


  if (!values.length) {

    return 'C0001';

  }


  const headers =
    values[0];


  const idColumn =
    headers.findIndex(
      function (header) {

        return String(header)
          .trim()
          .toLowerCase() ===
          'company id';

      }
    );


  if (idColumn === -1) {

    throw new Error(
      'Company ID column not found.'
    );

  }


  const usedIds =
    {};


  for (
    let i = 1;
    i < values.length;
    i++
  ) {

    const value =
      String(
        values[i][idColumn] || ''
      )
      .trim()
      .toUpperCase();


    const match =
      value.match(
        /^C(\d+)$/
      );


    if (match) {

      usedIds[
        Number(match[1])
      ] = true;

    }

  }


  let number = 1;


  while (
    usedIds[number]
  ) {

    number++;

  }


  return 'C' +
    String(number)
      .padStart(
        4,
        '0'
      );

}


/*************************************************
 * SAVE NEW CUSTOMER
 *************************************************/

function saveNewCustomerRecord(
  data
) {

  const lock =
    LockService.getScriptLock();


  lock.waitLock(
    30000
  );


  try {

    const ss =
      SpreadsheetApp
        .getActiveSpreadsheet();


    const sheet =
      ss.getSheetByName(
        CUSTOMERS_SHEET
      );


    if (!sheet) {

      throw new Error(
        'Customers sheet not found.'
      );

    }


    if (
      !data ||
      !data.customerName
    ) {

      throw new Error(
        'Customer Name is required.'
      );

    }


    const values =
      sheet
        .getDataRange()
        .getValues();


    if (!values.length) {

      throw new Error(
        'Customers sheet is empty.'
      );

    }


    const headers =
      values[0].map(
        function (header) {

          return String(header)
            .trim()
            .toLowerCase();

        }
      );


    const idColumn =
      headers.indexOf(
        'company id'
      );


    if (idColumn === -1) {

      throw new Error(
        'Company ID column not found.'
      );

    }


    /*
     * Generate again while locked.
     *
     * This guarantees that two users
     * cannot accidentally get the same ID.
     */

    const companyId =
      getNextAvailableCompanyId();


    /*
     * Create a blank row based on
     * the existing sheet columns.
     */

    const newRow =
      new Array(
        headers.length
      )
      .fill('');


    /*
     * Helper to safely set
     * a column by header.
     */

    function setField(
      header,
      value
    ) {

      const column =
        headers.indexOf(
          header
        );


      if (
        column !== -1
      ) {

        newRow[column] =
          value || '';

      }

    }


    setField(
      'company id',
      companyId
    );


    setField(
      'customer name',
      data.customerName
    );


    setField(
      'industrial area',
      data.industrialArea
    );


    setField(
      'industrial hub',
      data.industrialHub
    );


    setField(
      'city',
      data.city
    );


    setField(
      'type',
      data.type
    );


    setField(
      'status',
      data.status
    );


    setField(
      'potential',
      data.potential
    );


    setField(
      'website',
      data.website
    );


    setField(
      'notes',
      data.notes
    );


    /*
     * Add customer.
     */

    sheet
      .appendRow(
        newRow
      );


    return companyId;


  } finally {

    lock.releaseLock();

  }

}

/*************************************************
 * UNIVERSAL SEARCH
 *************************************************/

function universalSearch(
  query
) {

  const search =
    String(
      query || ''
    )
    .trim()
    .toLowerCase();


  if (!search) {

    return [];

  }


  const results = [];


  /*
   * INDUSTRIAL HUBS
   */

  const hubs =
    getIndustrialHubs();


  hubs.forEach(
    function (hub) {

      const text =
        [
          hub.hub,
          hub.city
        ]
        .join(' ')
        .toLowerCase();


      if (
        text.indexOf(
          search
        ) !== -1
      ) {

        results.push({

          type:
            'hub',

          name:
            hub.hub,

          city:
            hub.city

        });

      }

    }
  );


  /*
   * CUSTOMERS
   */

  const customerSheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(
        CUSTOMERS_SHEET
      );


  if (customerSheet) {

    const data =
      customerSheet
        .getDataRange()
        .getDisplayValues();


    if (
      data.length > 1
    ) {

      const headers =
        data[0].map(
          normalizeHeader
        );


      const companyIdIndex =
        findHeader(
          headers,
          [
            'company id'
          ]
        );


      const nameIndex =
        findHeader(
          headers,
          [
            'customer name'
          ]
        );


      const areaIndex =
        findHeader(
          headers,
          [
            'industrial area'
          ]
        );


      const hubIndex =
        findHeader(
          headers,
          [
            'industrial hub'
          ]
        );


      const cityIndex =
        findHeader(
          headers,
          [
            'city'
          ]
        );


      const typeIndex =
        findHeader(
          headers,
          [
            'type'
          ]
        );


      const statusIndex =
        findHeader(
          headers,
          [
            'status'
          ]
        );


      const potentialIndex =
        findHeader(
          headers,
          [
            'potential'
          ]
        );


      const websiteIndex =
        findHeader(
          headers,
          [
            'website'
          ]
        );


      const notesIndex =
        findHeader(
          headers,
          [
            'notes'
          ]
        );


      for (
        let i = 1;
        i < data.length;
        i++
      ) {

        const row =
          data[i];


        const companyId =
          getCell(
            row,
            companyIdIndex
          );


        const name =
          getCell(
            row,
            nameIndex
          );


        const area =
          getCell(
            row,
            areaIndex
          );


        const hub =
          getCell(
            row,
            hubIndex
          );


        const city =
          getCell(
            row,
            cityIndex
          );


        const type =
          getCell(
            row,
            typeIndex
          );


        const status =
          getCell(
            row,
            statusIndex
          );


        const potential =
          getCell(
            row,
            potentialIndex
          );


        const website =
          getCell(
            row,
            websiteIndex
          );


        const notes =
          getCell(
            row,
            notesIndex
          );


        const searchable =
          [
            companyId,
            name,
            area,
            hub,
            city,
            type,
            status,
            potential,
            website,
            notes
          ]
          .join(' ')
          .toLowerCase();


        if (
          searchable.indexOf(
            search
          ) !== -1
        ) {

          results.push({

            type:
              'customer',

            name:
              name,

            companyId:
              companyId,

            industrialArea:
              area,

            industrialHub:
              hub,

            city:
              city

          });

        }

      }

    }

  }


  /*
   * PEOPLE
   */

  const peopleSheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(
        PEOPLE_SHEET
      );


  if (peopleSheet) {

    const data =
      peopleSheet
        .getDataRange()
        .getDisplayValues();


    if (
      data.length > 1
    ) {

      const headers =
        data[0].map(
          normalizeHeader
        );


      const peopleIdIndex =
        findHeader(
          headers,
          [
            'people id'
          ]
        );


      const companyIdIndex =
        findHeader(
          headers,
          [
            'company id'
          ]
        );


      const companyIndex =
        findHeader(
          headers,
          [
            'company'
          ]
        );


      const nameIndex =
        findHeader(
          headers,
          [
            'name'
          ]
        );


      const designationIndex =
        findHeader(
          headers,
          [
            'designation'
          ]
        );


      const phoneIndex =
        findHeader(
          headers,
          [
            'phone'
          ]
        );


      const emailIndex =
        findHeader(
          headers,
          [
            'email'
          ]
        );


      /*
       * Build Company ID → Hub
       */

      const companyHubMap =
        {};


      const customerData =
        customerSheet
          ? customerSheet
              .getDataRange()
              .getDisplayValues()
          : [];


      if (
        customerData.length > 1
      ) {

        const customerHeaders =
          customerData[0].map(
            normalizeHeader
          );


        const cId =
          findHeader(
            customerHeaders,
            [
              'company id'
            ]
          );


        const cHub =
          findHeader(
            customerHeaders,
            [
              'industrial hub'
            ]
          );


        for (
          let i = 1;
          i < customerData.length;
          i++
        ) {

          companyHubMap[
            getCell(
              customerData[i],
              cId
            ).toLowerCase()
          ] =
            getCell(
              customerData[i],
              cHub
            );

        }

      }


      for (
        let i = 1;
        i < data.length;
        i++
      ) {

        const row =
          data[i];


        const peopleId =
          getCell(
            row,
            peopleIdIndex
          );


        const companyId =
          getCell(
            row,
            companyIdIndex
          );


        const company =
          getCell(
            row,
            companyIndex
          );


        const name =
          getCell(
            row,
            nameIndex
          );


        const designation =
          getCell(
            row,
            designationIndex
          );


        const phone =
          getCell(
            row,
            phoneIndex
          );


        const email =
          getCell(
            row,
            emailIndex
          );


        const hub =
          companyHubMap[
            companyId.toLowerCase()
          ] || '';


        const searchable =
          [
            peopleId,
            companyId,
            company,
            name,
            designation,
            phone,
            email,
            hub
          ]
          .join(' ')
          .toLowerCase();


        if (
          searchable.indexOf(
            search
          ) !== -1
        ) {

          results.push({

            type:
              'person',

            name:
              name,

            peopleId:
              peopleId,

            companyId:
              companyId,

            company:
              company,

            designation:
              designation,

            phone:
              phone,

            email:
              email,

            industrialHub:
              hub

          });

        }

      }

    }

  }


  /*
   * Keep result list manageable.
   */

  return results.slice(
    0,
    30
  );

}

/*************************************************
 * SETTINGS MANAGER
 * ADD / EDIT / DELETE SETTINGS VALUES
 *************************************************/

function getSettingList(settingName) {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const sheet =
    ss.getSheetByName('Settings');

  if (!sheet) {
    throw new Error('Settings sheet not found.');
  }

  const values =
    sheet.getDataRange().getValues();

  if (!values.length) {
    throw new Error('Settings sheet is empty.');
  }

  const headers =
    values[0].map(function (h) {
      return String(h || '').trim();
    });

  const column =
    headers.findIndex(function (header) {
      return header.toLowerCase() ===
        String(settingName || '').trim().toLowerCase();
    });

  if (column === -1) {
    throw new Error(
      'Setting column not found: ' + settingName
    );
  }

  const list = [];

  for (
    let row = 1;
    row < values.length;
    row++
  ) {

    const value =
      String(
        values[row][column] || ''
      ).trim();

    if (
      value &&
      !list.some(function (item) {
        return item.toLowerCase() === value.toLowerCase();
      })
    ) {

      list.push(value);

    }

  }

  return list;
}


/*************************************************
 * ADD SETTING VALUE
 *************************************************/

function addSettingValue(
  settingName,
  value
) {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const sheet =
    ss.getSheetByName('Settings');

  if (!sheet) {
    throw new Error('Settings sheet not found.');
  }

  const cleanName =
    String(settingName || '').trim();

  const cleanValue =
    String(value || '').trim();

  if (!cleanName) {
    throw new Error('Setting name is required.');
  }

  if (!cleanValue) {
    throw new Error('Value is required.');
  }

  const values =
    sheet.getDataRange().getValues();

  const headers =
    values[0].map(function (h) {
      return String(h || '').trim();
    });

  const column =
    headers.findIndex(function (header) {
      return header.toLowerCase() ===
        cleanName.toLowerCase();
    });

  if (column === -1) {
    throw new Error(
      'Setting column not found: ' + cleanName
    );
  }

  /*
   * Prevent duplicate values.
   */

  for (
    let row = 1;
    row < values.length;
    row++
  ) {

    const existing =
      String(
        values[row][column] || ''
      ).trim();

    if (
      existing.toLowerCase() ===
      cleanValue.toLowerCase()
    ) {

      throw new Error(
        '"' + cleanValue +
        '" already exists.'
      );

    }

  }

  /*
   * Put new value in first empty cell.
   */

  let targetRow = -1;

  for (
    let row = 1;
    row < values.length;
    row++
  ) {

    const existing =
      String(
        values[row][column] || ''
      ).trim();

    if (!existing) {

      targetRow =
        row + 1;

      break;

    }

  }

  /*
   * If no empty cell exists,
   * append a new row.
   */

  if (targetRow === -1) {

    targetRow =
      sheet.getLastRow() + 1;

  }

  sheet
    .getRange(
      targetRow,
      column + 1
    )
    .setValue(
      cleanValue
    );

  return true;
}


/*************************************************
 * EDIT SETTING VALUE
 *************************************************/

function updateSettingValue(
  settingName,
  oldValue,
  newValue
) {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const sheet =
    ss.getSheetByName('Settings');

  if (!sheet) {
    throw new Error('Settings sheet not found.');
  }

  const cleanName =
    String(settingName || '').trim();

  const oldText =
    String(oldValue || '').trim();

  const newText =
    String(newValue || '').trim();

  if (!newText) {
    throw new Error('Value is required.');
  }

  const values =
    sheet.getDataRange().getValues();

  const headers =
    values[0].map(function (h) {
      return String(h || '').trim();
    });

  const column =
    headers.findIndex(function (header) {
      return header.toLowerCase() ===
        cleanName.toLowerCase();
    });

  if (column === -1) {
    throw new Error(
      'Setting column not found: ' + cleanName
    );
  }

  let foundRow = -1;

  for (
    let row = 1;
    row < values.length;
    row++
  ) {

    const existing =
      String(
        values[row][column] || ''
      ).trim();

    if (
      existing.toLowerCase() ===
      oldText.toLowerCase()
    ) {

      foundRow =
        row + 1;

      break;

    }

  }

  if (foundRow === -1) {

    throw new Error(
      'Existing value not found: ' +
      oldText
    );

  }

  /*
   * Prevent duplicate after editing.
   */

  for (
    let row = 1;
    row < values.length;
    row++
  ) {

    if (
      row + 1 === foundRow
    ) {
      continue;
    }

    const existing =
      String(
        values[row][column] || ''
      ).trim();

    if (
      existing.toLowerCase() ===
      newText.toLowerCase()
    ) {

      throw new Error(
        '"' + newText +
        '" already exists.'
      );

    }

  }

  sheet
    .getRange(
      foundRow,
      column + 1
    )
    .setValue(
      newText
    );

  return true;
}


/*************************************************
 * DELETE SETTING VALUE
 *************************************************/

function deleteSettingValue(
  settingName,
  value
) {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const sheet =
    ss.getSheetByName('Settings');

  if (!sheet) {
    throw new Error('Settings sheet not found.');
  }

  const cleanName =
    String(settingName || '').trim();

  const cleanValue =
    String(value || '').trim();

  const values =
    sheet.getDataRange().getValues();

  const headers =
    values[0].map(function (h) {
      return String(h || '').trim();
    });

  const column =
    headers.findIndex(function (header) {
      return header.toLowerCase() ===
        cleanName.toLowerCase();
    });

  if (column === -1) {
    throw new Error(
      'Setting column not found: ' + cleanName
    );
  }

  for (
    let row = 1;
    row < values.length;
    row++
  ) {

    const existing =
      String(
        values[row][column] || ''
      ).trim();

    if (
      existing.toLowerCase() ===
      cleanValue.toLowerCase()
    ) {

      /*
       * Clear only this cell.
       * Do NOT delete the entire row because
       * other Settings columns may contain data.
       */

      sheet
        .getRange(
          row + 1,
          column + 1
        )
        .clearContent();

      return true;

    }

  }

  throw new Error(
    'Value not found: ' + cleanValue
  );

}