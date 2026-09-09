/*************************************************
 * CUSTOMERS HUB
 * INDUSTRIAL AREA / CUSTOMER MANAGEMENT
 *
 * DATA RELATIONSHIP
 *
 * Industrial Area
 *       ↓
 * Customers
 *       ↓
 * People
 *************************************************/


const AREAS_SHEET = 'Areas';
const CUSTOMERS_SHEET = 'Customers';
const PEOPLE_SHEET = 'People';
const API_KEY_SHEET = 'API Key';
const SETTINGS_SHEET = 'Settings';


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

  for (
    let i = 0;
    i < sheets.length;
    i++
  ) {

    const name =
      sheets[i]
        .getName()
        .trim()
        .toLowerCase();

    if (
      name ===
      API_KEY_SHEET.toLowerCase()
    ) {

      sheet =
        sheets[i];

      break;

    }

  }

  if (!sheet) {

    throw new Error(
      'API Key sheet not found. Available sheets: ' +
      sheets
        .map(function(s) {
          return s.getName();
        })
        .join(', ')
    );

  }

  const values =
    sheet
      .getDataRange()
      .getDisplayValues();

  for (
    let r = 0;
    r < values.length;
    r++
  ) {

    for (
      let c = 0;
      c < values[r].length;
      c++
    ) {

      const value =
        String(
          values[r][c] || ''
        ).trim();

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
 * INDUSTRIAL AREAS
 *
 * Areas sheet:
 *
 * Industrial Area
 * City
 * Latitude
 * Longitude
 *************************************************/

function getIndustrialAreas() {

  const ss =
    SpreadsheetApp
      .getActiveSpreadsheet();

  const sheet =
    ss.getSheetByName(
      AREAS_SHEET
    );

  if (!sheet) {

    throw new Error(
      'Areas sheet not found.'
    );

  }

  const values =
    sheet
      .getDataRange()
      .getDisplayValues();

  if (
    values.length < 2
  ) {

    return [];

  }

  const headers =
    values[0].map(
      normalizeHeader
    );

  const areaColumn =
    findHeader(
      headers,
      [
        'industrial area',
        'area'
      ]
    );

  const cityColumn =
    findHeader(
      headers,
      [
        'city'
      ]
    );

  const latitudeColumn =
    findHeader(
      headers,
      [
        'latitude',
        'lat'
      ]
    );

  const longitudeColumn =
    findHeader(
      headers,
      [
        'longitude',
        'long',
        'lng'
      ]
    );

  if (
    areaColumn === -1 ||
    cityColumn === -1 ||
    latitudeColumn === -1 ||
    longitudeColumn === -1
  ) {

    throw new Error(
      'Areas sheet must contain Industrial Area, City, Latitude and Longitude columns.'
    );

  }

  const areas = [];

  for (
    let i = 1;
    i < values.length;
    i++
  ) {

    const row =
      values[i];

    const industrialArea =
      getCell(
        row,
        areaColumn
      );

    if (!industrialArea) {
      continue;
    }

    const city =
      getCell(
        row,
        cityColumn
      );

    const latitude =
      Number(
        row[latitudeColumn]
      );

    const longitude =
      Number(
        row[longitudeColumn]
      );

    if (
      !Number.isFinite(latitude) ||
      !Number.isFinite(longitude)
    ) {

      continue;

    }

    areas.push({

      industrialArea:
        industrialArea,

      city:
        city,

      latitude:
        latitude,

      longitude:
        longitude,

      rowNumber:
        i + 1

    });

  }

  areas.sort(
    function(a, b) {

      return String(
        a.industrialArea
      ).localeCompare(
        String(
          b.industrialArea
        )
      );

    }
  );

  return areas;

}


/*************************************************
 * GET CUSTOMERS BY INDUSTRIAL AREA
 *************************************************/

function getCustomersByArea(
  industrialArea
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

  const values =
    sheet
      .getDataRange()
      .getDisplayValues();

  if (
    values.length < 2
  ) {

    return [];

  }

  const headers =
    values[0].map(
      normalizeHeader
    );

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
          'area'
        ]
      ),

    city:
      findHeader(
        headers,
        [
          'city'
        ]
      ),

    type:
      findHeader(
        headers,
        [
          'type'
        ]
      ),

    status:
      findHeader(
        headers,
        [
          'status'
        ]
      ),

    potential:
      findHeader(
        headers,
        [
          'potential'
        ]
      ),

    website:
      findHeader(
        headers,
        [
          'website'
        ]
      ),

    notes:
      findHeader(
        headers,
        [
          'notes'
        ]
      )

  };

  if (
    indexes.companyId === -1 ||
    indexes.customerName === -1 ||
    indexes.industrialArea === -1
  ) {

    throw new Error(
      'Customers sheet must contain Company ID, Customer Name and Industrial Area columns.'
    );

  }

  const wantedArea =
    String(
      industrialArea || ''
    )
      .trim()
      .toLowerCase();

  const customers = [];

  for (
    let i = 1;
    i < values.length;
    i++
  ) {

    const row =
      values[i];

    const rowArea =
      getCell(
        row,
        indexes.industrialArea
      )
      .toLowerCase();

    if (
      rowArea !== wantedArea
    ) {

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
    function(a, b) {

      return String(
        a.customerName || ''
      ).localeCompare(
        String(
          b.customerName || ''
        )
      );

    }
  );

}


/*************************************************
 * GET ALL CUSTOMERS
 *************************************************/

function getAllCustomers() {

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

  const values =
    sheet
      .getDataRange()
      .getDisplayValues();

  if (
    values.length < 2
  ) {

    return [];

  }

  const headers =
    values[0].map(
      normalizeHeader
    );

  const indexes = {

    companyId:
      findHeader(
        headers,
        [
          'company id'
        ]
      ),

    customerName:
      findHeader(
        headers,
        [
          'customer name'
        ]
      ),

    industrialArea:
      findHeader(
        headers,
        [
          'industrial area'
        ]
      ),

    city:
      findHeader(
        headers,
        [
          'city'
        ]
      ),

    type:
      findHeader(
        headers,
        [
          'type'
        ]
      ),

    status:
      findHeader(
        headers,
        [
          'status'
        ]
      ),

    potential:
      findHeader(
        headers,
        [
          'potential'
        ]
      ),

    website:
      findHeader(
        headers,
        [
          'website'
        ]
      ),

    notes:
      findHeader(
        headers,
        [
          'notes'
        ]
      )

  };

  const customers = [];

  for (
    let i = 1;
    i < values.length;
    i++
  ) {

    const row =
      values[i];

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
    function(a, b) {

      return String(
        a.customerName || ''
      ).localeCompare(
        String(
          b.customerName || ''
        )
      );

    }
  );

}

/*************************************************
 * GET COMPLETE CUSTOMERS HUB DATA
 *
 * Loads all core data in one backend request:
 *
 * - Industrial Areas
 * - Customers
 * - People / Contacts
 *************************************************/

function getCustomersHubData() {

  const areas =
    getIndustrialAreas();


  const customers =
    getAllCustomers();


  const ss =
    SpreadsheetApp
      .getActiveSpreadsheet();


  const peopleSheet =
    ss.getSheetByName(
      PEOPLE_SHEET
    );


  const people = [];


  if (peopleSheet) {

    const values =
      peopleSheet
        .getDataRange()
        .getDisplayValues();


    if (
      values.length > 1
    ) {

      const headers =
        values[0].map(
          normalizeHeader
        );


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
            [
              'company'
            ]
          ),

        name:
          findHeader(
            headers,
            [
              'name'
            ]
          ),

        designation:
          findHeader(
            headers,
            [
              'designation'
            ]
          ),

        phone:
          findHeader(
            headers,
            [
              'phone'
            ]
          ),

        email:
          findHeader(
            headers,
            [
              'email'
            ]
          ),

        photo:
          findHeader(
            headers,
            [
              'photo'
            ]
          )

      };


      for (
        let i = 1;
        i < values.length;
        i++
      ) {

        const row =
          values[i];


        const photo =
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
            photo,

          visitingCardUrl:
            getGitHubImageUrl(
              photo
            )

        });

      }


      people.sort(
        function(a, b) {

          return String(
            a.name || ''
          ).localeCompare(
            String(
              b.name || ''
            )
          );

        }
      );

    }

  }


  return {

    areas:
      areas,

    customers:
      customers,

    people:
      people

  };

}


/*************************************************
 * GET CUSTOMER BY ID
 *************************************************/

function getCustomerById(
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

  const values =
    sheet
      .getDataRange()
      .getValues();

  if (
    values.length < 2
  ) {

    return null;

  }

  const headers =
    values[0];

  const idColumn =
    headers.findIndex(
      function(header) {

        return normalizeHeader(
          header
        ) === 'company id';

      }
    );

  if (
    idColumn === -1
  ) {

    throw new Error(
      'Company ID column not found.'
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
    i < values.length;
    i++
  ) {

    const rowId =
      String(
        values[i][idColumn] || ''
      )
      .trim()
      .toLowerCase();

    if (
      rowId === wantedId
    ) {

      const result = {};

      headers.forEach(
        function(header, index) {

          result[
            String(
              header
            ).trim()
          ] =
            values[i][index];

        }
      );

      return result;

    }

  }

  return null;

}


/*************************************************
 * GET PEOPLE BY COMPANY
 *************************************************/

function getPeopleByCompanyId(
  companyId
) {

  const ss =
    SpreadsheetApp
      .getActiveSpreadsheet();

  const sheet =
    ss.getSheetByName(
      PEOPLE_SHEET
    );

  if (!sheet) {

    throw new Error(
      'People sheet not found.'
    );

  }

  const values =
    sheet
      .getDataRange()
      .getDisplayValues();

  if (
    values.length < 2
  ) {

    return [];

  }

  const headers =
    values[0].map(
      normalizeHeader
    );

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
        [
          'company'
        ]
      ),

    name:
      findHeader(
        headers,
        [
          'name'
        ]
      ),

    designation:
      findHeader(
        headers,
        [
          'designation'
        ]
      ),

    phone:
      findHeader(
        headers,
        [
          'phone'
        ]
      ),

    email:
      findHeader(
        headers,
        [
          'email'
        ]
      ),

    photo:
      findHeader(
        headers,
        [
          'photo'
        ]
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

  for (
    let i = 1;
    i < values.length;
    i++
  ) {

    const row =
      values[i];

    const rowCompanyId =
      getCell(
        row,
        indexes.companyId
      )
      .toLowerCase();

    if (
      rowCompanyId !== wantedId
    ) {

      continue;

    }

    const photo =
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
        photo,

      visitingCardUrl:
        getGitHubImageUrl(
          photo
        )

    });

  }

  /*
   * PEOPLE A-Z
   */

  return people.sort(
    function(a, b) {

      return String(
        a.name || ''
      ).localeCompare(
        String(
          b.name || ''
        )
      );

    }
  );

}


/*************************************************
 * SAVE NEW CUSTOMER
 *************************************************/

function saveNewCustomerRecord(
  data
) {

  const lock =
    LockService
      .getScriptLock();

  lock.waitLock(30000);

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
      !String(
        data.customerName || ''
      ).trim()
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
        normalizeHeader
      );

    const idColumn =
      headers.indexOf(
        'company id'
      );

    if (
      idColumn === -1
    ) {

      throw new Error(
        'Company ID column not found.'
      );

    }

    const companyId =
      getNextAvailableCompanyId();

    const newRow =
      new Array(
        headers.length
      ).fill('');

    setRowField(
      newRow,
      headers,
      'company id',
      companyId
    );

    setRowField(
      newRow,
      headers,
      'customer name',
      data.customerName
    );

    setRowField(
      newRow,
      headers,
      'industrial area',
      data.industrialArea
    );

    setRowField(
      newRow,
      headers,
      'city',
      data.city
    );

    setRowField(
      newRow,
      headers,
      'type',
      data.type
    );

    setRowField(
      newRow,
      headers,
      'status',
      data.status
    );

    setRowField(
      newRow,
      headers,
      'potential',
      data.potential
    );

    setRowField(
      newRow,
      headers,
      'website',
      data.website
    );

    setRowField(
      newRow,
      headers,
      'notes',
      data.notes
    );

    sheet.appendRow(
      newRow
    );

    return companyId;

  } finally {

    lock.releaseLock();

  }

}


/*************************************************
 * UPDATE CUSTOMER
 *************************************************/

function updateCustomerRecord(
  companyId,
  data
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

  const values =
    sheet
      .getDataRange()
      .getValues();

  if (
    values.length < 2
  ) {

    throw new Error(
      'Customers sheet is empty.'
    );

  }

  const headers =
    values[0].map(
      normalizeHeader
    );

  const idColumn =
    headers.indexOf(
      'company id'
    );

  if (
    idColumn === -1
  ) {

    throw new Error(
      'Company ID column not found.'
    );

  }

  const wantedId =
    String(
      companyId || ''
    )
      .trim()
      .toLowerCase();

  let rowNumber =
    -1;

  for (
    let i = 1;
    i < values.length;
    i++
  ) {

    const rowId =
      String(
        values[i][idColumn] || ''
      )
      .trim()
      .toLowerCase();

    if (
      rowId === wantedId
    ) {

      rowNumber =
        i + 1;

      break;

    }

  }

  if (
    rowNumber === -1
  ) {

    throw new Error(
      'Customer not found.'
    );

  }

  const fields = {

    'customer name':
      data.customerName,

    'industrial area':
      data.industrialArea,

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
    .forEach(
      function(header) {

        const column =
          headers.indexOf(
            header
          );

        if (
          column === -1
        ) {

          return;

        }

        sheet
          .getRange(
            rowNumber,
            column + 1
          )
          .setValue(
            fields[header] || ''
          );

      }
    );

  return true;

}


/*************************************************
 * GET NEXT AVAILABLE COMPANY ID
 *************************************************/

function getNextAvailableCompanyId() {

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

  const values =
    sheet
      .getDataRange()
      .getValues();

  if (
    !values.length
  ) {

    return 'C0001';

  }

  const headers =
    values[0];

  const idColumn =
    headers.findIndex(
      function(header) {

        return normalizeHeader(
          header
        ) === 'company id';

      }
    );

  if (
    idColumn === -1
  ) {

    throw new Error(
      'Company ID column not found.'
    );

  }

  const used =
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

      used[
        Number(
          match[1]
        )
      ] = true;

    }

  }

  let number =
    1;

  while (
    used[number]
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
 * UPDATE INDUSTRIAL AREA
 *************************************************/

function updateAreaRecord(
  rowNumber,
  data
) {

  const ss =
    SpreadsheetApp
      .getActiveSpreadsheet();

  const sheet =
    ss.getSheetByName(
      AREAS_SHEET
    );

  if (!sheet) {

    throw new Error(
      'Areas sheet not found.'
    );

  }

  const row =
    Number(
      rowNumber
    );

  if (
    !row ||
    row < 2 ||
    row > sheet.getLastRow()
  ) {

    throw new Error(
      'Invalid Areas sheet row.'
    );

  }

  if (!data) {

    throw new Error(
      'Area data is required.'
    );

  }

  const values =
    sheet
      .getDataRange()
      .getValues();

  const headers =
    values[0].map(
      normalizeHeader
    );

  const areaColumn =
    headers.indexOf(
      'industrial area'
    );

  const cityColumn =
    headers.indexOf(
      'city'
    );

  const latitudeColumn =
    headers.indexOf(
      'latitude'
    );

  const longitudeColumn =
    headers.indexOf(
      'longitude'
    );

  if (
    areaColumn === -1 ||
    cityColumn === -1 ||
    latitudeColumn === -1 ||
    longitudeColumn === -1
  ) {

    throw new Error(
      'Areas sheet must contain Industrial Area, City, Latitude and Longitude columns.'
    );

  }

  const area =
    String(
      data.industrialArea || ''
    ).trim();

  const city =
    String(
      data.city || ''
    ).trim();

  const latitude =
    Number(
      data.latitude
    );

  const longitude =
    Number(
      data.longitude
    );

  if (!area) {

    throw new Error(
      'Industrial Area is required.'
    );

  }

  if (
    !Number.isFinite(latitude)
  ) {

    throw new Error(
      'Invalid Latitude.'
    );

  }

  if (
    !Number.isFinite(longitude)
  ) {

    throw new Error(
      'Invalid Longitude.'
    );

  }

  sheet
    .getRange(
      row,
      areaColumn + 1
    )
    .setValue(
      area
    );

  sheet
    .getRange(
      row,
      cityColumn + 1
    )
    .setValue(
      city
    );

  sheet
    .getRange(
      row,
      latitudeColumn + 1
    )
    .setValue(
      latitude
    );

  sheet
    .getRange(
      row,
      longitudeColumn + 1
    )
    .setValue(
      longitude
    );

  return true;

}


/*************************************************
 * UNIVERSAL SEARCH
 *
 * Searches:
 * - Industrial Areas
 * - Customers
 * - People
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


  /***********************************************
   * INDUSTRIAL AREAS
   ***********************************************/

  const areas =
    getIndustrialAreas();

  areas.forEach(
    function(area) {

      const text =
        [
          area.industrialArea,
          area.city
        ]
        .join(' ')
        .toLowerCase();

      if (
        text.indexOf(search) !== -1
      ) {

        results.push({

          type:
            'area',

          industrialArea:
            area.industrialArea,

          city:
            area.city,

          latitude:
            area.latitude,

          longitude:
            area.longitude

        });

      }

    }
  );


  /***********************************************
   * CUSTOMERS
   ***********************************************/

  const customers =
    getAllCustomers();

  customers.forEach(
    function(customer) {

      const text =
        [
          customer.companyId,
          customer.customerName,
          customer.industrialArea,
          customer.city,
          customer.type,
          customer.status,
          customer.potential,
          customer.website,
          customer.notes
        ]
        .join(' ')
        .toLowerCase();

      if (
        text.indexOf(search) !== -1
      ) {

        results.push({

          type:
            'customer',

          name:
            customer.customerName,

          companyId:
            customer.companyId,

          industrialArea:
            customer.industrialArea,

          city:
            customer.city

        });

      }

    }
  );


  /***********************************************
   * PEOPLE
   ***********************************************/

  const ss =
    SpreadsheetApp
      .getActiveSpreadsheet();

  const peopleSheet =
    ss.getSheetByName(
      PEOPLE_SHEET
    );

  if (
    peopleSheet
  ) {

    const values =
      peopleSheet
        .getDataRange()
        .getDisplayValues();

    if (
      values.length > 1
    ) {

      const headers =
        values[0].map(
          normalizeHeader
        );

      const peopleIdColumn =
        findHeader(
          headers,
          [
            'people id'
          ]
        );

      const companyIdColumn =
        findHeader(
          headers,
          [
            'company id'
          ]
        );

      const companyColumn =
        findHeader(
          headers,
          [
            'company'
          ]
        );

      const nameColumn =
        findHeader(
          headers,
          [
            'name'
          ]
        );

      const designationColumn =
        findHeader(
          headers,
          [
            'designation'
          ]
        );

      const phoneColumn =
        findHeader(
          headers,
          [
            'phone'
          ]
        );

      const emailColumn =
        findHeader(
          headers,
          [
            'email'
          ]
        );


      /*
       * Company ID → Industrial Area
       */

      const companyAreaMap =
        {};

      customers.forEach(
        function(customer) {

          companyAreaMap[
            String(
              customer.companyId
            )
            .trim()
            .toLowerCase()
          ] =
            customer.industrialArea;

        }
      );


      for (
        let i = 1;
        i < values.length;
        i++
      ) {

        const row =
          values[i];

        const peopleId =
          getCell(
            row,
            peopleIdColumn
          );

        const companyId =
          getCell(
            row,
            companyIdColumn
          );

        const company =
          getCell(
            row,
            companyColumn
          );

        const name =
          getCell(
            row,
            nameColumn
          );

        const designation =
          getCell(
            row,
            designationColumn
          );

        const phone =
          getCell(
            row,
            phoneColumn
          );

        const email =
          getCell(
            row,
            emailColumn
          );

        const area =
          companyAreaMap[
            companyId
              .trim()
              .toLowerCase()
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
            area
          ]
          .join(' ')
          .toLowerCase();

        if (
          searchable.indexOf(search) !== -1
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

            industrialArea:
              area

          });

        }

      }

    }

  }


  return results.slice(
    0,
    30
  );

}


/*************************************************
 * TOTAL CUSTOMER COUNT
 *************************************************/

function getTotalCustomerCount() {

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

  return Math.max(
    0,
    sheet.getLastRow() - 1
  );

}


/*************************************************
 * SPREADSHEET URL
 *************************************************/

function getSpreadsheetUrl() {

  return SpreadsheetApp
    .getActiveSpreadsheet()
    .getUrl();

}


/*************************************************
 * SETTINGS
 *
 * Settings sheet contains:
 *
 * Industrial Area
 * City
 * Type
 * Status
 * Designation
 *
 * Industrial Hub is no longer used.
 *************************************************/

function getSettingsOptions() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SETTINGS_SHEET);

  if (!sheet) {
    throw new Error('Settings sheet not found.');
  }

  const values = sheet.getDataRange().getValues();

  if (!values.length) {
    return {};
  }

  const result = {};

  values[0].forEach(function(header, column) {

    const key = String(header || '').trim();

    if (!key) return;

    // Ignore old Industrial Hub column
    if (key.toLowerCase() === 'industrial hub') {
      return;
    }

    const list = [];

    for (let row = 1; row < values.length; row++) {

      const value = String(values[row][column] || '').trim();

      if (value && list.indexOf(value) === -1) {
        list.push(value);
      }
    }

    // Alphabetical order
    list.sort(function(a, b) {
      return a.localeCompare(b, undefined, {
        sensitivity: 'base'
      });
    });

    result[key] = list;
  });

  return result;
}

/*************************************************
 * GET ONE SETTINGS LIST
 *************************************************/

/*************************************************
 * GET ONE SETTINGS LIST
 *************************************************/

function getSettingList(
  settingName
) {

  const ss =
    SpreadsheetApp
      .getActiveSpreadsheet();

  const sheet =
    ss.getSheetByName(
      SETTINGS_SHEET
    );

  if (!sheet) {

    throw new Error(
      'Settings sheet not found.'
    );

  }

  const values =
    sheet
      .getDataRange()
      .getValues();

  if (!values.length) {

    return [];

  }

  const headers =
    values[0].map(
      function(header) {

        return String(
          header || ''
        ).trim();

      }
    );

  const column =
    headers.findIndex(
      function(header) {

        return header
          .toLowerCase() ===
          String(
            settingName || ''
          )
          .trim()
          .toLowerCase();

      }
    );

  if (
    column === -1
  ) {

    throw new Error(
      'Setting column not found: ' +
      settingName
    );

  }

  const list =
    [];

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
      list.indexOf(value) === -1
    ) {

      list.push(
        value
      );

    }

  }

  // Alphabetical A-Z
  list.sort(
    function(a, b) {

      return String(a).localeCompare(
        String(b),
        undefined,
        {
          sensitivity: 'base'
        }
      );

    }
  );

  return list;

}

/*************************************************
 * ADD SETTING
 *************************************************/

function addSettingValue(
  settingName,
  value
) {

  const ss =
    SpreadsheetApp
      .getActiveSpreadsheet();

  const sheet =
    ss.getSheetByName(
      SETTINGS_SHEET
    );

  if (!sheet) {

    throw new Error(
      'Settings sheet not found.'
    );

  }

  const cleanName =
    String(
      settingName || ''
    ).trim();

  const cleanValue =
    String(
      value || ''
    ).trim();

  if (!cleanName) {

    throw new Error(
      'Setting name is required.'
    );

  }

  if (!cleanValue) {

    throw new Error(
      'Value is required.'
    );

  }

  const values =
    sheet
      .getDataRange()
      .getValues();

  const headers =
    values[0].map(
      function(header) {

        return String(
          header || ''
        ).trim();

      }
    );

  const column =
    headers.findIndex(
      function(header) {

        return header
          .toLowerCase() ===
          cleanName.toLowerCase();

      }
    );

  if (
    column === -1
  ) {

    throw new Error(
      'Setting column not found: ' +
      cleanName
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

      throw new Error(
        '"' +
        cleanValue +
        '" already exists.'
      );

    }

  }

  let targetRow =
    -1;

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

  if (
    targetRow === -1
  ) {

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
 * EDIT SETTING
 *************************************************/

function updateSettingValue(
  settingName,
  oldValue,
  newValue
) {

  const ss =
    SpreadsheetApp
      .getActiveSpreadsheet();

  const sheet =
    ss.getSheetByName(
      SETTINGS_SHEET
    );

  if (!sheet) {

    throw new Error(
      'Settings sheet not found.'
    );

  }

  const cleanName =
    String(
      settingName || ''
    ).trim();

  const oldText =
    String(
      oldValue || ''
    ).trim();

  const newText =
    String(
      newValue || ''
    ).trim();

  if (!newText) {

    throw new Error(
      'Value is required.'
    );

  }

  const values =
    sheet
      .getDataRange()
      .getValues();

  const headers =
    values[0].map(
      function(header) {

        return String(
          header || ''
        ).trim();

      }
    );

  const column =
    headers.findIndex(
      function(header) {

        return header
          .toLowerCase() ===
          cleanName.toLowerCase();

      }
    );

  if (
    column === -1
  ) {

    throw new Error(
      'Setting column not found: ' +
      cleanName
    );

  }

  let foundRow =
    -1;

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

  if (
    foundRow === -1
  ) {

    throw new Error(
      'Existing value not found: ' +
      oldText
    );

  }

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
        '"' +
        newText +
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
 * DELETE SETTING
 *************************************************/

function deleteSettingValue(
  settingName,
  value
) {

  const ss =
    SpreadsheetApp
      .getActiveSpreadsheet();

  const sheet =
    ss.getSheetByName(
      SETTINGS_SHEET
    );

  if (!sheet) {

    throw new Error(
      'Settings sheet not found.'
    );

  }

  const cleanName =
    String(
      settingName || ''
    ).trim();

  const cleanValue =
    String(
      value || ''
    ).trim();

  const values =
    sheet
      .getDataRange()
      .getValues();

  const headers =
    values[0].map(
      function(header) {

        return String(
          header || ''
        ).trim();

      }
    );

  const column =
    headers.findIndex(
      function(header) {

        return header
          .toLowerCase() ===
          cleanName.toLowerCase();

      }
    );

  if (
    column === -1
  ) {

    throw new Error(
      'Setting column not found: ' +
      cleanName
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
    'Value not found: ' +
    cleanValue
  );

}


/*************************************************
 * SAVE PERSON
 *************************************************/

function savePersonRecord(
  peopleId,
  companyId,
  data
) {

  const ss =
    SpreadsheetApp
      .getActiveSpreadsheet();

  const sheet =
    ss.getSheetByName(
      PEOPLE_SHEET
    );

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
    values[0].map(
      normalizeHeader
    );

  const idColumn =
    headers.indexOf(
      'people id'
    );

  if (
    idColumn === -1
  ) {

    throw new Error(
      'People ID column not found.'
    );

  }

  let rowNumber =
    -1;

  if (peopleId) {

    for (
      let i = 1;
      i < values.length;
      i++
    ) {

      if (
        String(
          values[i][idColumn] || ''
        ).trim() ===
        String(
          peopleId
        ).trim()
      ) {

        rowNumber =
          i + 1;

        break;

      }

    }

  }

  if (
    rowNumber === -1
  ) {

    rowNumber =
      sheet.getLastRow() + 1;

    peopleId =
      generateNextPeopleId(
        values,
        idColumn
      );

    sheet
      .getRange(
        rowNumber,
        idColumn + 1
      )
      .setValue(
        peopleId
      );

  }

  const companyName =
    getCompanyNameById(
      companyId
    );

  const photoColumn =
    headers.indexOf(
      'photo'
    );

  let oldPhoto =
    '';

  if (
    photoColumn !== -1 &&
    rowNumber <=
      sheet.getLastRow()
  ) {

    oldPhoto =
      String(
        sheet
          .getRange(
            rowNumber,
            photoColumn + 1
          )
          .getValue() ||
        ''
      ).trim();

  }

  let finalPhoto =
    oldPhoto;

  if (
    data &&
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
    .forEach(
      function(header) {

        const column =
          headers.indexOf(
            header
          );

        if (
          column !== -1
        ) {

          sheet
            .getRange(
              rowNumber,
              column + 1
            )
            .setValue(
              fields[header] || ''
            );

        }

      }
    );

  return true;

}


/*************************************************
 * COMPANY NAME
 *************************************************/

function getCompanyNameById(
  companyId
) {

  const customer =
    getCustomerById(
      companyId
    );

  if (!customer) {

    return '';

  }

  return String(
    customer['Customer Name'] ||
    customer['customer name'] ||
    ''
  ).trim();

}


/*************************************************
 * PEOPLE ID
 *************************************************/

function generateNextPeopleId(
  values,
  idColumn
) {

  let max =
    0;

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
          Number(
            match[1]
          )
        );

    }

  }

  return 'P' +
    String(
      max + 1
    ).padStart(
      4,
      '0'
    );

}


/*************************************************
 * GET PERSON BY ID
 *************************************************/

function getPersonById(
  peopleId
) {

  const ss =
    SpreadsheetApp
      .getActiveSpreadsheet();

  const sheet =
    ss.getSheetByName(
      PEOPLE_SHEET
    );

  if (!sheet) {

    throw new Error(
      'People sheet not found.'
    );

  }

  const values =
    sheet
      .getDataRange()
      .getValues();

  if (
    values.length < 2
  ) {

    return null;

  }

  const headers =
    values[0];

  const idColumn =
    headers.findIndex(
      function(header) {

        return normalizeHeader(
          header
        ) === 'people id';

      }
    );

  if (
    idColumn === -1
  ) {

    throw new Error(
      'People ID column not found.'
    );

  }

  for (
    let i = 1;
    i < values.length;
    i++
  ) {

    if (
      String(
        values[i][idColumn] || ''
      ).trim() ===
      String(
        peopleId
      ).trim()
    ) {

      const person =
        {};

      headers.forEach(
        function(header, index) {

          person[
            String(
              header
            ).trim()
          ] =
            values[i][index];

        }
      );

      return person;

    }

  }

  return null;

}


/*************************************************
 * GITHUB CONFIG
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
 * GITHUB TOKEN
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
        .getRange(
          'A3'
        )
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
 * GITHUB CONTENT URL
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
 * GET GITHUB FILE
 *************************************************/

function getGitHubFile(
  fileName,
  token
) {

  const response =
    UrlFetchApp.fetch(
      getGitHubContentsUrl(
        fileName
      ),
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
 * DELETE GITHUB FILE
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
      getGitHubContentsUrl(
        fileName
      ),
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

  if (
    response.getResponseCode() !==
    200
  ) {

    throw new Error(
      'Unable to delete old GitHub visiting card: ' +
      response.getContentText()
    );

  }

}


/*************************************************
 * UPLOAD VISITING CARD TO GITHUB
 *************************************************/

function uploadVisitingCardToGitHub(
  contactName,
  fileData,
  oldPhotoName
) {

  if (!contactName) {

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

  let extension =
    'png';

  const originalName =
    String(
      fileData.fileName || ''
    );

  const match =
    originalName.match(
      /\.([a-zA-Z0-9]+)$/
    );

  if (match) {

    extension =
      match[1].toLowerCase();

  }

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

  const newFileName =
    cleanName +
    '.' +
    extension;

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

  const existing =
    getGitHubFile(
      newFileName,
      token
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
      getGitHubContentsUrl(
        newFileName
      ),
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

  if (
    code !== 200 &&
    code !== 201
  ) {

    throw new Error(
      'GitHub upload failed (' +
      code +
      '): ' +
      response.getContentText()
    );

  }

  return newFileName;

}


/*************************************************
 * GITHUB IMAGE URL
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

  const folder =
    getVisitingCardFolder();

  if (oldPhotoName) {

    try {

      const oldFiles =
        folder.getFilesByName(
          oldPhotoName
        );

      while (
        oldFiles.hasNext()
      ) {

        oldFiles
          .next()
          .setTrashed(
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

  const safePeopleId =
    String(
      peopleId
    )
      .replace(
        /[^a-zA-Z0-9_-]/g,
        ''
      );

  const fileName =
    safePeopleId +
    '_VisitingCard.' +
    extension;

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

  const file =
    folder.createFile(
      blob
    );

  return file.getName();

}


/*************************************************
 * DRIVE VISITING CARD FOLDER
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


/*************************************************
 * HELPERS
 *************************************************/

function getCell(
  row,
  index
) {

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


function normalizeHeader(
  value
) {

  return String(
    value || ''
  )
    .trim()
    .toLowerCase()
    .replace(
      /\s+/g,
      ' '
    );

}


function findHeader(
  headers,
  names
) {

  for (
    let i = 0;
    i < headers.length;
    i++
  ) {

    if (
      names.indexOf(
        headers[i]
      ) !== -1
    ) {

      return i;

    }

  }

  return -1;

}


function setRowField(
  row,
  headers,
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

    row[column] =
      value || '';

  }

}

/*************************************************
 * GET CUSTOMER COUNTS BY INDUSTRIAL AREA
 *************************************************/

function getCustomerCountsByArea() {

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

  const values =
    sheet
      .getDataRange()
      .getDisplayValues();

  if (
    values.length < 2
  ) {

    return {};

  }

  const headers =
    values[0].map(
      normalizeHeader
    );

  const areaIndex =
    findHeader(
      headers,
      [
        'industrial area'
      ]
    );

  if (
    areaIndex === -1
  ) {

    throw new Error(
      'Customers sheet must contain Industrial Area column.'
    );

  }

  const counts = {};

  for (
    let i = 1;
    i < values.length;
    i++
  ) {

    const area =
      String(
        values[i][areaIndex] || ''
      )
      .trim();

    if (!area) {

      continue;

    }

    const key =
      area.toLowerCase();

    counts[key] =
      (counts[key] || 0) + 1;

  }

  return counts;

}