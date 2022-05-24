/**
 * @OnlyCurrentDoc
 */

// 👆 narrows down the scopes required by this plugin, see https://developers.google.com/apps-script/add-ons/concepts/editor-scopes#editor_add-on_scopes 

function runImport(token) {
  const pbUrl = "https://api.productboard.com"
  // Fetch all the custom fields definitions. We will use it as columns.
  const customFields = fetchAll(token, `${pbUrl}/hierarchy-entities/custom-fields?type=text,number,dropdown,member`)
  // Fetch all the custom fields values. These will be the values in the rows.
  const customFieldsValues = fetchAll(token, `${pbUrl}/hierarchy-entities/custom-fields-values?type=text,number,dropdown,member`)

  // Group custom fields values by feature for faster look-up
  const customFieldsValuesByFeature = {}
  customFieldsValues.forEach(value => customFieldsValuesByFeature[getCustomFieldValueKey(value.hierarchyEntity, value.customField)] = value)

  // Set up the sheet
  const sheet = createSheet()
  insertHeader(sheet, customFields)

  // Let's process features page-by-page to optimize memory usage for very large datasets
  // - create an iterator over "pages" of features
  const featuresPageIterator = pagedResourceFetcher(token, `${pbUrl}/features`)
  // - process and append each page
  let currentRow = 2  // row 1 is the header
  for(features of featuresPageIterator) {
      // Iterate through all the features
    const rows = features.map(feature => {
      // First three columns will be feature's name, description and status.
      const featureProperties = [feature.name, stripHtml(feature.description), feature.status.name]
      // The rest of the columns will be custom fields values for this feature.
      // Iterate through all of the custom fields (columns) and see if they have value set for this feature.
      const customFieldsValues = customFields.map(field => findCustomFieldValue(field, feature))
      // Create a whole row
      return featureProperties.concat(customFieldsValues)
    });

    // Insert all rows with features at once
    sheet.getRange(currentRow, 1, rows.length, rows[0].length).setValues(rows)
    currentRow += rows.length
  }

  // Fix the height of all rows to prevent overly high rows on long content
  sheet.setRowHeightsForced(1, currentRow, 21)
  // Fix the width of the first column (feature's name) to fit the content
  sheet.autoResizeColumn(1)

  // This is a local function on purpose so that it has access to `customFieldsValuesByFeature`
  function findCustomFieldValue(field, feature) {
    const value = customFieldsValuesByFeature[getCustomFieldValueKey(feature, field)]
    if (value == null) return
    if (value.type == "dropdown") {
        // value of dropdown custom fields is an object so we need to handle it separately
        return value.value.label
    } if (value.type == "member") {
        return value.value.email
    } else {
        return value.value
    }
  }

  function getCustomFieldValueKey(feature, customField) {
    return `${feature.id}:${customField.id}`
  }
}

function createSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet()
  sheet.setName("Imported from Productboard")
  return sheet
}

function insertHeader(sheet, customFields) {
  const customFieldNames = customFields.map(cf => cf.name ?? `[Unnamed ${cf.type}]`)
  const header = ['Name', 'Description', 'Status', ...customFieldNames]
  sheet.appendRow(header)
  sheet.getRange(1, 1, 1, header.length).setFontWeight("bold")
}

// API-related code

/**
 * Convenience function for iterating over all "pages" of given resource starting with the `url`.
 */
function fetchAll(token, url) {
  const fetcher = pagedResourceFetcher(token, url)
  return Array.from(fetcher).flat()
}

/**
 * A generator function returning an interator over "pages" of data. 
 * The iterator yields a "page" on each iteration, following `links.next` until there is no next link returned.
 */
function* pagedResourceFetcher(token, url) {
  let nextLink = url
  do {
    let page = fetch(token, nextLink)
    nextLink = page.links.next
    yield page.data
  } while(nextLink)
}

/**
 * Issues a single GET request to given `url` using the `token` for authentication.
 * Returns an array of objects.
 */
function fetch(token, url) {
  const response = UrlFetchApp.fetch(url, {
    method: "GET",
    headers: {
      "Authorization": "Bearer " + token,
      "X-Version": "1"
    }
  })

  return JSON.parse(response.getContentText())
}

// Code for interfacing with the Google Sheets UI

const PB_TOKEN_KEY = "com.productboard.api.token"

function onInstall(e) {
  // Need to onOpen() to make sure menus are created as the document is already opened
  onOpen(e)
}

function onOpen(e) {
  SpreadsheetApp.getUi()
      .createMenu('Productboard')
      .addItem('Import features', 'showPrompt')
      .addItem('Clear token', 'clearToken')
      .addToUi();
}

function clearToken() {
  PropertiesService.getUserProperties().deleteProperty(PB_TOKEN_KEY)
}

function showPrompt() {
  const ui = SpreadsheetApp.getUi();

  // check properties for already stored token
  let token = PropertiesService.getUserProperties().getProperty(PB_TOKEN_KEY)

  if (!token) {
    // prompt the user for a token if no stored token found
    const result = ui.prompt(
        'Let\'s import some features!',
        'Please enter your API token:',
        ui.ButtonSet.OK_CANCEL);

    const button = result.getSelectedButton();
  
    if (button == ui.Button.OK) {
      token = result.getResponseText();
    } else {
      // do nothing if the user clicked anything but "OK"
      return
    }
  }

  try {
    runImport(token)
    // Store the token for later use by the same user
    PropertiesService.getUserProperties().setProperty(PB_TOKEN_KEY, token)
  } catch(e) {
    Logger.log(e)
    ui.alert('Import has failed', e, ui.ButtonSet.OK)
  }
}

// General utilities

function stripHtml(html) {
  // Don't try this at home, there are better ways to strip HTML in javascript, but we are limited by the Apps Script environment.
   return html.replace(/<[^>]*>?/gm, '');
}

