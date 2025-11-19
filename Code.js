/**
 * @OnlyCurrentDoc
 */

function onHomepage(event) {
  return createHomeCard_();
}

function createHomeCard_(parameters) {
  parameters = Object.assign({
    items: '[]',
    token: null
  }, parameters);

  const items = JSON.parse(parameters['items']);
  const grid = CardService.newGrid()
    .setNumColumns(1)
    .setOnClickAction(CardService.newAction()
      .setFunctionName(clickSpreadsheetList_.name)
      .setLoadIndicator(CardService.LoadIndicator.SPINNER))
    .setBorderStyle(CardService.newBorderStyle()
      .setType(CardService.BorderType.NO_BORDER));
  for (let i = 0; i < items.length; i++) {
    grid.addItem(CardService.newGridItem()
      .setIdentifier(items[i].id)
      .setTitle(items[i].title)
      .setSubtitle(items[i].subtitle));
  }
  const recent = getRecentSpreadsheets_(parameters['token']);
  for (const file of recent.files) {
    const f = DriveApp.getFileById(file.id);
    const date = f.getLastUpdated();
    const user = f.getOwner();
    const item = {
      id: file.id,
      title: file.name,
      subtitle: `${date.toLocaleDateString()} ${date.toLocaleTimeString()} ${user && user.getName() || ''}`
    };
    items.push(item);
    grid.addItem(CardService.newGridItem()
      .setIdentifier(item.id)
      .setTitle(item.title)
      .setSubtitle(item.subtitle));
  }
  return CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader()
      .setTitle('Select Sheet'))
    .addSection(CardService.newCardSection()
      .addWidget(CardService.newButtonSet()
        .addButton(CardService.newTextButton()
          .setText('Search')
          .setOnClickAction(CardService.newAction()
            .setFunctionName(clickSearchButton_.name)
            .setLoadIndicator(CardService.LoadIndicator.SPINNER))
          .setTextButtonStyle(CardService.TextButtonStyle.FILLED))))
    .addSection(CardService.newCardSection()
      .setHeader('Recent files')
      .addWidget(grid)
      .addWidget(CardService.newTextButton()
        .setText('More')
        .setOnClickAction(CardService.newAction()
          .setFunctionName(clickMoreRecentButton_.name)
          .setParameters({
            items: JSON.stringify(items),
            token: recent.nextPageToken
          })
          .setLoadIndicator(CardService.LoadIndicator.SPINNER))
        .setTextButtonStyle(CardService.TextButtonStyle.TEXT)))
    .setFixedFooter(buildFixedFooter_({
      primary: false,
      secondary: true
    }))
    .build();
}

function clickSearchButton_(event) {
  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation()
      .pushCard(createSearchCard_(event)))
    .build();
}

function createSearchCard_(event) {
  const value = event.formInput['search'] || '';
  const grid = CardService.newGrid()
    .setNumColumns(1)
    .setOnClickAction(CardService.newAction()
      .setFunctionName(clickSpreadsheetList_.name)
      .setLoadIndicator(CardService.LoadIndicator.SPINNER))
    .setBorderStyle(CardService.newBorderStyle()
      .setType(CardService.BorderType.NO_BORDER));
  if (value) {
    const files = DriveApp.searchFiles(`mimeType = 'application/vnd.google-apps.spreadsheet' and fullText contains '${value}'`);
    while (files.hasNext()) {
      const file = files.next();
      const name = file.getName();
      const date = file.getLastUpdated();
      const user = file.getOwner();
      grid.addItem(CardService.newGridItem()
        .setIdentifier(file.getId())
        .setTitle(name)
        .setSubtitle(`${date.toLocaleDateString()} ${date.toLocaleTimeString()} ${user && user.getName() || ''}`));
    }
  }
  return CardService.newCardBuilder()
    .addSection(CardService.newCardSection()
      .addWidget(CardService.newTextInput()
        .setTitle('Search')
        .setValue(value)
        .setFieldName('search')
        .setOnChangeAction(CardService.newAction()
          .setFunctionName(changeSearchValue_.name)
          .setParameters({
            ...event.parameters
          })
          .setLoadIndicator(CardService.LoadIndicator.SPINNER))
        .setMultiline(false)))
    .addSection(CardService.newCardSection()
      .setHeader('Search Result')
      .addWidget(grid))
    .setFixedFooter(buildFixedFooter_({
      primary: false,
      secondary: true
    }))
    .build();
}

function changeSearchValue_(event) {
  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation()
      .updateCard(createSearchCard_(event)))
    .build();
}

function clickMoreRecentButton_(event) {
  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation()
      .updateCard(createHomeCard_(event.parameters)))
    .build();
}

function clickMoreSearchButton_(event) {
  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation()
      .updateCard(createSearchCard_(event.parameters)))
    .build();
}

function clickSpreadsheetList_(event) {
  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation()
      .pushCard(fieldCodeSetting_({
        ...event.parameters,
        id: event.parameters.grid_item_identifier
      })))
    .build();
}

function openLinkCallback_(event) {
  return CardService.newActionResponseBuilder()
    .setOpenLink(CardService.newOpenLink()
      .setUrl(event.parameters.url))
    .build();
}

function fieldCodeSetting_(parameters) {
  const spreadsheet = getSpreadsheet_(parameters.id);
  spreadsheet.sheets.sort((a, b) => a.properties.index - b.properties.index);
  let currentSheet = spreadsheet.sheets[0];
  if (parameters.sheetId) {
    for (const sheet of spreadsheet.sheets) {
      if (parameters.sheetId === sheet.properties.sheetId.toString()) {
        currentSheet = sheet;
        break;
      }
    }
  }
  const sheetData = getSheetsDataByFilters_(spreadsheet.spreadsheetId, {
    gridRange: {
      sheetId: currentSheet.properties.sheetId,
      startRowIndex: 0,
      endRowIndex: 1,
    }
  });
  const buttonSet = CardService.newButtonSet();
  for (const sheet of spreadsheet.sheets) {
    const button = CardService.newTextButton()
      .setText(sheet.properties.title)
      .setOnClickAction(CardService.newAction()
        .setFunctionName(changeSheet_.name)
        .setParameters({
          ...parameters,
          sheetName: sheet.properties.title,
          sheetId: JSON.stringify(sheet.properties.sheetId),
        })
        .setLoadIndicator(CardService.LoadIndicator.SPINNER))
      .setTextButtonStyle(CardService.TextButtonStyle.TEXT);
    if (currentSheet.properties.sheetId === sheet.properties.sheetId) {
      button.setDisabled(true);
    }
    buttonSet.addButton(button);
  }

  const cardSection = CardService.newCardSection()
    .setHeader('Field Code');
  for (const headerValue of sheetData.sheets[0].data[0].rowData[0].values) {
    if (headerValue.formattedValue) {
      const textInput = CardService.newTextInput()
        .setTitle(headerValue.formattedValue)
        .setFieldName(headerValue.formattedValue)
        .setValue(`{{${headerValue.formattedValue}}}`)
        .setSuggestions(CardService.newSuggestions()
          .addSuggestions([
            `{{${headerValue.formattedValue}}}`,
          ]))
        .setOnChangeAction(CardService.newAction()
          .setFunctionName(changeFieldCode_.name)
          .setLoadIndicator(CardService.LoadIndicator.NONE))
        .setMultiline(false);
      cardSection.addWidget(textInput);
    }
  }
  return CardService.newCardBuilder()
    .addCardAction(CardService.newCardAction()
      .setOpenLink(CardService.newOpenLink()
        .setOpenAs(CardService.OpenAs.OVERLAY)
        .setOnClose(CardService.OnClose.NOTHING)
        .setUrl(spreadsheet.spreadsheetUrl))
      .setText(`Open ${spreadsheet.properties.title}`))
    .setHeader(CardService.newCardHeader()
      .setTitle('Match the field name'))
    .addSection(CardService.newCardSection()
      .setHeader('Sheets')
      .addWidget(buttonSet))
    .addSection(cardSection)
    .setFixedFooter(buildFixedFooter_({
      primary: true,
      secondary: true
    }, {
      spreadsheet: JSON.stringify(spreadsheet),
      sheet: JSON.stringify(currentSheet),
    }))
    .build();
}

function changeSheet_(event) {
  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation()
      .updateCard(fieldCodeSetting_(event.parameters)))
    .build();
}

function changeFieldCode_(arg) {
  console.log(arg);
}

function buildFixedFooter_(conditions, context) {
  conditions = Object.assign({
    primary: false,
    secondary: false
  }, conditions);

  return CardService.newFixedFooter()
    .setPrimaryButton(CardService.newTextButton()
      .setText('Merge')
      .setDisabled(!conditions.primary)
      .setOnClickAction(CardService.newAction()
        .setFunctionName(clickMergeButton_.name)
        .setParameters(context)))
    .setSecondaryButton(CardService.newTextButton()
      .setText('Back')
      .setDisabled(!conditions.secondary)
      .setOnClickAction(CardService.newAction()
        .setFunctionName(gotoPreviousCard_.name)));
}

function gotoPreviousCard_(event) {
  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation()
      .popCard())
    .build();
}

function clickMergeButton_(event) {
  const settings = event.formInput;

  const htmlTemplate = HtmlService.createTemplateFromFile('ModalDialog');
  htmlTemplate.spreadsheet = event.parameters.spreadsheet;
  htmlTemplate.sheet = event.parameters.sheet;
  htmlTemplate.settings = JSON.stringify(settings);

  const htmlOutput = htmlTemplate.evaluate();
  htmlOutput.setSandboxMode(HtmlService.SandboxMode.IFRAME)
  htmlOutput.setWidth(800)
  htmlOutput.setHeight(600)
  htmlOutput.addMetaTag("viewport", "width=device-width, initial-scale=1");

  const ui = DocumentApp.getUi();
  ui.showModalDialog(htmlOutput, 'Sashikomi');

  return CardService.newActionResponseBuilder()
    .setNotification(CardService.newNotification()
      .setText('Merging...'))
    .build();
}

function createTargetFolder(templateDocumentId) {
  const templateDocument = DriveApp.getFileById(templateDocumentId);
  const parentFolders = getParentsFolders_(templateDocument);
  const availableFolders = selectAvailableFolders_(parentFolders);
  const availableFolder = availableFolders[0];
  if (availableFolders.length > 1) {
    console.log('Multiple available folders found. Using the first one.');
  }
  const targetFolder = DriveApp.createFolder(`[Sashikomi]${templateDocument.getName()}`);
  targetFolder.moveTo(availableFolder);

  return {
    name: targetFolder.getName(),
    url: targetFolder.getUrl(),
    id: targetFolder.getId(),
  };
}

function createMergeDocument(templateDocumentId, targetFolderId, name) {
  const templateFile = DriveApp.getFileById(templateDocumentId);
  const targetFolder = DriveApp.getFolderById(targetFolderId);
  const mergeFile = templateFile.makeCopy(name, targetFolder);
  const url = mergeFile.getUrl();
  const mergeDocument = Docs.Documents.get(mergeFile.getId(), {
    includeTabsContent: false,
  });
  return {
    url: url,
    document: mergeDocument,
  };
}

function replaceMergeDocument(documentId, revisionId, replaceAllTextRequests) {
  const response = Docs.Documents.batchUpdate({
    requests: replaceAllTextRequests,
    writeControl: {
      requiredRevisionId: revisionId,
    }
  }, documentId);
  return response;
}

function getTemplateDocument() {
  const activeDocument = DocumentApp.getActiveDocument();
  const templateDocument = Docs.Documents.get(activeDocument.getId(), {
    includeTabsContent: false,
  });
  return templateDocument;
}

function getSheetsData(spreadsheetId, sheetName) {
  const fields = 'sheets(data(rowData(values(formattedValue))))';

  return Sheets.Spreadsheets.get(spreadsheetId, {
    fields: fields,
    ranges: [sheetName],
    includeGridData: true,
    excludeTablesInBandedRanges: false,
  });
}

function getParentsFolders_(file) {
  const parentFolders = [];
  const folders = file.getParents();
  while (folders.hasNext()) {
    const folder = folders.next();
    parentFolders.push(folder);
  }
  if (parentFolders.length === 0) {
    parentFolders.push(DriveApp.getRootFolder());
  }

  return parentFolders;
}

function selectAvailableFolders_(folders) {
  const user = Session.getActiveUser();
  const availableFolders = [];
  for (const folder of folders) {
    const permission = folder.getAccess(user);
    if (permission === DriveApp.Permission.EDIT ||
      permission === DriveApp.Permission.OWNER ||
      permission === DriveApp.Permission.ORGANIZER ||
      permission === DriveApp.Permission.FILE_ORGANIZER) {
      availableFolders.push(folder);
    } else {
      console.log(`No permission to write: ${folder.getName()}`);
    }
  }
  if (availableFolders.length === 0) {
    availableFolders.push(DriveApp.getRootFolder());
  }

  return availableFolders;
}

function getSpreadsheet_(spreadsheetId) {
  const fields = 'spreadsheetId,spreadsheetUrl,properties,sheets(properties(sheetId,title,index,sheetType,gridProperties))';

  return Sheets.Spreadsheets.get(spreadsheetId, {
    fields: fields,
    includeGridData: false,
    excludeTablesInBandedRanges: false,
  });
}

function getSheetsDataByFilters_(spreadsheetId, ...dataFilters) {
  return Sheets.Spreadsheets.getByDataFilter({
    dataFilters: [
      ...dataFilters,
    ],
    includeGridData: true,
    excludeTablesInBandedRanges: false,
  }, spreadsheetId);
}

function getRecentSpreadsheets_(token = null) {
  const options = {
    corpora: 'allDrives',
    includeItemsFromAllDrives: true,
    orderBy: 'viewedByMeTime desc',
    pageSize: 5,
    q: `trashed = false and mimeType = 'application/vnd.google-apps.spreadsheet'`,
    supportsAllDrives: true,
  };

  if (token) {
    options.pageToken = token;
  }

  const page = Drive.Files.list(options);

  return page;
}

function searchSpreadsheets_(value = '', token = null) {
  const options = {
    corpora: 'allDrives',
    includeItemsFromAllDrives: true,
    pageSize: 10,
    q: `trashed = false and mimeType = 'application/vnd.google-apps.spreadsheet' and fullText contains '${value}'`,
    supportsAllDrives: true,
  };

  if (token) {
    options.pageToken = token;
  }

  const page = Drive.Files.list(options);

  return page;
}
