/**
 * @OnlyCurrentDoc
 */


function onHomepage(event) {
  return createHomeCard_();
}

function createHomeCard_(parameters = {
  items: items = '[]',
  token: null
}) {
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
      .setTitle('Select database'))
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
          .setParameters({ ...event.parameters })
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
      .pushCard(fieldCodeSetting_({ ...event.parameters, id: event.parameters.grid_item_identifier })))
    .build();
}

function openLinkCallback_(event) {
  return CardService.newActionResponseBuilder()
    .setOpenLink(CardService.newOpenLink()
      .setUrl(event.parameters.url))
    .build();
}

function fieldCodeSetting_(parameters) {
  const spreadsheet = SpreadsheetApp.openById(parameters.id);
  const sheets = spreadsheet.getSheets();
  const sheet = spreadsheet.getSheetByName(parameters.sheetName) || sheets[0];
  const range = sheet.getDataRange();
  const values = range.getDisplayValues();
  const headers = values[0];
  const cardSection = CardService.newCardSection()
    .setHeader('Field Code');
  const document = DocumentApp.getActiveDocument();
  const text = document.getBody().getText();
  const match = text.match(/{{[^{}]+?}}/g);
  const buttonSet = CardService.newButtonSet();
  const currentSheetName = sheet.getName();
  for (let i = 0; i < sheets.length; i++) {
    const sheetName = sheets[i].getName();
    const button = CardService.newTextButton()
      .setText(sheetName)
      .setOnClickAction(CardService.newAction()
        .setFunctionName(changeSheet_.name)
        .setParameters({
          ...parameters,
          sheetName: sheetName
        })
        .setLoadIndicator(CardService.LoadIndicator.SPINNER))
      .setTextButtonStyle(CardService.TextButtonStyle.TEXT);
    if (currentSheetName === sheetName) {
      button.setDisabled(true);
    }
    buttonSet.addButton(button);
  }
  for (const header of headers) {
    if (header) {
      const textInput = CardService.newTextInput()
        .setTitle(header)
        .setFieldName(header)
        .setValue(`{{${header}}}`)
        .setSuggestions(CardService.newSuggestions()
          .addSuggestions([...(match || []), `{{${header}}}`]))
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
        .setUrl(spreadsheet.getUrl()))
      .setText(`Open ${spreadsheet.getName()}`))
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
      data: JSON.stringify(values),
      template: text
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

function buildFixedFooter_(conditions = {
  primary: primary = false,
  secondary: secondary = false
}, content = {
  data: data = '',
  template: template = ''
}) {
  return CardService.newFixedFooter()
    .setPrimaryButton(CardService.newTextButton()
      .setText('Merge')
      .setDisabled(!conditions.primary)
      .setOnClickAction(CardService.newAction()
        .setFunctionName(clickMergeButton_.name)
        .setParameters(content)))
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
  const content = event.parameters;
  const template = content.template;
  const data = JSON.parse(content.data);
  const ui = DocumentApp.getUi();
  const htmlTemplate = HtmlService.createTemplateFromFile('ModalDialog');
  htmlTemplate.data = JSON.stringify(data);
  htmlTemplate.settings = JSON.stringify(settings);
  htmlTemplate.template = JSON.stringify(template);
  ui.showModalDialog(htmlTemplate.evaluate(), 'sashikomi');
  return CardService.newActionResponseBuilder()
    .setNotification(CardService.newNotification()
      .setText('Merging...'))
    .build();
}

function createMergeDocument() {
  const templateDocument = DocumentApp.getActiveDocument();
  const templateFile = DriveApp.getFileById(templateDocument.getId());
  const parents = templateFile.getParents();
  let destination = undefined;
  if (parents.hasNext()) {
    destination = parents.next();
  }
  const documentFile = templateFile.makeCopy('[Sashikomi]' + templateFile.getName(), destination);
  const document = DocumentApp.openById(documentFile.getId());
  const documentBody = document.getBody();
  documentBody.clear();
  document.saveAndClose();
  const url = documentFile.getUrl();
  return url;
}

function mergeDocument(url, entry) {
  const template = DocumentApp.getActiveDocument();
  const document = DocumentApp.openByUrl(url);
  const documentBody = document.getBody();
  const templateBody = template.getBody();
  // const templateHeader = template.getHeader();
  // const templateFooter = template.getFooter();
  // const templateFootnotes = template.getFootnotes();
  const body = templateBody.copy();
  for (const [fieldCode, text] of Object.entries(entry)) {
    body.replaceText(fieldCode, text);
    console.log(`Replace Text: ${fieldCode}, ${text}`);
  }
  for (let i = 0; i < body.getNumChildren(); i++) {
    const child = body.getChild(i);
    switch (child.getType()) {
      case DocumentApp.ElementType.LIST_ITEM:
        documentBody.appendListItem(child.asListItem().copy());
        console.log(`Copy ListItem`);
        break;
      case DocumentApp.ElementType.PARAGRAPH:
        documentBody.appendParagraph(child.asParagraph().copy());
        console.log(`Copy Paragraph`);
        break;
      case DocumentApp.ElementType.TABLE:
        documentBody.appendTable(child.asTable().copy());
        console.log(`Copy Table`);
        break;
      default:
        console.log(`Unknown Type: ${child.getType()}`);
        break;
    }
  }
  document.saveAndClose();
  return;
}

function merge_(template, entries) {
  const merged = [];
  for (const entry of entries) {
    merged.push(template.replace(/{{[^{}]+?}}/g, (match) => {
      return entry[match];
    }));
  }
  return merged;
}

function getRecentSpreadsheets_(token = null) {
  const options = {
    corpora: 'allDrives',
    includeItemsFromAllDrives: true,
    orderBy: 'viewedByMeTime desc',
    pageSize: 10,
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
