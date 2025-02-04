function onHomepage(event) {
  return createHomeCard();
}

function createHomeCard(parameters = {
  items: items = '[]',
  token: null
}) {
  const files = parameters['token'] ? DriveApp.continueFileIterator(parameters['token']) : DriveApp.getFilesByType(MimeType.GOOGLE_SHEETS);
  const items = JSON.parse(parameters['items']);
  const grid = CardService.newGrid()
    .setNumColumns(1)
    .setOnClickAction(CardService.newAction()
      .setFunctionName('clickSpreadsheetList')
      .setLoadIndicator(CardService.LoadIndicator.SPINNER))
    .setBorderStyle(CardService.newBorderStyle()
      .setType(CardService.BorderType.NO_BORDER));
  for (let i = 0; i < items.length; i++) {
    grid.addItem(CardService.newGridItem()
      .setIdentifier(items[i].id)
      .setTitle(items[i].title)
      .setSubtitle(items[i].subtitle));
  }
  let count = 0;
  while (files.hasNext()) {
    if (count < 5) {
      const file = files.next();
      const date = file.getLastUpdated();
      const user = file.getOwner();
      const item = {
        id: file.getId(),
        title: file.getName(),
        subtitle: `${date.toLocaleDateString()} ${date.toLocaleTimeString()} ${user && user.getName() || ''}`
      };
      items.push(item);
      grid.addItem(CardService.newGridItem()
        .setIdentifier(item.id)
        .setTitle(item.title)
        .setSubtitle(item.subtitle));
    } else {
      break;
    }
    count++;
  }
  return CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader()
      .setTitle('Select database'))
    .addSection(CardService.newCardSection()
      .addWidget(CardService.newButtonSet()
        .addButton(CardService.newTextButton()
          .setText('Search')
          .setOnClickAction(CardService.newAction()
            .setFunctionName('clickSearchButton')
            .setLoadIndicator(CardService.LoadIndicator.SPINNER))
          .setTextButtonStyle(CardService.TextButtonStyle.FILLED))))
    .addSection(CardService.newCardSection()
      .setHeader('Recent files')
      .addWidget(grid)
      .addWidget(CardService.newTextButton()
        .setText('More')
        .setOnClickAction(CardService.newAction()
          .setFunctionName('clickMoreButton')
          .setParameters({ items: JSON.stringify(items), token: files.getContinuationToken() })
          .setLoadIndicator(CardService.LoadIndicator.SPINNER))
        .setTextButtonStyle(CardService.TextButtonStyle.TEXT)))
    .setFixedFooter(buildFixedFooter({
      primary: false,
      secondary: true
    }))
    .build();
}

function clickSearchButton(event) {
  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation()
      .pushCard(createSearchCard(event)))
    .build();
}

function createSearchCard(event) {
  const value = event.formInput['search'] || '';
  const grid = CardService.newGrid()
    .setNumColumns(1)
    .setOnClickAction(CardService.newAction()
      .setFunctionName('clickSpreadsheetList')
      .setLoadIndicator(CardService.LoadIndicator.SPINNER))
    .setBorderStyle(CardService.newBorderStyle()
      .setType(CardService.BorderType.NO_BORDER));
  if (value) {
    const files = DriveApp.searchFiles(`mimeType contains 'spreadsheet' and fullText contains '${value}'`);
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
          .setFunctionName('changeSearchValue')
          .setParameters({ ...event.parameters })
          .setLoadIndicator(CardService.LoadIndicator.SPINNER))
        .setMultiline(false)))
    .addSection(CardService.newCardSection()
      .setHeader('Search Result')
      .addWidget(grid))
    .setFixedFooter(buildFixedFooter({
      primary: false,
      secondary: true
    }))
    .build();
}

function changeSearchValue(event) {
  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation()
      .updateCard(createSearchCard(event)))
    .build();
}

function clickMoreButton(event) {
  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation()
      .updateCard(createHomeCard(event.parameters)))
    .build();
}

function clickSpreadsheetList(event) {
  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation()
      .pushCard(fieldCodeSetting({ ...event.parameters, id: event.parameters.grid_item_identifier })))
    .build();
}

function openLinkCallback(event) {
  return CardService.newActionResponseBuilder()
    .setOpenLink(CardService.newOpenLink()
      .setUrl(event.parameters.url))
    .build();
}

function fieldCodeSetting(parameters) {
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
        .setFunctionName('changeSheet')
        .setParameters({ ...parameters, sheetName: sheetName })
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
          .setFunctionName('changeFieldCode')
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
    .setFixedFooter(buildFixedFooter({
      primary: true,
      secondary: true
    }, {
      data: JSON.stringify(values),
      template: text
    }))
    .build();
}

function changeSheet(event) {
  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation()
      .updateCard(fieldCodeSetting(event.parameters)))
    .build();
}

function changeFieldCode(arg) {
  console.log(arg);
}

function buildFixedFooter(conditions = {
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
        .setFunctionName('clickMergeButton')
        .setParameters(content)))
    .setSecondaryButton(CardService.newTextButton()
      .setText('Back')
      .setDisabled(!conditions.secondary)
      .setOnClickAction(CardService.newAction()
        .setFunctionName('gotoPreviousCard')));
}

function gotoPreviousCard(event) {
  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation()
      .popCard())
    .build();
}

function clickMergeButton(event) {
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
  const template = DocumentApp.getActiveDocument();
  const document = DocumentApp.create('[Sashikomi]' + template.getName());
  const url = document.getUrl();
  document.saveAndClose();
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
  }
  for (let i = 0; i < body.getNumChildren(); i++) {
    const child = body.getChild(i);
    switch (child.getType()) {
      case DocumentApp.ElementType.LIST_ITEM:
        documentBody.appendListItem(child.asListItem().copy());
        break;
      case DocumentApp.ElementType.PARAGRAPH:
        documentBody.appendParagraph(child.asParagraph().copy());
        break;
      case DocumentApp.ElementType.TABLE:
        documentBody.appendTable(child.asTable().copy());
        break;
      default:
        break;
    }
  }
  document.saveAndClose();
  return;
}

function merge(template, entries) {
  const merged = [];
  for (const entry of entries) {
    merged.push(template.replace(/{{[^{}]+?}}/g, (match) => {
      return entry[match];
    }));
  }
  return merged;
}
