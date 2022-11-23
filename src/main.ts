const properties = PropertiesService.getScriptProperties()
const API_HOST = properties.getProperty('PORTFOLIO_API_HOST')
const TEMPLATE_DOC_ID = properties.getProperty('TEMPLATE_DOC_ID')

console.log(TEMPLATE_DOC_ID)

const fileCopy = (timestamp: Date) => {
  const formattedDate = Utilities.formatDate(timestamp, "JST", "yyyy-MM-dd")
  const fileName = "職務経歴書"
  const templateFile = DriveApp.getFileById(TEMPLATE_DOC_ID)
  const newFile = templateFile.makeCopy(`${fileName}-${formattedDate}`)

  return newFile.getId()
}

const findTableWithText = (body: GoogleAppsScript.Document.Body, text: string) => {
  for (let i=0; i<body.getNumChildren(); i++) {
    const child = body.getChild(i)

    if (child.getType() ==  DocumentApp.ElementType.TABLE) {
      const table = child.asTable()
      if (table.findText(text)) {
        return i
      }
    }
  }
}

const findListItemWithText = (body: GoogleAppsScript.Document.Body | GoogleAppsScript.Document.TableCell, text) => {
  let index = -1

  for (let i=0; i<body.getNumChildren(); i++) {
    const child = body.getChild(i)

    if (child.getType() ==  DocumentApp.ElementType.LIST_ITEM) {
      const listItem = child.asListItem()
      if (listItem.getText() == text) {
         index = i
      }
    }
  }
  return index
}

const replaceListItem = (body: GoogleAppsScript.Document.Body | GoogleAppsScript.Document.TableCell, placeholder: string, list: string[]) => {
  const index = findListItemWithText(body, placeholder)
  const listItem = body.getChild(index).asListItem()
  const glyphType = listItem.getGlyphType()
  listItem.setGlyphType(glyphType)
  listItem.setText(list[0])

  for (let i=1; i<list.length; i++) {
    const li = body.insertListItem(index + i, list[i])  
    li.setGlyphType(glyphType)
    li.setListId(listItem)
    li.setIndentFirstLine(listItem.getIndentFirstLine())
    li.setIndentStart(listItem.getIndentStart())
    li.setIndentEnd(listItem.getIndentEnd())
  }
}

const fetchProfile = () => {
  const response = UrlFetchApp.fetch(`${API_HOST}/profile`).getContentText("UTF-8") 
  return JSON.parse(response).data
}

const fetchHistories = () => {
  const response = UrlFetchApp.fetch(`${API_HOST}/histories`).getContentText("UTF-8") 
  return JSON.parse(response).data
}  

const getHistoryTemplate = (body: GoogleAppsScript.Document.Body) => {
  for (let i=0; i<body.getNumChildren(); i++) {
    const child = body.getChild(i)
    if (child.getType() === DocumentApp.ElementType.PARAGRAPH)  {
      if ((child as GoogleAppsScript.Document.Paragraph).getText() == "***history-template-start***") {
        const historyTemplate = []
        for (let j=i+1; j<body.getNumChildren(); j++) {
          const child2 = body.getChild(j)
          if (child2.getType() === DocumentApp.ElementType.PARAGRAPH)  {
            if ((child2 as GoogleAppsScript.Document.Paragraph).getText() == "***history-template-end***") {
              return {historyTemplate, start: i, end: j}
            }
          }
          historyTemplate.push(child2.copy())
        }
      }
    }
  }
}

function main() {
  const timestamp = new Date()


  const docID = fileCopy(timestamp)
  const doc = DocumentApp.openById(docID)
  const body = doc.getBody()

  const profile = fetchProfile()

  body.replaceText("{timestamp}", Utilities.formatDate(timestamp, "JST", "yyyy/MM/dd"))
  body.replaceText("{profile.job}", profile.job)
  body.replaceText("{profile.description}", profile.description)
  replaceListItem(body, "{profile.skillDescription}", profile.skillDescription)
  replaceListItem(body, "{profile.licenses}", profile.licenses)
  body.replaceText("{profile.pr}", profile.pr)

  const { historyTemplate, start, end } = getHistoryTemplate(body)
  for (let i=start; i <= end; i++) {
    body.removeChild(body.getChild(start))
  }

  fetchHistories().forEach(history => {
    [...historyTemplate].reverse().forEach(historyTemplateItem => {
      switch(historyTemplateItem.getType()) {
        case DocumentApp.ElementType.PARAGRAPH: {
          body.insertParagraph(end, historyTemplateItem.copy())
          break
        }
        case DocumentApp.ElementType.TABLE: {
          body.insertTable(end, historyTemplateItem.copy())
          break
        }
      }

    })
    body.replaceText("{history.organization}", history.organization)
    body.replaceText("{history.startMonth}", `${history.startMonth.year}/${history.startMonth.month.toString().padStart(2, '0')}`)
    if (history.endMonth === null) {
      body.replaceText("{history.endMonth}", "現在")
    } else {
      body.replaceText("{history.endMonth}", `${history.endMonth.year}/${history.endMonth.month.toString().padStart(2, '0')}`)
    }
    const tableIndex = findTableWithText(body, "{product.title}")
    const table = body.getChild(tableIndex).asTable()
    const templateRow = table.getRow(1).copy()
    table.removeRow(1)
    history.products.forEach(product => {
      const row = templateRow.copy()
      row.replaceText("{product.startMonth}", `${product.startMonth.year}/${product.startMonth.month.toString().padStart(2, '0')}`)
      if (product.endMonth === null) {
        row.replaceText("{product.endMonth}", "現在")
      } else {
        row.replaceText("{product.endMonth}", `${product.endMonth.year}/${product.endMonth.month.toString().padStart(2, '0')}`)
      }
      row.replaceText("{product.title}", product.title)
      replaceListItem(row.getCell(1), "{product.description}", product.description)
      replaceListItem(row.getCell(2), "{product.technologies}", product.technologies)
      table.appendTableRow(row)
    }) 
  })


  doc.saveAndClose()
}
