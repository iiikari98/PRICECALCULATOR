import { useMemo, useState } from 'react'
import ExcelJS from 'exceljs'
import { saveAs } from 'file-saver'
import { jsPDF } from 'jspdf'
import autoTable from 'jspdf-autotable'
import { Calculator, Download, FileText, Plus, RefreshCw, Ship, Trash2 } from 'lucide-react'
import './App.css'

const termRules = {
  EXW: {
    title: 'EXW 工厂交货',
    note: '卖方报价到工厂/仓库交货点，通常不包含本地出口费用、海运费和保险。',
    needsLocal: false,
    needsOcean: false,
    needsInsurance: false,
  },
  FOB: {
    title: 'FOB 装运港船上交货',
    note: '卖方通常承担货物装上船前的本地费用和出口相关费用，不包含海运费和保险。',
    needsLocal: true,
    needsOcean: false,
    needsInsurance: false,
  },
  CFR: {
    title: 'CFR 成本加运费',
    note: '卖方报价包含本地费用和到目的港的海运费，但不负责购买运输保险。',
    needsLocal: true,
    needsOcean: true,
    needsInsurance: false,
  },
  CIF: {
    title: 'CIF 成本、保险加运费',
    note: '卖方报价包含本地费用、到目的港海运费和运输保险；低货值赠送保险时可勾选不计费。',
    needsLocal: true,
    needsOcean: true,
    needsInsurance: true,
  },
}

const companyProfiles = {
  shenzhen: {
    id: 'shenzhen',
    label: 'Shenzhen Jindaquan Technology Co., Ltd',
    name: 'Shenzhen Jindaquan Technology Co.,Ltd',
    address:
      'Room 1605, ShiHong Building, No. 2095 Bixin Road, Nanlian Community, Longgang Subdistrict, Longgang District, Shenzhen, Guangdong, China',
    tel: '+86-0755-28996208',
    fax: '+86-0755-28994568',
    bank: {
      accountNo: '000409407436',
      bankName: 'Shenzhen Rural Commercial Bank',
      swift: 'SRCCCNBSXXX',
      bankAddress: 'RCB BIdg, No.2028, Haixiu Road, Baoan Dist, Shenzhen China',
    },
  },
  guangdong: {
    id: 'guangdong',
    label: 'Guangdong Jindaquan Technology Co., Ltd',
    name: 'Guangdong Jindaquan Technology Co.,Ltd',
    address:
      'Bld 7#, Wanyang Zhongchuang Cheng, Shantangwei, Baishabu, Daling Street, Huizhou City, Guangdong Province, China',
    tel: '+86-0755-28996208',
    fax: '+86-0755-28994568',
    bank: {
      accountNo: '',
      bankName: 'BANK OF CHINA HUIDONG SUB-BRANCH',
      swift: 'BKCHCNBJ47A',
      bankAddress: 'NO.98, JIAN SHE ROAD, HUIDONG, GUANGDONG, CHINA',
    },
  },
}
const serialStorageKey = 'jdq-generated-document-nos'

const blankItem = () => ({
  id: crypto.randomUUID(),
  description: '',
  hsCode: '',
  qty: '',
  unitPrice: '',
  currency: 'USD',
})

const defaultCustomer = {
  company: '',
  attn: '',
  address: '',
  tel: '',
  buyer: '',
}

const defaultDoc = {
  no: '',
  date: new Date().toISOString().slice(0, 10),
  companyProfileId: 'shenzhen',
  by: '',
  customerType: 'company',
  documentSeq: '',
  countryCode: '',
  customerCode: '',
  customerOrderSeq: '',
  from: '',
  to: '',
  payment: '',
  leadTime: '',
}

const defaultFees = {
  incoterm: 'EXW',
  exchangeRate: '6.5',
  local: '',
  localCurrency: 'CNY',
  ocean: '',
  oceanCurrency: 'USD',
  bank: '',
  bankCurrency: 'USD',
  other: '',
  otherCurrency: 'USD',
  insuranceRate: '0.003',
  insuranceWaived: false,
}

const roundCurrency = (value) => Math.round((Number(value || 0) + Number.EPSILON) * 100) / 100

const money = (value) =>
  `US$${roundCurrency(value).toLocaleString('en-US', {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  })}`

const documentMoney = (value) => {
  const rounded = roundCurrency(value)
  const text = rounded.toLocaleString('en-US', {
    minimumFractionDigits: Number.isInteger(rounded) ? 0 : 2,
    maximumFractionDigits: 2,
  })
  return `${text}$`
}

const cny = (value) =>
  `RMB ${roundCurrency(value).toLocaleString('en-US', {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  })}`

const num = (value) => Number(value) || 0

function toUsd(value, currency, exchangeRate) {
  if (currency === 'CNY') return num(value) / Math.max(num(exchangeRate), 0.0001)
  return num(value)
}

function roundedRmbToUsd(value, exchangeRate) {
  return Math.round(num(value) / Math.max(num(exchangeRate), 0.0001))
}

function roundTwo(value) {
  return Math.round((num(value) + Number.EPSILON) * 100) / 100
}

function fromUsd(value, currency, exchangeRate) {
  if (currency === 'CNY') return num(value) * num(exchangeRate)
  return num(value)
}

const ones = [
  '',
  'ONE',
  'TWO',
  'THREE',
  'FOUR',
  'FIVE',
  'SIX',
  'SEVEN',
  'EIGHT',
  'NINE',
  'TEN',
  'ELEVEN',
  'TWELVE',
  'THIRTEEN',
  'FOURTEEN',
  'FIFTEEN',
  'SIXTEEN',
  'SEVENTEEN',
  'EIGHTEEN',
  'NINETEEN',
]

const tens = ['', '', 'TWENTY', 'THIRTY', 'FORTY', 'FIFTY', 'SIXTY', 'SEVENTY', 'EIGHTY', 'NINETY']

function underThousand(value) {
  const n = Math.floor(value)
  const hundred = Math.floor(n / 100)
  const rest = n % 100
  const parts = []
  if (hundred) parts.push(`${ones[hundred]} HUNDRED`)
  if (rest < 20) {
    if (rest) parts.push(ones[rest])
  } else {
    const ten = Math.floor(rest / 10)
    const one = rest % 10
    parts.push(one ? `${tens[ten]} ${ones[one]}` : tens[ten])
  }
  return parts.join(' ')
}

function amountWords(value) {
  const rounded = Math.round(num(value))
  if (!rounded) return 'SAY TOTAL U.S. DOLLARS ZERO ONLY'
  const millions = Math.floor(rounded / 1000000)
  const thousands = Math.floor((rounded % 1000000) / 1000)
  const rest = rounded % 1000
  const parts = []
  if (millions) parts.push(`${underThousand(millions)} MILLION`)
  if (thousands) parts.push(`${underThousand(thousands)} THOUSAND`)
  if (rest) parts.push(underThousand(rest))
  return `SAY TOTAL U.S. DOLLARS ${parts.join(' ')} ONLY`
}

function formatDate(date) {
  if (!date) return ''
  const [year, month, day] = date.split('-')
  return `${day}/${month}/${year}`
}

function formatCodeDate(date) {
  if (!date) return 'ddmmyy'
  const [year, month, day] = date.split('-')
  return `${day}${month}${year.slice(-2)}`
}

function twoDigits(value) {
  const n = Math.max(0, Math.floor(num(value)))
  return String(n).padStart(2, '0').slice(-2)
}

function suggestCustomerCode(customer, customerType) {
  const source = `${customer.company || customer.attn || customer.buyer || ''}`.trim().toUpperCase()
  const cleaned = source
    .replace(/\b(CO|COMPANY|LTD|LIMITED|INC|LLC|GMBH|SAS|SPA|TRADING|IMPORT|EXPORT)\b/g, ' ')
    .replace(/[^A-Z0-9 ]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
  const chars = cleaned.replace(/[^A-Z0-9]/g, '')
  if (!chars) return ''
  if (customerType === 'personal') return chars[0]
  const words = cleaned.split(' ').filter(Boolean)
  if (words.length >= 2) {
    return words.map((word) => word[0]).join('').slice(0, 4)
  }
  if (chars.length <= 3) return chars
  return `${chars[0]}${chars[1]}${chars[chars.length - 1]}`.slice(0, 3)
}

function buildDocumentNo(doc, customer) {
  const sequence = doc.documentSeq || doc.customerOrderSeq || 1
  const countryCode = (doc.countryCode || '').toUpperCase()
  const customerCode = (doc.customerCode || suggestCustomerCode(customer, doc.customerType)).toUpperCase()
  const customerOrderSeq = twoDigits(doc.customerOrderSeq || 1)
  return [
    'ALL',
    formatCodeDate(doc.date),
    twoDigits(sequence),
    '-',
    countryCode,
    customerCode,
    customerOrderSeq,
  ].join('')
}

function formatItemDescription(item) {
  const description = String(item.description || '').trim()
  const hsCode = String(item.hsCode || '').trim()
  if (!hsCode) return description
  return description ? `${description}\nHS Code: ${hsCode}` : `HS Code: ${hsCode}`
}

function readGeneratedNos() {
  try {
    return JSON.parse(localStorage.getItem(serialStorageKey) || '[]')
  } catch {
    return []
  }
}

function saveGeneratedNos(values) {
  localStorage.setItem(serialStorageKey, JSON.stringify(values.slice(-300)))
}

async function loadWorkbook(template) {
  const workbook = new ExcelJS.Workbook()
  const response = await fetch(`/templates/${template}`)
  const buffer = await response.arrayBuffer()
  await workbook.xlsx.load(buffer)
  return workbook
}

async function loadAssetDataUrl(path) {
  const response = await fetch(path)
  const blob = await response.blob()
  return new Promise((resolve) => {
    const reader = new FileReader()
    reader.onload = () => resolve(reader.result)
    reader.readAsDataURL(blob)
  })
}

async function loadAssetBuffer(path) {
  const response = await fetch(path)
  return response.arrayBuffer()
}

function setCell(sheet, address, value) {
  const cell = sheet.getCell(address)
  cell.value = value
  return cell
}

function setFittedCell(sheet, address, value, { baseSize = 12, minSize = 8, wrapAt = 28 } = {}) {
  const cell = setCell(sheet, address, value)
  const text = String(value || '')
  const shrinkSteps = Math.max(0, Math.ceil((text.length - wrapAt) / 12))
  cell.font = { ...(cell.font || {}), size: Math.max(minSize, baseSize - shrinkSteps) }
  cell.alignment = {
    ...(cell.alignment || {}),
    wrapText: text.length > wrapAt,
    vertical: cell.alignment?.vertical || 'center',
  }
  return cell
}

function adjustRowHeight(sheet, rowNumber, values, { baseHeight = 19, charsPerLine = 52, lineHeight = 12, maxHeight = 54 } = {}) {
  const longestLines = values
    .map((value) => String(value || '').split('\n').reduce((max, line) => Math.max(max, Math.ceil(line.length / charsPerLine)), 1))
    .reduce((max, lines) => Math.max(max, lines), 1)
  sheet.getRow(rowNumber).height = Math.min(maxHeight, Math.max(baseHeight, longestLines * lineHeight + 7))
}

function setMergedLine(sheet, rowNumber, value, { fontSize = 9, bold = false, wrap = false, height = 18, charsPerLine = 105 } = {}) {
  const range = `B${rowNumber}:I${rowNumber}`
  safeUnmerge(sheet, range)
  safeMerge(sheet, range)
  const cell = setCell(sheet, `B${rowNumber}`, value)
  cell.font = { ...(cell.font || {}), size: fontSize, bold }
  cell.alignment = {
    ...(cell.alignment || {}),
    horizontal: 'left',
    vertical: wrap ? 'top' : 'center',
    wrapText: wrap,
  }
  if (wrap) {
    adjustRowHeight(sheet, rowNumber, [value], { baseHeight: height, charsPerLine, lineHeight: 15, maxHeight: 78 })
  } else {
    sheet.getRow(rowNumber).height = height
  }
  return cell
}

function setMergedRangeLine(
  sheet,
  rowNumber,
  startCol,
  endCol,
  value,
  { fontSize = 10, wrap = true, height = 18, charsPerLine = 70, maxHeight = 88 } = {},
) {
  const range = `${startCol}${rowNumber}:${endCol}${rowNumber}`
  safeUnmerge(sheet, range)
  safeMerge(sheet, range)
  const cell = setCell(sheet, `${startCol}${rowNumber}`, value)
  cell.font = { ...(cell.font || {}), size: fontSize }
  cell.alignment = {
    ...(cell.alignment || {}),
    horizontal: 'left',
    vertical: wrap ? 'top' : 'center',
    wrapText: wrap,
  }
  if (wrap) {
    adjustRowHeight(sheet, rowNumber, [value], { baseHeight: height, charsPerLine, lineHeight: 13, maxHeight })
  } else {
    sheet.getRow(rowNumber).height = height
  }
  return cell
}

function copyStyle(from, to) {
  to.style = JSON.parse(JSON.stringify(from.style || {}))
  to.numFmt = from.numFmt
  to.alignment = from.alignment ? { ...from.alignment } : undefined
  to.border = from.border ? { ...from.border } : undefined
  to.fill = from.fill ? { ...from.fill } : undefined
  to.font = from.font ? { ...from.font } : undefined
}

function clearInvoiceRows(sheet, start, end) {
  for (let row = start; row <= end; row += 1) {
    ;['B', 'C', 'D', 'E', 'G', 'I'].forEach((col) => {
      sheet.getCell(`${col}${row}`).value = ''
    })
  }
}

function safeUnmerge(sheet, range) {
  try {
    sheet.unMergeCells(range)
  } catch {
    // Some templates do not merge this range.
  }
}

function safeMerge(sheet, range) {
  try {
    sheet.mergeCells(range)
  } catch {
    // The template may already have this range merged.
  }
}

function addSheetLogo(workbook, sheet) {
  return loadAssetBuffer('/assets/logo-mark.png').then((buffer) => {
    const logoId = workbook.addImage({ buffer, extension: 'png' })
    sheet.addImage(logoId, 'B1:C3')
  })
}

function clearTemplateSampleData(sheet) {
  ;[
    'B4',
    'D4',
    'D5',
    'D6',
    'D7',
    'D8',
    'G5',
    'H5',
    'H6',
    'H7',
    'G8',
    'I8',
    'C19',
    'C20',
    'C21',
    'G19',
    'G20',
    'G21',
    'G28',
  ].forEach((cell) => {
    setCell(sheet, cell, '')
  })
  clearInvoiceRows(sheet, 10, 18)
  for (let row = 19; row <= 30; row += 1) {
    ;['B', 'C', 'G'].forEach((col) => {
      sheet.getCell(`${col}${row}`).value = ''
    })
  }
}

function fillInvoice(sheet, payload, title) {
  const { customer, doc, rows, totals, fees, companyProfile } = payload
  const isPi = title === 'PROFORMA INVOICE'
  clearTemplateSampleData(sheet)
  const noLabel = title === 'COMMERCIAL INVOICE' ? 'CI NO. :' : title === 'PROFORMA INVOICE' ? 'PI NO. :' : 'QUOTE NO. :'
  setFittedCell(sheet, 'A1', companyProfile.name, { baseSize: 16, minSize: 11, wrapAt: 42 })
  setFittedCell(sheet, 'A2', companyProfile.address, { baseSize: 10, minSize: 7, wrapAt: 95 })
  sheet.getRow(2).height = 34
  safeUnmerge(sheet, 'A4:I4')
  safeMerge(sheet, 'A4:I4')
  const titleCell = setCell(sheet, 'A4', title)
  titleCell.font = { ...(titleCell.font || {}), name: 'Calibri', size: 14, bold: true }
  titleCell.alignment = { horizontal: 'center', vertical: 'middle' }
  setCell(sheet, 'G5', noLabel)
  setMergedRangeLine(sheet, 5, 'D', 'F', customer.company, { fontSize: 11, height: 24, charsPerLine: 36, maxHeight: 54 })
  setMergedRangeLine(sheet, 6, 'D', 'F', customer.attn, { fontSize: 11, height: 20, charsPerLine: 36, maxHeight: 44 })
  setMergedRangeLine(sheet, 7, 'D', 'F', customer.address, { fontSize: 10, height: 26, charsPerLine: 44, maxHeight: 72 })
  setFittedCell(sheet, 'D8', customer.tel, { baseSize: 10.5, minSize: 8, wrapAt: 34 })
  setCell(sheet, 'H5', doc.no)
  setCell(sheet, 'H6', formatDate(doc.date))
  setCell(sheet, 'H7', doc.by)
  setCell(sheet, 'G8', doc.from)
  setCell(sheet, 'I8', doc.to)
  adjustRowHeight(sheet, 5, [customer.company], { baseHeight: 24, charsPerLine: 38, maxHeight: 42 })
  adjustRowHeight(sheet, 7, [customer.address], { baseHeight: 33, charsPerLine: 48, maxHeight: 58 })

  const itemStartRow = 10
  const baseItemRows = 4
  const extraItemRows = Math.max(0, rows.length - baseItemRows)
  const baseTotalRow = 14
  if (extraItemRows > 0) {
    sheet.spliceRows(baseTotalRow, 0, ...Array.from({ length: extraItemRows }, () => []))
    for (let row = baseTotalRow; row < baseTotalRow + extraItemRows; row += 1) {
      for (let col = 1; col <= 13; col += 1) {
        copyStyle(sheet.getRow(13).getCell(col), sheet.getRow(row).getCell(col))
      }
    }
  }

  const totalRow = baseTotalRow + extraItemRows
  clearInvoiceRows(sheet, itemStartRow, totalRow)

  rows.forEach((row, index) => {
    const excelRow = itemStartRow + index
    setCell(sheet, `B${excelRow}`, index + 1)
    setMergedRangeLine(sheet, excelRow, 'C', 'D', row.description, { fontSize: 9, height: 22, charsPerLine: 48, maxHeight: 96 })
    setCell(sheet, `E${excelRow}`, row.qty === null ? '***' : `${row.qty}KG`)
    setCell(sheet, `G${excelRow}`, row.unitPrice === null ? '***' : documentMoney(row.unitPrice))
    setCell(sheet, `I${excelRow}`, documentMoney(row.subtotal))
  })

  setCell(sheet, `B${totalRow}`, '')
  setCell(sheet, `C${totalRow}`, '')
  setCell(sheet, `E${totalRow}`, '')
  setCell(sheet, `G${totalRow}`, '')
  setCell(sheet, `I${totalRow}`, documentMoney(totals.grand))

  setMergedLine(sheet, totalRow + 1, amountWords(totals.grand), { fontSize: 9, wrap: true, height: 20, charsPerLine: 100 })
  setCell(sheet, `B${totalRow + 2}`, 'Payment: ')
  setMergedRangeLine(sheet, totalRow + 2, 'D', 'I', doc.payment, { fontSize: 9, height: 20, charsPerLine: 74, maxHeight: 72 })
  setCell(sheet, `B${totalRow + 3}`, 'Lead-time: ')
  setMergedRangeLine(sheet, totalRow + 3, 'D', 'I', doc.leadTime, { fontSize: 9, height: 20, charsPerLine: 74, maxHeight: 72 })

  const rule = termRules[fees.incoterm]
  const insuranceNote = rule.needsInsurance
    ? fees.insuranceWaived
      ? 'Insurance: waived / included by shipping company'
      : `Insurance: ${(num(fees.insuranceRate) * 100).toFixed(3)}%`
    : ''
  const hasBankInfo = Boolean(companyProfile.bank.accountNo || companyProfile.bank.bankName || companyProfile.bank.swift)
  const bankRow = totalRow + 5
  if (isPi && hasBankInfo) {
    setMergedLine(sheet, bankRow, 'BANK ACCOUNT:', { fontSize: 10, bold: true, height: 20 })
    setMergedLine(sheet, bankRow + 1, `Beneficiary Name :   ${companyProfile.name.toUpperCase()}`, { height: 20 })
    setMergedLine(sheet, bankRow + 2, `Address of Beneficiary :   ${companyProfile.address}`, { wrap: true, height: 20 })
    setMergedLine(sheet, bankRow + 3, `Beneficiary Account No. :  ${companyProfile.bank.accountNo || ''}`, { height: 20 })
    setMergedLine(sheet, bankRow + 4, `Beneficiary bank:  ${companyProfile.bank.bankName || ''}`, { height: 20 })
    setMergedLine(sheet, bankRow + 5, `SWIFT Code:  ${companyProfile.bank.swift || ''}`, { height: 20 })
    setMergedLine(sheet, bankRow + 6, `Bank Address:  ${companyProfile.bank.bankAddress || ''}`, { wrap: true, height: 20 })
  }

  let buyerRow = isPi && hasBankInfo ? bankRow + 9 : totalRow + 5
  if (!isPi) {
    const noteRows = [
      insuranceNote,
      `Exchange rate: 1 USD = RMB ${num(fees.exchangeRate).toFixed(4)}`,
      `Final unit price: ${money(totals.cifPerKg)} / KG`,
    ].filter(Boolean)
    noteRows.forEach((note, index) => setMergedLine(sheet, totalRow + 4 + index, note, { fontSize: 9, wrap: true, height: 18 }))
    buyerRow = totalRow + 5 + noteRows.length
  }
  setCell(sheet, `C${buyerRow - 1}`, 'The seller')
  setCell(sheet, `G${buyerRow - 1}`, 'The buyer')
  setMergedRangeLine(sheet, buyerRow, 'C', 'F', companyProfile.name, { fontSize: 10, height: 24, charsPerLine: 44, maxHeight: 66 })
  setMergedRangeLine(sheet, buyerRow, 'G', 'I', customer.buyer || customer.company, { fontSize: 10, height: 24, charsPerLine: 34, maxHeight: 66 })
}

function buildRows(items, fees, totals) {
  const rule = termRules[fees.incoterm]
  const goods = items
    .filter((item) => item.description || item.hsCode || num(item.qty) || num(item.unitPrice))
    .map((item) => {
      const unitPriceUsd = toUsd(item.unitPrice, item.currency, fees.exchangeRate)
      return {
        description: formatItemDescription(item),
        qty: num(item.qty),
        unitPrice: unitPriceUsd,
        subtotal: num(item.qty) * unitPriceUsd,
      }
    })
  const extraRows = []
  if (rule.needsLocal && totals.local > 0) {
    extraRows.push({ description: 'Local charge', qty: null, unitPrice: null, subtotal: totals.local })
  }
  if (rule.needsOcean && totals.ocean > 0) {
    extraRows.push({ description: 'Freight', qty: null, unitPrice: null, subtotal: totals.ocean })
  }
  if (rule.needsInsurance && !fees.insuranceWaived && totals.insurance > 0) {
    extraRows.push({ description: 'Insurance', qty: null, unitPrice: null, subtotal: totals.insurance })
  }
  if (totals.bank > 0) {
    extraRows.push({ description: 'Bank transfer Fee', qty: null, unitPrice: null, subtotal: totals.bank })
  }
  if (totals.other > 0) {
    extraRows.push({ description: 'Other charge', qty: null, unitPrice: null, subtotal: totals.other })
  }
  const rows = [...goods, ...extraRows]
  const rowsTotal = rows.reduce((sum, row) => sum + num(row.subtotal), 0)
  const adjustment = roundTwo(totals.grand - rowsTotal)
  if (Math.abs(adjustment) >= 0.01) {
    rows.push({ description: 'RMB conversion rounding adjustment', qty: null, unitPrice: null, subtotal: adjustment })
  }
  return rows
}

function downloadWorkbook(workbook, filename) {
  return workbook.xlsx.writeBuffer().then((buffer) => {
    saveAs(new Blob([buffer]), filename)
  })
}

async function exportPdf(kind, payload) {
  const { customer, doc, rows, totals, companyProfile } = payload
  const title = kind === 'PI' ? 'PROFORMA INVOICE' : kind === 'CI' ? 'COMMERCIAL INVOICE' : 'QUOTATION'
  const pdf = new jsPDF({ unit: 'pt', format: 'a4' })
  const width = pdf.internal.pageSize.getWidth()
  const logoDataUrl = await loadAssetDataUrl('/assets/logo-mark.png')

  const fitText = (text, x, y, maxWidth, { size = 12, minSize = 8, style = 'normal', align = 'left', lineHeight = 1.15 } = {}) => {
    const value = String(text || '')
    let currentSize = size
    pdf.setFont('helvetica', style)
    pdf.setFontSize(currentSize)
    while (currentSize > minSize && pdf.getTextWidth(value) > maxWidth) {
      currentSize -= 0.5
      pdf.setFontSize(currentSize)
    }
    const lines = pdf.splitTextToSize(value, maxWidth)
    pdf.text(lines, x, y, { align, lineHeightFactor: lineHeight })
    return lines.length * currentSize * lineHeight
  }

  const labelValue = (label, value, x, y, labelWidth, valueWidth, valueSize = 12) => {
    if (!value) return 0
    pdf.setFont('helvetica', 'bold')
    pdf.setFontSize(12)
    pdf.text(label, x + labelWidth, y, { align: 'right' })
    return fitText(value, x + labelWidth + 5, y, valueWidth, { size: valueSize, minSize: 8 })
  }

  const compactLabelValue = (label, value, x, y, valueWidth, valueSize = 12) => {
    pdf.setFont('helvetica', 'bold')
    pdf.setFontSize(12)
    pdf.text(label, x, y)
    const labelGap = pdf.getTextWidth(label) + 5
    return fitText(value, x + labelGap, y, valueWidth - labelGap, { size: valueSize, minSize: 8 })
  }

  pdf.setFont('helvetica', 'bold')
  pdf.setFontSize(18)
  pdf.addImage(logoDataUrl, 'PNG', 105, 48, 34, 34)
  pdf.text(companyProfile.name, 152, 68)
  pdf.setFont('helvetica', 'normal')
  pdf.setFontSize(10.5)
  pdf.text(companyProfile.address, width / 2, 96, { align: 'center' })
  pdf.text(`Tel: ${companyProfile.tel}`, width * 0.36, 114, { align: 'center' })
  pdf.text(`Fax: ${companyProfile.fax}`, width * 0.67, 114, { align: 'center' })
  pdf.setFont('helvetica', 'bold')
  pdf.setFontSize(15)
  pdf.text(title, width / 2, 136, { align: 'center' })

  const noLabel = kind === 'CI' ? 'CI NO. :' : kind === 'PI' ? 'PI NO. :' : 'QUOTE NO. :'
  let leftY = 160
  leftY += Math.max(24, compactLabelValue('Company:', customer.company, 40, leftY, 300, 12) + 10)
  leftY += Math.max(24, compactLabelValue('ATTN:', customer.attn, 40, leftY, 300, 12) + 10)
  leftY += Math.max(28, compactLabelValue('Add.', customer.address, 40, leftY, 300, 10.5) + 12)
  compactLabelValue('Tel:', customer.tel, 40, leftY, 300, 11)
  labelValue(noLabel, doc.no, 315, 165, 70, 175, 12)
  labelValue('Date:', formatDate(doc.date), 315, 192, 70, 175, 12)
  labelValue('By:', doc.by, 315, 220, 70, 175, 12)
  labelValue('From:', doc.from, 315, 250, 52, 90, 11)
  labelValue('To', doc.to, 430, 250, 24, 105, 11)
  const tableStartY = Math.max(262, leftY + 18)

  autoTable(pdf, {
    startY: tableStartY,
    margin: { left: 20, right: 20 },
    head: [['Item', 'Description', 'QTY(KG)', 'Unit Price (USD)', 'Subtotal']],
    body: rows.map((row, index) => [
      index + 1,
      row.description,
      row.qty === null ? '***' : `${row.qty}KG`,
      row.unitPrice === null ? '***' : documentMoney(row.unitPrice),
      documentMoney(row.subtotal),
    ]),
    foot: [['', '', '', '', documentMoney(totals.grand)]],
    theme: 'grid',
    styles: {
      font: 'helvetica',
      fontSize: 9,
      cellPadding: { top: 5, right: 3, bottom: 5, left: 3 },
      lineColor: [0, 0, 0],
      lineWidth: 0.75,
      textColor: [0, 0, 0],
      valign: 'middle',
      overflow: 'linebreak',
    },
    headStyles: { fillColor: [255, 255, 255], textColor: [0, 0, 0], fontStyle: 'bold', halign: 'center' },
    footStyles: { fillColor: [255, 255, 255], textColor: [0, 0, 0], fontStyle: 'normal', halign: 'center' },
    columnStyles: {
      0: { cellWidth: 34, halign: 'center' },
      1: { cellWidth: 260, halign: 'left' },
      2: { cellWidth: 62, halign: 'center' },
      3: { cellWidth: 92, halign: 'center' },
      4: { cellWidth: 98, halign: 'center' },
    },
    didParseCell: (data) => {
      const text = Array.isArray(data.cell.text) ? data.cell.text.join(' ') : String(data.cell.text || '')
      if (text.length > 42 && data.column.index === 1) data.cell.styles.fontSize = 8
      if (text.length > 70 && data.column.index === 1) data.cell.styles.fontSize = 7
    },
  })

  const hasBankInfo = kind === 'PI' && Boolean(companyProfile.bank.accountNo || companyProfile.bank.bankName || companyProfile.bank.swift)
  const bankY = pdf.lastAutoTable.finalY + 22
  let bankEndY = bankY
  if (hasBankInfo) {
    pdf.setFont('helvetica', 'bold')
    pdf.setFontSize(9)
    pdf.text('BANK ACCOUNT:', 40, bankY)
    let cursorY = bankY + 15
    ;[
      `Beneficiary Name :   ${companyProfile.name.toUpperCase()}`,
      `Address of Beneficiary :   ${companyProfile.address}`,
      `Beneficiary Account No. :  ${companyProfile.bank.accountNo || ''}`,
      `Beneficiary bank:  ${companyProfile.bank.bankName || ''}`,
      `SWIFT Code:  ${companyProfile.bank.swift || ''}`,
      `Bank Address:  ${companyProfile.bank.bankAddress || ''}`,
    ].forEach((line) => {
      const height = fitText(line, 40, cursorY, 515, { size: 8.5, minSize: 7, lineHeight: 1.1 })
      cursorY += Math.max(14, height + 3)
    })
    bankEndY = cursorY
  }

  const signY = hasBankInfo ? Math.max(535, bankEndY + 32) : Math.max(455, bankY + 42)
  pdf.setFont('helvetica', 'normal')
  pdf.setFontSize(11)
  pdf.text('The seller', 55, signY)
  pdf.text('The buyer', 385, signY)
  pdf.text(companyProfile.name, 55, signY + 45)
  pdf.line(55, signY + 49, 235, signY + 49)
  fitText(customer.buyer || customer.company || '', 385, signY + 45, 170, { size: 11, minSize: 8 })
  pdf.line(385, signY + 49, 555, signY + 49)
  pdf.setFontSize(10)
  pdf.text('Page  1  ,  Total  1  Pages', width / 2, 820, { align: 'center' })

  const cleanNo = doc.no.replace(/[^\w-]/g, '') || `${kind}-${doc.date || new Date().toISOString().slice(0, 10)}`
  pdf.save(`${cleanNo}-${kind}.pdf`)
}

function Input({ label, value, onChange, type = 'text', step, placeholder }) {
  return (
    <label className="field">
      <span>{label}</span>
      <input
        type={type}
        step={step}
        value={value ?? ''}
        placeholder={placeholder}
        onChange={(event) => onChange(event.target.value)}
      />
    </label>
  )
}

function CurrencyInput({ label, value, currency, onValueChange, onCurrencyChange, hiddenCurrency = false }) {
  return (
    <div className={`currency-field ${hiddenCurrency ? 'single' : ''}`}>
      <Input label={label} type="number" step="0.01" value={value} placeholder="0.00" onChange={onValueChange} />
      {!hiddenCurrency && (
        <label className="field">
          <span>币种</span>
          <select value={currency ?? 'USD'} onChange={(event) => onCurrencyChange(event.target.value)}>
            <option>USD</option>
            <option>CNY</option>
          </select>
        </label>
      )}
    </div>
  )
}

function App() {
  const [customer, setCustomer] = useState(defaultCustomer)
  const [doc, setDoc] = useState(defaultDoc)
  const [fees, setFees] = useState(defaultFees)
  const [items, setItems] = useState([blankItem()])
  const [status, setStatus] = useState('准备生成单据')
  const [generatedNos, setGeneratedNos] = useState(readGeneratedNos)

  const activeRule = termRules[fees.incoterm]
  const activeCompanyProfile = companyProfiles[doc.companyProfileId] || companyProfiles.shenzhen

  const totals = useMemo(() => {
    const goodsCny = items.reduce(
      (sum, item) => sum + (item.currency === 'CNY' ? num(item.qty) * num(item.unitPrice) : 0),
      0,
    )
    const goodsUsd = items.reduce(
      (sum, item) => sum + (item.currency === 'USD' ? num(item.qty) * num(item.unitPrice) : 0),
      0,
    )
    const qty = items.reduce((sum, item) => sum + num(item.qty), 0)
    const localCny = activeRule.needsLocal && fees.localCurrency === 'CNY' ? num(fees.local) : 0
    const localUsd = activeRule.needsLocal && fees.localCurrency === 'USD' ? num(fees.local) : 0
    const oceanCny = activeRule.needsOcean && fees.oceanCurrency === 'CNY' ? num(fees.ocean) : 0
    const oceanUsd = activeRule.needsOcean && fees.oceanCurrency === 'USD' ? num(fees.ocean) : 0
    const bankCny = fees.bankCurrency === 'CNY' ? num(fees.bank) : 0
    const bankUsd = fees.bankCurrency === 'USD' ? num(fees.bank) : 0
    const otherCny = fees.otherCurrency === 'CNY' ? num(fees.other) : 0
    const otherUsd = fees.otherCurrency === 'USD' ? num(fees.other) : 0

    const rmbBase = goodsCny + localCny + oceanCny
    const rmbAddons = bankCny + otherCny
    const rmbBaseUsd = roundedRmbToUsd(rmbBase, fees.exchangeRate)
    const rmbAddonsUsd = roundedRmbToUsd(rmbAddons, fees.exchangeRate)
    const goods = goodsUsd + roundedRmbToUsd(goodsCny, fees.exchangeRate)
    const local = localUsd + roundedRmbToUsd(localCny, fees.exchangeRate)
    const ocean = oceanUsd + roundedRmbToUsd(oceanCny, fees.exchangeRate)
    const bank = bankUsd + roundedRmbToUsd(bankCny, fees.exchangeRate)
    const other = otherUsd + roundedRmbToUsd(otherCny, fees.exchangeRate)
    const fob = rmbBaseUsd + goodsUsd + localUsd
    const cfr = rmbBaseUsd + goodsUsd + localUsd + oceanUsd
    const insurance = activeRule.needsInsurance && !fees.insuranceWaived ? cfr * num(fees.insuranceRate) * 1.1 : 0
    const cif = cfr + insurance
    const baseByTerm = { EXW: goods, FOB: fob, CFR: cfr, CIF: cif }[fees.incoterm] || goods
    const grand = baseByTerm + rmbAddonsUsd + bankUsd + otherUsd
    const cifPerKg = qty > 0 ? roundTwo(grand / qty) : 0
    return {
      goods,
      goodsCny,
      goodsUsd,
      qty,
      local,
      ocean,
      bank,
      other,
      rmbBase,
      rmbAddons,
      rmbBaseUsd,
      rmbAddonsUsd,
      usdBase: goodsUsd + localUsd + oceanUsd,
      usdAddons: bankUsd + otherUsd,
      fob,
      cfr,
      insurance,
      cif,
      cifPerKg,
      grand,
    }
  }, [activeRule, fees, items])

  const payload = { customer, doc, rows: [], totals, fees, companyProfile: activeCompanyProfile }

  const updateCustomer = (key, value) => setCustomer((current) => ({ ...current, [key]: value }))
  const updateDoc = (key, value) => setDoc((current) => ({ ...current, [key]: value }))
  const updateFees = (key, value) => setFees((current) => ({ ...current, [key]: value }))
  const updateItem = (id, key, value) =>
    setItems((current) => current.map((item) => (item.id === id ? { ...item, [key]: value } : item)))
  const generatedCustomerCode = suggestCustomerCode(customer, doc.customerType)
  const activeCustomerCode = doc.customerCode || generatedCustomerCode
  const generatedDocNo = buildDocumentNo({ ...doc, customerCode: activeCustomerCode }, customer)
  const docNoLooksAuto = /^ALL\d{8}(?:-[A-Z]{2}[A-Z0-9]{1,4}\d{2})?$/i.test(doc.no)
  const activeDocNo = !doc.no || docNoLooksAuto ? generatedDocNo : doc.no
  const isDuplicateNo = generatedNos.includes(activeDocNo)

  const resetForm = () => {
    setCustomer(defaultCustomer)
    setDoc(defaultDoc)
    setFees(defaultFees)
    setItems([blankItem()])
    setStatus('准备生成单据')
  }

  const rememberNo = () => {
    if (activeDocNo && !generatedNos.includes(activeDocNo)) {
      const nextNos = [...generatedNos, activeDocNo]
      setGeneratedNos(nextNos)
      saveGeneratedNos(nextNos)
    }
  }

  const exportDocument = async (kind) => {
    setStatus(`正在生成 ${kind} Excel...`)
    const isCi = kind === 'CI'
    const template = isCi ? 'commercial-invoice-template.xlsx' : 'proforma-invoice-template.xlsx'
    const workbook = await loadWorkbook(template)
    const sheet = workbook.getWorksheet('PI') || workbook.worksheets[0]
    const title = kind === 'PI' ? 'PROFORMA INVOICE' : kind === 'CI' ? 'COMMERCIAL INVOICE' : 'QUOTATION'
    fillInvoice(
      sheet,
      {
        ...payload,
        doc: { ...doc, no: activeDocNo },
        rows: buildRows(items, fees, totals),
      },
      title,
    )
    await addSheetLogo(workbook, sheet)
    const cleanNo = activeDocNo.replace(/[^\w-]/g, '') || `${kind}-${doc.date || new Date().toISOString().slice(0, 10)}`
    await downloadWorkbook(workbook, `${cleanNo}-${kind}.xlsx`)
    rememberNo()
    setStatus(`${kind} Excel 已生成，可在浏览器下载记录里查看`)
  }

  const exportDocumentPdf = async (kind) => {
    setStatus(`正在生成 ${kind} PDF...`)
    await exportPdf(kind, {
      ...payload,
      doc: { ...doc, no: activeDocNo },
      rows: buildRows(items, fees, totals),
    })
    rememberNo()
    setStatus(`${kind} PDF 已生成，可在浏览器下载记录里查看`)
  }

  return (
    <main className="app-shell">
      <header className="topbar">
        <div className="brand">
          <img className="brand-logo" src="/assets/company-logo.png" alt="ARGIOPE 金蛛王" />
          <div>
            <h1>报价计算器</h1>
            <p>Shenzhen Jindaquan PI / CI / Quotation</p>
          </div>
        </div>
        <button className="ghost-button" type="button" onClick={resetForm}>
          <RefreshCw size={16} aria-hidden="true" />
          清空重置
        </button>
      </header>

      <section className="workspace">
        <aside className="panel form-panel">
          <div className="panel-title">
            <h2>客户信息</h2>
            <span>Buyer</span>
          </div>
          <Input label="Company" value={customer.company} onChange={(value) => updateCustomer('company', value)} />
          <Input label="ATTN" value={customer.attn} onChange={(value) => updateCustomer('attn', value)} />
          <label className="field">
            <span>Address</span>
            <textarea value={customer.address} onChange={(event) => updateCustomer('address', event.target.value)} />
          </label>
          <Input label="Tel" value={customer.tel} onChange={(value) => updateCustomer('tel', value)} />
          <Input label="Buyer sign" value={customer.buyer} onChange={(value) => updateCustomer('buyer', value)} />

          <div className="panel-title spaced">
            <h2>单据信息</h2>
            <span>Document</span>
          </div>
          <label className="field">
            <span>公司主体</span>
            <select value={doc.companyProfileId} onChange={(event) => updateDoc('companyProfileId', event.target.value)}>
              {Object.values(companyProfiles).map((profile) => (
                <option key={profile.id} value={profile.id}>
                  {profile.label}
                </option>
              ))}
            </select>
          </label>
          <Input label="Date" type="date" value={doc.date} onChange={(value) => updateDoc('date', value)} />
          <Input
            label="单据流水号"
            type="number"
            value={doc.documentSeq}
            placeholder="01"
            onChange={(value) => updateDoc('documentSeq', value)}
          />
          <div className="serial-box">
            <div className="serial-preview">
              <span>自动单号</span>
              <strong>{generatedDocNo}</strong>
            </div>
            <button type="button" className="small-button" onClick={() => updateDoc('no', generatedDocNo)}>
              填入单号
            </button>
          </div>
          {isDuplicateNo && <p className="duplicate-warning">这个单号已在本机生成过，请调整单据流水号。</p>}
          <Input label="PI / Quote No." value={activeDocNo} onChange={(value) => updateDoc('no', value)} />
          <label className="field">
            <span>客户类型</span>
            <select value={doc.customerType} onChange={(event) => updateDoc('customerType', event.target.value)}>
              <option value="company">公司采购</option>
              <option value="personal">个人采购</option>
            </select>
          </label>
          <div className="two-col">
            <Input
              label="国家代码"
              value={doc.countryCode}
              placeholder="IT"
              onChange={(value) => updateDoc('countryCode', value.toUpperCase().slice(0, 2))}
            />
            <Input
              label="客户编码"
              value={activeCustomerCode}
              placeholder={generatedCustomerCode || 'ABC'}
              onChange={(value) => updateDoc('customerCode', value.toUpperCase().replace(/[^A-Z0-9]/g, '').slice(0, 4))}
            />
          </div>
          <Input
            label="该客户第几次订单"
            type="number"
            value={doc.customerOrderSeq}
            placeholder="02"
            onChange={(value) => updateDoc('customerOrderSeq', value)}
          />
          <div className="two-col">
            <Input label="By" value={doc.by} onChange={(value) => updateDoc('by', value)} />
            <Input label="From" value={doc.from} onChange={(value) => updateDoc('from', value)} />
          </div>
          <Input label="To" value={doc.to} onChange={(value) => updateDoc('to', value)} />
        </aside>

        <section className="panel main-panel">
          <div className="panel-title">
            <h2>货品明细</h2>
            <button type="button" className="small-button" onClick={() => setItems((current) => [...current, blankItem()])}>
              <Plus size={16} aria-hidden="true" />
              添加行
            </button>
          </div>

          <div className="item-table" role="table" aria-label="货品明细">
            <div className="table-head" role="row">
              <span>Description</span>
              <span>HS CODE</span>
              <span>QTY(KG)</span>
              <span>Unit Price</span>
              <span>币种</span>
              <span>Subtotal USD</span>
              <span></span>
            </div>
            {items.map((item) => (
              <div className="table-row" role="row" key={item.id}>
                <input
                  value={item.description}
                  placeholder="输入产品描述"
                  onChange={(event) => updateItem(item.id, 'description', event.target.value)}
                />
                <input
                  value={item.hsCode ?? ''}
                  placeholder="HS CODE"
                  onChange={(event) => updateItem(item.id, 'hsCode', event.target.value)}
                />
                <input
                  type="number"
                  step="0.01"
                  value={item.qty}
                  placeholder="0"
                  onChange={(event) => updateItem(item.id, 'qty', event.target.value)}
                />
                <input
                  type="number"
                  step="0.01"
                  value={item.unitPrice}
                  placeholder="0.00"
                  onChange={(event) => updateItem(item.id, 'unitPrice', event.target.value)}
                />
                <select value={item.currency} onChange={(event) => updateItem(item.id, 'currency', event.target.value)}>
                  <option>USD</option>
                  <option>CNY</option>
                </select>
                <strong>{money(num(item.qty) * toUsd(item.unitPrice, item.currency, fees.exchangeRate))}</strong>
                <button
                  className="icon-button"
                  type="button"
                  aria-label="删除货品行"
                  onClick={() => setItems((current) => current.filter((row) => row.id !== item.id))}
                  disabled={items.length === 1}
                >
                  <Trash2 size={16} aria-hidden="true" />
                </button>
              </div>
            ))}
          </div>

          <div className="rules-grid">
            <div className="rule-card">
              <div className="rule-heading">
                <Calculator size={18} aria-hidden="true" />
                <h2>费用规则</h2>
              </div>
              <label className="field">
                <span>报价术语</span>
                <select value={fees.incoterm} onChange={(event) => updateFees('incoterm', event.target.value)}>
                  {Object.keys(termRules).map((term) => (
                    <option key={term}>{term}</option>
                  ))}
                </select>
              </label>
              <Input
                label="汇率：1 USD = RMB"
                type="number"
                step="0.0001"
                value={fees.exchangeRate}
                placeholder="6.5"
                onChange={(value) => updateFees('exchangeRate', value)}
              />
              <div className="term-note">
                <strong>{activeRule.title}</strong>
                <p>{activeRule.note}</p>
              </div>

              {activeRule.needsLocal && (
                <CurrencyInput
                  label="本地费用 / 出口前费用"
                  value={fees.local}
                  currency={fees.localCurrency}
                  onValueChange={(value) => updateFees('local', value)}
                  onCurrencyChange={(value) => updateFees('localCurrency', value)}
                />
              )}
              {activeRule.needsOcean && (
                <CurrencyInput
                  label="海运费 / Freight"
                  value={fees.ocean}
                  currency={fees.oceanCurrency}
                  onValueChange={(value) => updateFees('ocean', value)}
                  onCurrencyChange={(value) => updateFees('oceanCurrency', value)}
                />
              )}
              {activeRule.needsInsurance && (
                <>
                  <label className="switch-row">
                    <input
                      type="checkbox"
                      checked={fees.insuranceWaived}
                      onChange={(event) => updateFees('insuranceWaived', event.target.checked)}
                    />
                    <span>保险由船公司赠送 / 不计费</span>
                  </label>
                  {!fees.insuranceWaived && (
                    <Input
                      label="保险费率"
                      type="number"
                      step="0.0001"
                      value={fees.insuranceRate}
                      placeholder="0.003"
                      onChange={(value) => updateFees('insuranceRate', value)}
                    />
                  )}
                </>
              )}

              <div className="optional-fees">
                <h3>单据附加费用</h3>
                <CurrencyInput
                  label="银行手续费"
                  value={fees.bank}
                  currency={fees.bankCurrency}
                  onValueChange={(value) => updateFees('bank', value)}
                  onCurrencyChange={(value) => updateFees('bankCurrency', value)}
                />
                <CurrencyInput
                  label="其他费用"
                  value={fees.other}
                  currency={fees.otherCurrency}
                  onValueChange={(value) => updateFees('other', value)}
                  onCurrencyChange={(value) => updateFees('otherCurrency', value)}
                />
              </div>
            </div>

            <div className="rule-card">
              <div className="rule-heading">
                <Ship size={18} aria-hidden="true" />
                <h2>条款备注</h2>
              </div>
              <label className="field">
                <span>Payment</span>
                <textarea value={doc.payment} onChange={(event) => updateDoc('payment', event.target.value)} />
              </label>
              <label className="field">
                <span>Lead-time</span>
                <textarea value={doc.leadTime} onChange={(event) => updateDoc('leadTime', event.target.value)} />
              </label>
            </div>
          </div>
        </section>

        <aside className="panel summary-panel">
          <div className="panel-title">
            <h2>即时报价</h2>
            <span>{fees.incoterm}</span>
          </div>
          <dl className="totals">
            <div>
              <dt>EXW 货值</dt>
              <dd>{money(totals.goods)}</dd>
            </div>
            {activeRule.needsLocal && (
              <div>
                <dt>FOB = EXW + 本地费用</dt>
                <dd>{money(totals.fob)}</dd>
              </div>
            )}
            {activeRule.needsOcean && (
              <div>
                <dt>CFR = FOB + 海运费</dt>
                <dd>{money(totals.cfr)}</dd>
              </div>
            )}
            {activeRule.needsInsurance && (
              <div>
                <dt>保险费</dt>
                <dd>{fees.insuranceWaived ? 'Waived' : money(totals.insurance)}</dd>
              </div>
            )}
            <div>
              <dt>人民币报价池</dt>
              <dd>{cny(totals.rmbBase)} → {money(totals.rmbBaseUsd)}</dd>
            </div>
            {(totals.rmbAddons > 0 || totals.rmbAddonsUsd > 0) && (
              <div>
                <dt>人民币附加费池</dt>
                <dd>{cny(totals.rmbAddons)} → {money(totals.rmbAddonsUsd)}</dd>
              </div>
            )}
            <div>
              <dt>最终每公斤</dt>
              <dd>{money(totals.cifPerKg)} / KG</dd>
            </div>
            <div>
              <dt>最终每公斤 RMB</dt>
              <dd>{cny(fromUsd(totals.cifPerKg, 'CNY', fees.exchangeRate))} / KG</dd>
            </div>
            <div>
              <dt>银行/其他附加</dt>
              <dd>{money(totals.bank + totals.other)}</dd>
            </div>
            <div className="grand">
              <dt>最终报价</dt>
              <dd>{money(totals.grand)}</dd>
            </div>
          </dl>
          <div className="export-stack">
            <button type="button" onClick={() => exportDocument('PI')}>
              <Download size={17} aria-hidden="true" />
              Excel PI
            </button>
            <button type="button" onClick={() => exportDocument('CI')}>
              <Download size={17} aria-hidden="true" />
              Excel CI
            </button>
            <button type="button" onClick={() => exportDocument('QUOTATION')}>
              <Download size={17} aria-hidden="true" />
              Excel 报价单
            </button>
          </div>
          <div className="export-stack pdf-stack">
            <button type="button" onClick={() => exportDocumentPdf('PI')}>
              <FileText size={17} aria-hidden="true" />
              PDF PI
            </button>
            <button type="button" onClick={() => exportDocumentPdf('CI')}>
              <FileText size={17} aria-hidden="true" />
              PDF CI
            </button>
            <button type="button" onClick={() => exportDocumentPdf('QUOTATION')}>
              <FileText size={17} aria-hidden="true" />
              PDF 报价单
            </button>
          </div>
          <p className="status">{status}</p>
        </aside>
      </section>
    </main>
  )
}

export default App
