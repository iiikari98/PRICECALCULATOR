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

const companyName = 'Shenzhen Jindaquan Technology Co.,Ltd'
const companyAddress = 'Bld 16#, Wulian Xieping Ailian Industrial, Longcheng Longgang District, Shenzhen, China'
const serialStorageKey = 'jdq-generated-document-nos'

const blankItem = () => ({
  id: crypto.randomUUID(),
  description: '',
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
  by: '',
  customerType: 'company',
  companyDailySeq: '',
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

function firstCustomerInitial(customer) {
  const source = `${customer.company || customer.attn || customer.buyer || ''}`.trim()
  const match = source.match(/[A-Za-z0-9]/)
  return match ? match[0].toUpperCase() : ''
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
    const initials = words.map((word) => word[0]).join('')
    return (initials + chars).slice(0, 3)
  }
  if (chars.length <= 3) return chars
  return `${chars[0]}${chars[1]}${chars[chars.length - 1]}`.slice(0, 3)
}

function buildDocumentNo(doc, customer) {
  const customerCode = (doc.customerCode || suggestCustomerCode(customer, doc.customerType) || firstCustomerInitial(customer)).toUpperCase()
  return [
    'ALL',
    formatCodeDate(doc.date),
    twoDigits(doc.companyDailySeq),
    '-',
    (doc.countryCode || '').toUpperCase(),
    customerCode,
    twoDigits(doc.customerOrderSeq),
  ].join('')
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

function setCell(sheet, address, value) {
  const cell = sheet.getCell(address)
  cell.value = value
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

function clearTemplateSampleData(sheet) {
  ;['D5', 'D6', 'D7', 'D8', 'H5', 'H6', 'H7', 'G8', 'I8', 'G19', 'G28'].forEach((cell) => {
    setCell(sheet, cell, '')
  })
  clearInvoiceRows(sheet, 10, 18)
}

function fillInvoice(sheet, payload, title) {
  const { customer, doc, rows, totals, fees } = payload
  clearTemplateSampleData(sheet)
  setCell(sheet, title === 'COMMERCIAL INVOICE' ? 'D4' : 'B4', title)
  setCell(sheet, 'D5', customer.company)
  setCell(sheet, 'D6', customer.attn)
  setCell(sheet, 'D7', customer.address)
  setCell(sheet, 'D8', customer.tel)
  setCell(sheet, 'H5', doc.no)
  setCell(sheet, 'H6', formatDate(doc.date))
  setCell(sheet, 'H7', doc.by)
  setCell(sheet, 'G8', doc.from)
  setCell(sheet, 'I8', doc.to)

  const neededRows = Math.max(rows.length + 2, 3)
  const baseTotalRow = 14
  if (neededRows > 4) {
    sheet.spliceRows(baseTotalRow, 0, ...Array.from({ length: neededRows - 4 }, () => []))
    for (let row = baseTotalRow; row < baseTotalRow + neededRows - 4; row += 1) {
      for (let col = 1; col <= 13; col += 1) {
        copyStyle(sheet.getRow(13).getCell(col), sheet.getRow(row).getCell(col))
      }
    }
  }

  clearInvoiceRows(sheet, 10, 14 + Math.max(0, neededRows - 4))

  rows.forEach((row, index) => {
    const excelRow = 10 + index
    setCell(sheet, `B${excelRow}`, index + 1)
    setCell(sheet, `C${excelRow}`, row.description)
    setCell(sheet, `E${excelRow}`, row.qty === null ? '***' : row.qty)
    setCell(sheet, `G${excelRow}`, row.unitPrice === null ? '***' : money(row.unitPrice))
    setCell(sheet, `I${excelRow}`, money(row.subtotal))
  })

  const totalRow = 10 + rows.length
  safeUnmerge(sheet, `B${totalRow}:I${totalRow}`)
  setCell(sheet, `B${totalRow}`, 'Total')
  setCell(sheet, `E${totalRow}`, totals.qty || '')
  setCell(sheet, `G${totalRow}`, '***')
  setCell(sheet, `I${totalRow}`, money(totals.grand))

  setCell(sheet, `B${totalRow + 1}`, amountWords(totals.grand))
  setCell(sheet, `B${totalRow + 2}`, 'Payment: ')
  setCell(sheet, `D${totalRow + 2}`, doc.payment)
  setCell(sheet, `B${totalRow + 3}`, 'Lead-time: ')
  setCell(sheet, `D${totalRow + 3}`, doc.leadTime)

  const rule = termRules[fees.incoterm]
  const insuranceNote = rule.needsInsurance
    ? fees.insuranceWaived
      ? 'Insurance: waived / included by shipping company'
      : `Insurance: ${(num(fees.insuranceRate) * 100).toFixed(3)}%`
    : ''
  setCell(sheet, `B${totalRow + 4}`, insuranceNote)
  setCell(sheet, `B${totalRow + 5}`, `Exchange rate: 1 USD = RMB ${num(fees.exchangeRate).toFixed(4)}`)
  setCell(sheet, `B${totalRow + 6}`, `Final unit price: ${money(totals.cifPerKg)} / KG`)

  const buyerRow = title === 'COMMERCIAL INVOICE' ? 19 : Math.max(28, totalRow + 14)
  setCell(sheet, `G${buyerRow}`, customer.buyer || customer.company)
}

function buildRows(items, fees, totals) {
  const rule = termRules[fees.incoterm]
  const goods = items
    .filter((item) => item.description || num(item.qty) || num(item.unitPrice))
    .map((item) => {
      const unitPriceUsd = toUsd(item.unitPrice, item.currency, fees.exchangeRate)
      return {
        description: item.description,
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

function exportPdf(kind, payload) {
  const { customer, doc, rows, totals, fees } = payload
  const title = kind === 'PI' ? 'PROFORMA INVOICE' : kind === 'CI' ? 'COMMERCIAL INVOICE' : 'QUOTATION'
  const pdf = new jsPDF({ unit: 'pt', format: 'a4' })
  const width = pdf.internal.pageSize.getWidth()

  pdf.setFont('helvetica', 'bold')
  pdf.setFontSize(15)
  pdf.text(companyName, 40, 42)
  pdf.setFont('helvetica', 'normal')
  pdf.setFontSize(8.5)
  pdf.text(companyAddress, 40, 58)
  pdf.text('Tel: +86-0755-28996208    Fax: +86-0755-28994568', 40, 72)
  pdf.setFont('helvetica', 'bold')
  pdf.setFontSize(17)
  pdf.text(title, width / 2, 104, { align: 'center' })

  pdf.setFontSize(9)
  pdf.setFont('helvetica', 'normal')
  const left = [
    ['Company', customer.company],
    ['ATTN', customer.attn],
    ['Address', customer.address],
    ['Tel', customer.tel],
  ]
  const right = [
    ['No.', doc.no],
    ['Date', formatDate(doc.date)],
    ['By', doc.by],
    ['From / To', `${doc.from || ''} / ${doc.to || ''}`],
  ]
  left.forEach(([label, value], index) => pdf.text(`${label}: ${value || ''}`, 40, 130 + index * 16, { maxWidth: 285 }))
  right.forEach(([label, value], index) => pdf.text(`${label}: ${value || ''}`, 365, 130 + index * 16, { maxWidth: 180 }))

  autoTable(pdf, {
    startY: 205,
    head: [['Item', 'Description', 'QTY(KG)', 'Unit Price (USD)', 'Subtotal']],
    body: rows.map((row, index) => [
      index + 1,
      row.description,
      row.qty === null ? '***' : row.qty,
      row.unitPrice === null ? '***' : money(row.unitPrice),
      money(row.subtotal),
    ]),
    foot: [['Total', '', totals.qty || '', '***', money(totals.grand)]],
    styles: { font: 'helvetica', fontSize: 9, cellPadding: 5 },
    headStyles: { fillColor: [18, 32, 51] },
    footStyles: { fillColor: [238, 244, 248], textColor: [23, 33, 47], fontStyle: 'bold' },
  })

  const endY = pdf.lastAutoTable.finalY + 18
  pdf.setFontSize(9)
  pdf.text(amountWords(totals.grand), 40, endY)
  pdf.text(`Payment: ${doc.payment || ''}`, 40, endY + 18)
  pdf.text(`Lead-time: ${doc.leadTime || ''}`, 40, endY + 36)
  pdf.text(`Exchange rate: 1 USD = RMB ${num(fees.exchangeRate).toFixed(4)}`, 40, endY + 54)
  pdf.text(`Final unit price: ${money(totals.cifPerKg)} / KG (${cny(fromUsd(totals.cifPerKg, 'CNY', fees.exchangeRate))} / KG)`, 40, endY + 72)

  pdf.text('The seller', 80, endY + 122)
  pdf.text('The buyer', 390, endY + 122)
  pdf.text(companyName, 60, endY + 140)
  pdf.text(customer.buyer || customer.company || '', 370, endY + 140)

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
        value={value}
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
          <select value={currency} onChange={(event) => onCurrencyChange(event.target.value)}>
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

  const payload = { customer, doc, rows: [], totals, fees }

  const updateCustomer = (key, value) => setCustomer((current) => ({ ...current, [key]: value }))
  const updateDoc = (key, value) => setDoc((current) => ({ ...current, [key]: value }))
  const updateFees = (key, value) => setFees((current) => ({ ...current, [key]: value }))
  const updateItem = (id, key, value) =>
    setItems((current) => current.map((item) => (item.id === id ? { ...item, [key]: value } : item)))
  const generatedDocNo = buildDocumentNo(doc, customer)
  const activeDocNo = doc.no || generatedDocNo
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
    const cleanNo = activeDocNo.replace(/[^\w-]/g, '') || `${kind}-${doc.date || new Date().toISOString().slice(0, 10)}`
    await downloadWorkbook(workbook, `${cleanNo}-${kind}.xlsx`)
    rememberNo()
    setStatus(`${kind} Excel 已生成，可在浏览器下载记录里查看`)
  }

  const exportDocumentPdf = (kind) => {
    setStatus(`正在生成 ${kind} PDF...`)
    exportPdf(kind, {
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
          <Input label="Date" type="date" value={doc.date} onChange={(value) => updateDoc('date', value)} />
          <div className="serial-box">
            <div className="serial-preview">
              <span>自动单号</span>
              <strong>{generatedDocNo}</strong>
            </div>
            <button type="button" className="small-button" onClick={() => updateDoc('no', generatedDocNo)}>
              填入单号
            </button>
          </div>
          {isDuplicateNo && <p className="duplicate-warning">这个单号已在本机生成过，请调整公司当天序号或客户订单序号。</p>}
          <Input label="PI / Quote No." value={doc.no} onChange={(value) => updateDoc('no', value)} />
          <label className="field">
            <span>客户类型</span>
            <select value={doc.customerType} onChange={(event) => updateDoc('customerType', event.target.value)}>
              <option value="company">公司采购</option>
              <option value="personal">个人采购</option>
            </select>
          </label>
          <Input
            label="公司当天第几单"
            type="number"
            value={doc.companyDailySeq}
            placeholder="03"
            onChange={(value) => updateDoc('companyDailySeq', value)}
          />
          <div className="two-col">
            <Input
              label="国家代码"
              value={doc.countryCode}
              placeholder="IT"
              onChange={(value) => updateDoc('countryCode', value.toUpperCase().slice(0, 2))}
            />
            <Input
              label="客户编码"
              value={doc.customerCode}
              placeholder={suggestCustomerCode(customer, doc.customerType) || 'ABC'}
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
