import {
  startTransition,
  useCallback,
  useDeferredValue,
  useEffect,
  useLayoutEffect,
  useMemo,
  useRef,
  useState,
  type CSSProperties,
} from 'react'
import './App.css'

type FormatName = 'hwp' | 'hwpx'
type ZoomPreset = 'fitWidth' | '100' | '125' | '150'

type TextStyle = {
  fontFamily?: string | null
  fontSize?: number | null
  textColor?: string | null
  backgroundColor?: string | null
  widthRatio?: number | null
  letterSpacing?: number | null
  relativeSize?: number | null
  baselineOffset?: number | null
  useFontSpace: boolean
  useKerning: boolean
  bold: boolean
  italic: boolean
  underline: boolean
}

type TextRun = {
  text: string
  style?: TextStyle | null
}

type ParagraphStyle = {
  align?: string | null
  indent?: number | null
  marginLeft?: number | null
  marginRight?: number | null
  marginPrev?: number | null
  marginNext?: number | null
  lineSpacingType?: string | null
  lineSpacing?: number | null
  headingType?: string | null
  headingIdRef?: number | null
  headingLevel?: number | null
  markerAlign?: string | null
  markerWidthAdjust?: number | null
  markerTextOffsetType?: string | null
  markerTextOffset?: number | null
  keepWithNext: boolean
  keepLines: boolean
}

type ParagraphValue = {
  marker?: TextRun | null
  runs: TextRun[]
  style?: ParagraphStyle | null
  lineSegmentCount?: number | null
  layoutHeightHint?: number | null
  pageBreakBefore: boolean
}

type TableCell = {
  text: string
  blocks: DocumentBlock[]
  colSpan?: number | null
  rowSpan?: number | null
  width?: number | null
  height?: number | null
  paddingLeft?: number | null
  paddingRight?: number | null
  paddingTop?: number | null
  paddingBottom?: number | null
  style?: TableCellStyle | null
  isHeader: boolean
}

type TableBorder = {
  style?: string | null
  width?: string | null
  color?: string | null
}

type TableDiagonal = {
  style?: string | null
  width?: string | null
  color?: string | null
  slashType?: string | null
  backSlashType?: string | null
}

type TableCellStyle = {
  backgroundColor?: string | null
  borderLeft?: TableBorder | null
  borderRight?: TableBorder | null
  borderTop?: TableBorder | null
  borderBottom?: TableBorder | null
  diagonal?: TableDiagonal | null
}

type TableRow = {
  cells: TableCell[]
}

type TableValue = {
  rows: TableRow[]
  width?: number | null
  height?: number | null
  cellSpacing?: number | null
  style?: TableCellStyle | null
  noAdjust: boolean
  repeatHeader: boolean
  headerRowCount?: number | null
}

type ImageValue = {
  kind: string
  assetId: string
  altText?: string | null
  width?: number | null
  height?: number | null
  widthRelTo?: string | null
  heightRelTo?: string | null
  treatAsChar: boolean
  textWrap?: string | null
  zOrder?: number | null
  vertRelTo?: string | null
  horzRelTo?: string | null
  vertAlign?: string | null
  horzAlign?: string | null
  vertOffset?: number | null
  horzOffset?: number | null
  distanceLeft?: number | null
  distanceRight?: number | null
  distanceTop?: number | null
  distanceBottom?: number | null
  rotation?: number | null
  caption?: string | null
}

type UnsupportedValue = {
  kind: string
  reason?: string | null
}

type FootnoteValue = {
  kind: string
  number?: number | null
  blocks: DocumentBlock[]
}

type DocumentBlock =
  | { type: 'paragraph'; value: ParagraphValue }
  | { type: 'table'; value: TableValue }
  | { type: 'image'; value: ImageValue }
  | { type: 'footnote'; value: FootnoteValue }
  | { type: 'unsupported'; value: UnsupportedValue }

type PageLayout = {
  width?: number | null
  height?: number | null
  landscape: boolean
  marginLeft?: number | null
  marginRight?: number | null
  marginTop?: number | null
  marginBottom?: number | null
  marginHeader?: number | null
  marginFooter?: number | null
  marginGutter?: number | null
  pageBorder?: TableCellStyle | null
}

type DocumentSection = {
  id: number
  blocks: DocumentBlock[]
  pageLayout?: PageLayout | null
  headers: HeaderFooter[]
  footers: HeaderFooter[]
  pageStartNumber?: number | null
}

type HeaderFooter = {
  applyPageType?: string | null
  blocks: DocumentBlock[]
}

type DocumentMetadata = {
  title?: string | null
  author?: string | null
  pageCount?: number | null
  language?: string | null
}

type AssetRef = {
  id: string
  mediaType: string
  sourcePath?: string | null
  dataUri?: string | null
}

type DocumentModel = {
  format?: FormatName | null
  metadata: DocumentMetadata
  sections: DocumentSection[]
  assets: AssetRef[]
}

type DocumentDiagnostics = {
  format: FormatName
  sectionCount: number
  isEncrypted: boolean
}

type DetailedLayoutSummary = {
  sectionCount: number
  paragraphCount: number
  tableCount: number
  imageCount: number
  unsupportedCount: number
}

type LoadedDocument = {
  fileName: string
  document: DocumentModel
  diagnostics: DocumentDiagnostics
  layout: DetailedLayoutSummary
  plainText: string
}

type PlaceholderContext = {
  pageNumber: number
  totalPages: number
}

type RenderedPage = {
  key: string
  sectionId: number
  displayPageNumber: number
  sectionPageNumber: number
  pageLayout?: PageLayout | null
  blocks: DocumentBlock[]
  header?: HeaderFooter | null
  footer?: HeaderFooter | null
}

type PaginationItem = {
  key: string
  block: DocumentBlock
  actualHeight: number
  hintHeight: number | null
  pageBreakBefore: boolean
  keepWithNext: boolean
  keepLines: boolean
  splitMeasurement?: SplittableBlockMeasurement | null
}

type SplittableBlockMeasurement =
  | SingleCellTableSplitMeasurement
  | TableRowSplitMeasurement
  | ParagraphSplitMeasurement

type SingleCellTableSplitMeasurement = {
  kind: 'singleCellTable'
  chromeHeight: number
  chromeHintHeight: number
  childItems: Array<{
    block: DocumentBlock
    actualHeight: number
    hintHeight: number | null
  }>
}

type TableRowSplitMeasurement = {
  kind: 'tableRows'
  chromeHeight: number
  chromeHintHeight: number
  headerRowCount: number
  safeBreakStarts: Set<number>
  rowItems: Array<{
    row: TableRow
    actualHeight: number
    hintHeight: number | null
  }>
}

type ParagraphSplitMeasurement = {
  kind: 'paragraphLines'
  lineRuns: TextRun[][]
  lineActualHeights: number[]
  lineHintHeights: number[]
  topMarginHeight: number
  bottomMarginHeight: number
}

const DEFAULT_PAGE_WIDTH = 794
const DEFAULT_PAGE_HEIGHT = 1123
const PAGINATION_TOLERANCE = 24
const MIN_PARAGRAPH_FRAGMENT_LINES = 3

async function openWithNativeDialog(): Promise<LoadedDocument | null> {
  const { invoke } = await import('@tauri-apps/api/core')
  return await invoke<LoadedDocument | null>('pick_and_open_document')
}

async function openWithBytes(fileName: string, bytes: Uint8Array): Promise<LoadedDocument | null> {
  const { invoke } = await import('@tauri-apps/api/core')
  return await invoke<LoadedDocument | null>('open_document', {
    fileName,
    bytes: Array.from(bytes),
  })
}

const RECENT_FILES_KEY = 'max-viewer-recent-files'
const MAX_RECENT_FILES = 10
type RecentFile = { name: string; openedAt: string }

function loadRecentFiles(): RecentFile[] {
  try {
    const raw = localStorage.getItem(RECENT_FILES_KEY)
    return raw ? JSON.parse(raw) : []
  } catch {
    return []
  }
}

function saveRecentFile(name: string): void {
  const recent = loadRecentFiles().filter((f) => f.name !== name)
  recent.unshift({ name, openedAt: new Date().toISOString() })
  if (recent.length > MAX_RECENT_FILES) recent.length = MAX_RECENT_FILES
  try {
    localStorage.setItem(RECENT_FILES_KEY, JSON.stringify(recent))
  } catch { /* ignore quota errors */ }
}

function formatLabel(format?: FormatName | null): string {
  return format?.toUpperCase() ?? 'Unknown'
}

function hwpToPx(value?: number | null): number | undefined {
  if (value == null) {
    return undefined
  }

  return Math.max(0, Math.round((value / 7200) * 96))
}

function fontSizeToCss(value?: number | null, zoomFactor = 1): string | undefined {
  if (value == null || value <= 0) {
    return undefined
  }

  return `${(value / 100) * zoomFactor}pt`
}

function effectiveFontSize(style?: TextStyle | null): number | undefined {
  if (!style?.fontSize || style.fontSize <= 0) {
    return undefined
  }

  const relativeScale =
    style.relativeSize && style.relativeSize > 0 ? style.relativeSize / 100 : 1
  return style.fontSize * relativeScale
}

function baselineOffsetToCss(style?: TextStyle | null): string | undefined {
  if (!style?.baselineOffset || style.baselineOffset === 0) {
    return undefined
  }

  return `${style.baselineOffset / 1000}em`
}

function letterSpacingToCss(style?: TextStyle | null): string | undefined {
  if (!style?.letterSpacing || style.letterSpacing === 0) {
    return undefined
  }

  return `${style.letterSpacing / 1000}em`
}

function lineHeightToCss(
  style?: ParagraphStyle | null,
  zoomFactor = 1,
): string | number | undefined {
  if (!style?.lineSpacing || style.lineSpacing <= 0) {
    return undefined
  }

  const spacingType = style.lineSpacingType ?? 'PERCENT'

  if (spacingType === 'PERCENT' || spacingType === '0') {
    // Hangul default: 160 = 160% of font size → CSS line-height: 1.6
    return style.lineSpacing / 100
  }

  if (spacingType === 'FIXED' || spacingType === '1') {
    const pixels = hwpToPx(style.lineSpacing)
    if (pixels) {
      return `${Math.round(pixels * zoomFactor)}px`
    }
  }

  if (spacingType === 'BETWEEN_LINES' || spacingType === '2') {
    // Space between lines (added to font height)
    const pixels = hwpToPx(style.lineSpacing)
    if (pixels) {
      return `calc(1em + ${Math.round(pixels * zoomFactor)}px)`
    }
  }

  // Fallback: treat as percent if numeric
  if (style.lineSpacing >= 50 && style.lineSpacing <= 500) {
    return style.lineSpacing / 100
  }

  const pixels = hwpToPx(style.lineSpacing)
  if (pixels) {
    return `${Math.round(pixels * zoomFactor)}px`
  }

  return undefined
}

function mapParagraphAlign(align?: string | null): CSSProperties['textAlign'] {
  switch (align) {
    case 'LEFT':
      return 'left'
    case 'RIGHT':
      return 'right'
    case 'CENTER':
      return 'center'
    case 'JUSTIFY':
    case 'DISTRIBUTE':
    case 'DISTRIBUTE_SPACE':
      return 'justify'
    default:
      return undefined
  }
}

function basePageWidth(pageLayout?: PageLayout | null): number {
  // In HWPX, width/height already represent the actual visual page dimensions.
  // The landscape flag indicates paper orientation metadata but does NOT mean
  // the dimensions are stored swapped — do NOT swap based on landscape.
  return hwpToPx(pageLayout?.width) ?? DEFAULT_PAGE_WIDTH
}

function basePageHeight(pageLayout?: PageLayout | null): number {
  // In HWPX, width/height already represent the actual visual page dimensions.
  // The landscape flag indicates paper orientation metadata but does NOT mean
  // the dimensions are stored swapped — do NOT swap based on landscape.
  return hwpToPx(pageLayout?.height) ?? DEFAULT_PAGE_HEIGHT
}

function resolveZoomFactor(
  pageLayout: PageLayout | null | undefined,
  zoomPreset: ZoomPreset,
  workspaceWidth: number | null,
): number {
  if (zoomPreset !== 'fitWidth') {
    return Number(zoomPreset) / 100
  }

  if (!workspaceWidth) {
    return 1
  }

  const availableWidth = Math.max(320, workspaceWidth - 48)
  return Math.min(1.5, Math.max(0.58, availableWidth / basePageWidth(pageLayout)))
}

function pageStyle(pageLayout?: PageLayout | null, zoomFactor = 1): CSSProperties {
  const width = Math.round(basePageWidth(pageLayout) * zoomFactor)
  const height = Math.round(basePageHeight(pageLayout) * zoomFactor)
  const paddingTop = Math.round((hwpToPx(pageLayout?.marginTop) ?? 76) * zoomFactor)
  const paddingRight = Math.round((hwpToPx(pageLayout?.marginRight) ?? 113) * zoomFactor)
  const paddingBottom = Math.round((hwpToPx(pageLayout?.marginBottom) ?? 57) * zoomFactor)
  const paddingLeft = Math.round((hwpToPx(pageLayout?.marginLeft) ?? 113) * zoomFactor)

  const style: CSSProperties = {
    width: `${width}px`,
    height: `${height}px`,
    paddingTop: `${paddingTop}px`,
    paddingRight: `${paddingRight}px`,
    paddingBottom: `${paddingBottom}px`,
    paddingLeft: `${paddingLeft}px`,
    flexShrink: 0,
  }

  if (pageLayout?.pageBorder) {
    applyTableCellStyle(style, pageLayout.pageBorder)
  }

  return style
}

function pageMeasureStyle(pageLayout?: PageLayout | null, zoomFactor = 1): CSSProperties {
  return {
    ...pageStyle(pageLayout, zoomFactor),
    height: 'auto',
    minHeight: undefined,
  }
}

function pageInnerHeight(pageLayout?: PageLayout | null, zoomFactor = 1): number {
  const height = basePageHeight(pageLayout) * zoomFactor
  const paddingTop = (hwpToPx(pageLayout?.marginTop) ?? 76) * zoomFactor
  const paddingBottom = (hwpToPx(pageLayout?.marginBottom) ?? 57) * zoomFactor
  return Math.max(120, Math.round(height - paddingTop - paddingBottom))
}

function pageContentWidth(pageLayout?: PageLayout | null, zoomFactor = 1): number {
  const width = basePageWidth(pageLayout) * zoomFactor
  const paddingLeft = (hwpToPx(pageLayout?.marginLeft) ?? 113) * zoomFactor
  const paddingRight = (hwpToPx(pageLayout?.marginRight) ?? 113) * zoomFactor
  return Math.max(120, Math.round(width - paddingLeft - paddingRight))
}

function paragraphStyle(style?: ParagraphStyle | null, zoomFactor = 1): CSSProperties {
  return {
    textAlign: mapParagraphAlign(style?.align),
    marginLeft: style?.marginLeft
      ? `${Math.round((hwpToPx(style.marginLeft) ?? 0) * zoomFactor)}px`
      : undefined,
    marginRight: style?.marginRight
      ? `${Math.round((hwpToPx(style.marginRight) ?? 0) * zoomFactor)}px`
      : undefined,
    marginTop: style?.marginPrev
      ? `${Math.round((hwpToPx(style.marginPrev) ?? 0) * zoomFactor)}px`
      : undefined,
    marginBottom: style?.marginNext
      ? `${Math.round((hwpToPx(style.marginNext) ?? 0) * zoomFactor)}px`
      : undefined,
    textIndent: style?.indent ? `${Math.round((hwpToPx(style.indent) ?? 0) * zoomFactor)}px` : undefined,
    lineHeight: lineHeightToCss(style, zoomFactor),
  }
}

function markerGap(style?: ParagraphStyle | null, zoomFactor = 1): number | undefined {
  if (!style?.markerTextOffset || style.markerTextOffset <= 0) {
    return undefined
  }

  if (style.markerTextOffsetType === 'PERCENT') {
    const width = style.markerWidthAdjust ? hwpToPx(style.markerWidthAdjust) ?? 0 : 0
    return Math.round(width * (style.markerTextOffset / 100) * zoomFactor)
  }

  return Math.round((hwpToPx(style.markerTextOffset) ?? 0) * zoomFactor)
}

function paragraphMarkerStyle(style?: ParagraphStyle | null, zoomFactor = 1): CSSProperties {
  return {
    display: 'inline-block',
    textAlign: mapParagraphAlign(style?.markerAlign),
    minWidth:
      style?.markerWidthAdjust && style.markerWidthAdjust > 0
        ? `${Math.round((hwpToPx(style.markerWidthAdjust) ?? 0) * zoomFactor)}px`
        : undefined,
    marginRight: markerGap(style, zoomFactor) ? `${markerGap(style, zoomFactor)}px` : undefined,
  }
}

function textStyle(style?: TextStyle | null, zoomFactor = 1): CSSProperties {
  const fontSize = effectiveFontSize(style)
  return {
    fontFamily: fontFamilyToCss(style?.fontFamily),
    fontSize: fontSizeToCss(fontSize, zoomFactor),
    color: style?.textColor ?? undefined,
    backgroundColor:
      style?.backgroundColor && style.backgroundColor !== '#FFFFFF'
        ? style.backgroundColor
        : undefined,
    fontStretch:
      style?.widthRatio && style.widthRatio > 0 ? `${style.widthRatio}%` : undefined,
    letterSpacing: letterSpacingToCss(style),
    verticalAlign: baselineOffsetToCss(style),
    fontKerning: style?.useKerning ? 'normal' : 'none',
    fontWeight: style?.bold ? 700 : undefined,
    fontStyle: style?.italic ? 'italic' : undefined,
    textDecoration: style?.underline ? 'underline' : undefined,
  }
}

function formatRoman(value: number): string {
  const numerals: Array<[number, string]> = [
    [1000, 'M'],
    [900, 'CM'],
    [500, 'D'],
    [400, 'CD'],
    [100, 'C'],
    [90, 'XC'],
    [50, 'L'],
    [40, 'XL'],
    [10, 'X'],
    [9, 'IX'],
    [5, 'V'],
    [4, 'IV'],
    [1, 'I'],
  ]

  let number = Math.max(1, value)
  let output = ''
  for (const [size, symbol] of numerals) {
    while (number >= size) {
      number -= size
      output += symbol
    }
  }
  return output
}

function formatLatin(value: number, upper: boolean): string {
  let number = Math.max(1, value)
  let output = ''
  while (number > 0) {
    number -= 1
    const code = (upper ? 65 : 97) + (number % 26)
    output = String.fromCharCode(code) + output
    number = Math.floor(number / 26)
  }
  return output
}

function formatPageNumber(value: number, format = 'DIGIT'): string {
  switch (format) {
    case 'ROMAN_CAPITAL':
      return formatRoman(value)
    case 'ROMAN_SMALL':
      return formatRoman(value).toLowerCase()
    case 'LATIN_CAPITAL':
      return formatLatin(value, true)
    case 'LATIN_SMALL':
      return formatLatin(value, false)
    case 'CIRCLED_DIGIT': {
      const circled = ['①', '②', '③', '④', '⑤', '⑥', '⑦', '⑧', '⑨', '⑩']
      return circled[value - 1] ?? String(value)
    }
    default:
      return String(value)
  }
}

function replacePlaceholders(text: string, context: PlaceholderContext): string {
  return text.replace(/\{\{(PAGE|TOTAL_PAGES):([^:}]*):([^}]*)\}\}/g, (_, kind, format, sideChar) => {
    const value = kind === 'TOTAL_PAGES' ? context.totalPages : context.pageNumber
    const formatted = formatPageNumber(value, format || 'DIGIT')
    return sideChar ? `${sideChar}${formatted}${sideChar}` : formatted
  })
}

function mapBorderStyle(style?: string | null): CSSProperties['borderLeftStyle'] {
  switch (style) {
    case 'NONE':
      return 'none'
    case 'SOLID':
      return 'solid'
    case 'DOUBLE':
    case 'DOUBLE_SLIM':
    case 'SLIM_THICK':
    case 'THICK_SLIM':
      return 'double'
    case 'DOT':
      return 'dotted'
    case 'DASH':
    case 'LONG_DASH':
    case 'DASH_DOT':
    case 'DASH_DOT_DOT':
      return 'dashed'
    default:
      return undefined
  }
}

function normalizeBorderWidth(width?: string | null): string | undefined {
  if (!width) {
    return undefined
  }

  const stripped = width.replace(/\s+/g, '')

  // Convert mm values to integer px so sub-pixel borders (e.g. 0.12mm ≈ 0.5px)
  // are always rendered as at least 1 visible pixel.
  const mmMatch = stripped.match(/^([\d.]+)mm$/)
  if (mmMatch) {
    const px = parseFloat(mmMatch[1]) * (96 / 25.4)
    return `${Math.max(1, Math.round(px))}px`
  }

  return stripped
}

function quoteFontFamily(name: string): string {
  return /[\s",]/.test(name) ? `"${name.replace(/"/g, '\\"')}"` : name
}

function isSerifFontFamily(name: string): boolean {
  return /명조|바탕|myungjo|myeongjo|batang/i.test(name)
}

function isSansFontFamily(name: string): boolean {
  return /고딕|돋움|굴림|헤드라인|산세리프|gothic|dotum|gulim|headline/i.test(name)
}

function fontFamilyToCss(fontFamily?: string | null): string | undefined {
  if (!fontFamily) {
    return undefined
  }

  const trimmed = fontFamily.trim()
  if (!trimmed) {
    return undefined
  }

  const families = new Set<string>([trimmed])

  if (/맑은 고딕/i.test(trimmed)) {
    families.add('Malgun Gothic')
  }
  if (/휴먼명조/i.test(trimmed)) {
    families.add('AppleMyungjo')
  }
  if (/한컴바탕|함초롬바탕|바탕체|바탕/i.test(trimmed)) {
    families.add('Batang')
    families.add('AppleMyungjo')
    families.add('Nanum Myeongjo')
  }
  if (/한컴고딕|함초롬돋움|견고딕|중고딕|고딕|돋움|굴림|헤드라인/i.test(trimmed)) {
    families.add('Dotum')
    families.add('Gulim')
    families.add('Apple SD Gothic Neo')
    families.add('NanumSquare')
  }
  if (/한양신명조|신명조|명조/i.test(trimmed)) {
    families.add('AppleMyungjo')
    families.add('Nanum Myeongjo')
  }
  if (/궁서|필기체/i.test(trimmed)) {
    families.add('Gungsuh')
    families.add('AppleMyungjo')
  }

  if (isSerifFontFamily(trimmed)) {
    families.add('AppleMyungjo')
    families.add('Nanum Myeongjo')
    families.add('serif')
  } else if (isSansFontFamily(trimmed)) {
    families.add('Malgun Gothic')
    families.add('Apple SD Gothic Neo')
    families.add('Noto Sans KR')
    families.add('NanumSquare')
    families.add('sans-serif')
  } else {
    families.add('Apple SD Gothic Neo')
    families.add('Noto Sans KR')
    families.add('sans-serif')
  }

  return Array.from(families).map(quoteFontFamily).join(', ')
}

function applyBorderSide(
  output: CSSProperties,
  side: 'Left' | 'Right' | 'Top' | 'Bottom',
  border?: TableBorder | null,
) {
  if (!border) {
    return
  }

  const borderStyle = mapBorderStyle(border.style)
  const borderWidth = normalizeBorderWidth(border.width)
  const borderColor = border.color ?? undefined

  switch (side) {
    case 'Left':
      output.borderLeftStyle = borderStyle
      output.borderLeftWidth = borderWidth
      output.borderLeftColor = borderColor
      break
    case 'Right':
      output.borderRightStyle = borderStyle
      output.borderRightWidth = borderWidth
      output.borderRightColor = borderColor
      break
    case 'Top':
      output.borderTopStyle = borderStyle
      output.borderTopWidth = borderWidth
      output.borderTopColor = borderColor
      break
    case 'Bottom':
      output.borderBottomStyle = borderStyle
      output.borderBottomWidth = borderWidth
      output.borderBottomColor = borderColor
      break
  }
}

function applyTableCellStyle(output: CSSProperties, style?: TableCellStyle | null) {
  if (!style) {
    return
  }

  if (style.backgroundColor) {
    output.backgroundColor = style.backgroundColor
  }

  applyBorderSide(output, 'Left', style.borderLeft)
  applyBorderSide(output, 'Right', style.borderRight)
  applyBorderSide(output, 'Top', style.borderTop)
  applyBorderSide(output, 'Bottom', style.borderBottom)
}

function hasVisibleBorder(border?: TableBorder | null): boolean {
  return Boolean(border?.style && border.style !== 'NONE')
}

function hasVisibleDiagonal(diagonal?: TableDiagonal | null): boolean {
  if (!diagonal) {
    return false
  }

  const hasDirection =
    (diagonal.slashType && diagonal.slashType !== 'NONE')
    || (diagonal.backSlashType && diagonal.backSlashType !== 'NONE')

  return Boolean(hasDirection && diagonal.style !== 'NONE')
}

function shouldUseSeparateBorderModel(table: TableValue): boolean {
  if (table.rows.length !== 1 || table.rows[0]?.cells.length !== 1) {
    return false
  }

  const cell = table.rows[0].cells[0]
  const tableHasBorder = [
    table.style?.borderLeft,
    table.style?.borderRight,
    table.style?.borderTop,
    table.style?.borderBottom,
  ].some(hasVisibleBorder)
  const cellHasBorder = [
    cell.style?.borderLeft,
    cell.style?.borderRight,
    cell.style?.borderTop,
    cell.style?.borderBottom,
  ].some(hasVisibleBorder)

  return cellHasBorder && !tableHasBorder
}

function shouldFitTableToContainer(table: TableValue): boolean {
  return !table.noAdjust && table.rows.length === 1 && (table.rows[0]?.cells.length ?? 0) === 1
}

function tableStyle(table: TableValue, zoomFactor: number): CSSProperties {
  const resolvedWidth = table.width
    ? `${Math.round((hwpToPx(table.width) ?? 0) * zoomFactor)}px`
    : undefined
  const style: CSSProperties = {
    width: shouldFitTableToContainer(table)
      ? '100%'
      : resolvedWidth
        ? table.noAdjust ? resolvedWidth : `min(${resolvedWidth}, 100%)`
      : '100%',
    maxWidth: '100%',
    boxSizing: 'border-box',
    tableLayout: table.width ? 'fixed' : 'auto',
  }

  if (shouldUseSeparateBorderModel(table)) {
    style.borderCollapse = 'separate'
    style.borderSpacing = 0
  }

  applyTableCellStyle(style, table.style)
  return style
}

function tableCellStyle(
  cell: TableCell,
  zoomFactor: number,
  options?: { ignoreWidth?: boolean },
): CSSProperties {
  const style: CSSProperties = {
    width:
      !options?.ignoreWidth && cell.width
        ? `${Math.round((hwpToPx(cell.width) ?? 0) * zoomFactor)}px`
        : undefined,
    height: cell.height ? `${Math.round((hwpToPx(cell.height) ?? 0) * zoomFactor)}px` : undefined,
    boxSizing: 'border-box',
    paddingLeft: cell.paddingLeft
      ? `${Math.round((hwpToPx(cell.paddingLeft) ?? 0) * zoomFactor)}px`
      : undefined,
    paddingRight: cell.paddingRight
      ? `${Math.round((hwpToPx(cell.paddingRight) ?? 0) * zoomFactor)}px`
      : undefined,
    paddingTop: cell.paddingTop
      ? `${Math.round((hwpToPx(cell.paddingTop) ?? 0) * zoomFactor)}px`
      : undefined,
    paddingBottom: cell.paddingBottom
      ? `${Math.round((hwpToPx(cell.paddingBottom) ?? 0) * zoomFactor)}px`
      : undefined,
  }

  applyTableCellStyle(style, cell.style)
  return style
}

function diagonalStrokeDasharray(style?: string | null): string | undefined {
  switch (style) {
    case 'DOT':
      return '1 3'
    case 'DASH':
      return '6 4'
    case 'LONG_DASH':
      return '10 4'
    case 'DASH_DOT':
      return '8 4 1 4'
    case 'DASH_DOT_DOT':
      return '8 4 1 4 1 4'
    default:
      return undefined
  }
}

function renderTableCellDiagonal(diagonal: TableDiagonal, keyPrefix: string) {
  if (!hasVisibleDiagonal(diagonal)) {
    return null
  }

  const strokeWidth = normalizeBorderWidth(diagonal.width) ?? '1px'
  const stroke = diagonal.color ?? '#000000'
  const strokeDasharray = diagonalStrokeDasharray(diagonal.style)

  return (
    <svg
      key={`${keyPrefix}-diagonal`}
      className="table-cell-diagonal"
      viewBox="-1 -1 102 102"
      preserveAspectRatio="none"
      aria-hidden="true"
    >
      {diagonal.backSlashType && diagonal.backSlashType !== 'NONE' ? (
        <line
          x1="0"
          y1="0"
          x2="100"
          y2="100"
          stroke={stroke}
          strokeWidth={strokeWidth}
          strokeDasharray={strokeDasharray}
          vectorEffect="non-scaling-stroke"
          strokeLinecap="square"
        />
      ) : null}
      {diagonal.slashType && diagonal.slashType !== 'NONE' ? (
        <line
          x1="100"
          y1="0"
          x2="0"
          y2="100"
          stroke={stroke}
          strokeWidth={strokeWidth}
          strokeDasharray={strokeDasharray}
          vectorEffect="non-scaling-stroke"
          strokeLinecap="square"
        />
      ) : null}
    </svg>
  )
}

function resolveAsset(assets: AssetRef[], assetId: string): AssetRef | undefined {
  return (
    assets.find((asset) => asset.id === assetId) ??
    assets.find((asset) => asset.sourcePath === assetId) ??
    assets.find((asset) => asset.id.split('/').pop() === assetId) ??
    assets.find((asset) => asset.id.split('/').pop()?.replace(/\.[^.]+$/, '') === assetId)
  )
}

function imageDistancePx(image: ImageValue, side: 'left' | 'right' | 'top' | 'bottom', zoomFactor: number): number {
  const raw =
    side === 'left' ? image.distanceLeft
    : side === 'right' ? image.distanceRight
    : side === 'top' ? image.distanceTop
    : image.distanceBottom
  if (raw != null) {
    return Math.max(0, Math.round((hwpToPx(raw) ?? 0) * zoomFactor))
  }
  return Math.round((side === 'top' || side === 'bottom' ? 8 : 10) * zoomFactor)
}

function resolveTextWrapMode(image: ImageValue): string {
  if (image.treatAsChar) return 'INLINE'
  const wrap = image.textWrap ?? ''
  if (wrap === 'BEHIND_TEXT' || wrap === 'IN_FRONT_OF_TEXT') return wrap
  if (wrap === 'TOP_AND_BOTTOM') return 'TOP_AND_BOTTOM'
  if (wrap === 'SQUARE' || wrap === 'TIGHT') return wrap
  if (image.horzAlign === 'LEFT' || image.horzAlign === 'RIGHT') return 'SQUARE'
  return 'TOP_AND_BOTTOM'
}

function objectPlacementStyle(image: ImageValue, zoomFactor: number): CSSProperties {
  const width = image.width ? Math.round((hwpToPx(image.width) ?? 0) * zoomFactor) : undefined
  const height = image.height ? Math.round((hwpToPx(image.height) ?? 0) * zoomFactor) : undefined

  if (image.treatAsChar) {
    return {
      width: width ? `${width}px` : undefined,
      height: height ? `${height}px` : undefined,
    }
  }

  const wrapMode = resolveTextWrapMode(image)

  const style: CSSProperties = {
    position: 'absolute',
    width: width ? `${width}px` : undefined,
    height: height ? `${height}px` : undefined,
  }

  if (wrapMode === 'BEHIND_TEXT') {
    style.zIndex = -1
    style.pointerEvents = 'none'
  } else if (wrapMode === 'IN_FRONT_OF_TEXT') {
    style.zIndex = Math.max(10, image.zOrder ?? 10)
    style.pointerEvents = 'none'
  } else {
    style.zIndex = image.zOrder ?? 1
  }

  const horizontalOffset = image.horzOffset ? Math.round((hwpToPx(image.horzOffset) ?? 0) * zoomFactor) : 0
  const verticalOffset = image.vertOffset ? Math.round((hwpToPx(image.vertOffset) ?? 0) * zoomFactor) : 0

  if (image.horzAlign === 'RIGHT') {
    style.right = `${horizontalOffset}px`
  } else if (image.horzAlign === 'CENTER') {
    style.left = `calc(50% + ${horizontalOffset}px)`
    style.transform = 'translateX(-50%)'
  } else {
    style.left = `${horizontalOffset}px`
  }

  if (image.vertAlign === 'BOTTOM') {
    style.bottom = `${verticalOffset}px`
  } else if (image.vertAlign === 'CENTER') {
    style.top = `calc(50% + ${verticalOffset}px)`
    style.transform = style.transform
      ? `${style.transform} translateY(-50%)`
      : 'translateY(-50%)'
  } else {
    style.top = `${verticalOffset}px`
  }

  // Rotation
  if (image.rotation && image.rotation !== 0) {
    const degrees = image.rotation / 10 // HWP unit: 1/10 degree
    const rotateStr = `rotate(${degrees}deg)`
    style.transform = style.transform ? `${style.transform} ${rotateStr}` : rotateStr
  }

  return style
}

function floatingObjectFootprintHeight(image: ImageValue, zoomFactor: number): number {
  const wrapMode = resolveTextWrapMode(image)
  if (wrapMode === 'BEHIND_TEXT' || wrapMode === 'IN_FRONT_OF_TEXT') {
    return 0
  }
  const imageHeight = image.height ? Math.round((hwpToPx(image.height) ?? 0) * zoomFactor) : 48
  const verticalOffset = image.vertOffset
    ? Math.max(0, Math.round((hwpToPx(image.vertOffset) ?? 0) * zoomFactor))
    : 0
  const captionAllowance = image.caption || image.altText ? Math.round(32 * zoomFactor) : 0
  const distTop = imageDistancePx(image, 'top', zoomFactor)
  const distBottom = imageDistancePx(image, 'bottom', zoomFactor)
  if (wrapMode === 'TOP_AND_BOTTOM') {
    return verticalOffset + distTop + imageHeight + distBottom + captionAllowance
  }
  return Math.max(imageHeight, verticalOffset + imageHeight + captionAllowance)
}

function floatingObjectFootprintWidth(image: ImageValue, zoomFactor: number): number {
  const imageWidth = image.width ? Math.round((hwpToPx(image.width) ?? 0) * zoomFactor) : 120
  const horizontalOffset = image.horzOffset
    ? Math.max(0, Math.round((hwpToPx(image.horzOffset) ?? 0) * zoomFactor))
    : 0
  const distLeft = imageDistancePx(image, 'left', zoomFactor)
  const distRight = imageDistancePx(image, 'right', zoomFactor)
  return imageWidth + horizontalOffset + distLeft + distRight
}

function supportsSideTextWrap(image: ImageValue): boolean {
  const wrapMode = resolveTextWrapMode(image)
  return wrapMode === 'SQUARE' || wrapMode === 'TIGHT'
}

function floatingObjectReserveStyle(image: ImageValue, zoomFactor: number): CSSProperties {
  if (image.treatAsChar) {
    return {}
  }

  const wrapMode = resolveTextWrapMode(image)

  if (wrapMode === 'BEHIND_TEXT' || wrapMode === 'IN_FRONT_OF_TEXT') {
    return { position: 'relative', minHeight: '0px' }
  }

  const style: CSSProperties = {
    minHeight: `${floatingObjectFootprintHeight(image, zoomFactor)}px`,
  }

  if (supportsSideTextWrap(image)) {
    style.width = `${floatingObjectFootprintWidth(image, zoomFactor)}px`
    style.float = image.horzAlign === 'RIGHT' ? 'right' : 'left'
    style.marginBottom = `${imageDistancePx(image, 'bottom', zoomFactor)}px`
    if (image.horzAlign === 'RIGHT') {
      style.marginLeft = `${imageDistancePx(image, 'left', zoomFactor)}px`
    } else {
      style.marginRight = `${imageDistancePx(image, 'right', zoomFactor)}px`
    }
    style.shapeOutside = 'margin-box'
  } else if (wrapMode === 'TOP_AND_BOTTOM') {
    style.clear = 'both'
    style.marginTop = `${imageDistancePx(image, 'top', zoomFactor)}px`
    style.marginBottom = `${imageDistancePx(image, 'bottom', zoomFactor)}px`
  }

  return style
}

type ExclusionZone = {
  left: number
  top: number
  right: number
  bottom: number
  wrapMode: string
  side: 'left' | 'right' | 'center'
}

function calculateExclusionZone(
  image: ImageValue,
  contentWidth: number,
  zoomFactor: number,
): ExclusionZone | null {
  if (image.treatAsChar) return null
  const wrapMode = resolveTextWrapMode(image)
  if (wrapMode === 'BEHIND_TEXT' || wrapMode === 'IN_FRONT_OF_TEXT' || wrapMode === 'INLINE') {
    return null
  }

  const objWidth = image.width ? Math.round((hwpToPx(image.width) ?? 0) * zoomFactor) : 120
  const objHeight = image.height ? Math.round((hwpToPx(image.height) ?? 0) * zoomFactor) : 48
  let horzOffset = image.horzOffset ? Math.round((hwpToPx(image.horzOffset) ?? 0) * zoomFactor) : 0
  const vertOffset = image.vertOffset ? Math.max(0, Math.round((hwpToPx(image.vertOffset) ?? 0) * zoomFactor)) : 0

  // Coordinate system correction: if horzRelTo is PAGE, offset is from page edge
  // but our contentWidth is the body area, so subtract left margin conceptually
  if (image.horzRelTo === 'PAGE' && horzOffset > 0) {
    // Approximate: page-relative offset likely already accounts for margins in most documents
    // No adjustment needed for typical cases, but clamp to content bounds
    horzOffset = Math.min(horzOffset, contentWidth - objWidth)
  }
  const dLeft = imageDistancePx(image, 'left', zoomFactor)
  const dRight = imageDistancePx(image, 'right', zoomFactor)
  const dTop = imageDistancePx(image, 'top', zoomFactor)
  const dBottom = imageDistancePx(image, 'bottom', zoomFactor)

  let left: number
  let side: 'left' | 'right' | 'center'
  if (image.horzAlign === 'RIGHT') {
    left = contentWidth - objWidth - horzOffset
    side = 'right'
  } else if (image.horzAlign === 'CENTER') {
    left = (contentWidth - objWidth) / 2 + horzOffset
    side = 'center'
  } else {
    left = horzOffset
    side = 'left'
  }

  return {
    left: left - dLeft,
    top: vertOffset - dTop,
    right: left + objWidth + dRight,
    bottom: vertOffset + objHeight + dBottom,
    wrapMode,
    side,
  }
}

function collectExclusionZones(
  blocks: DocumentBlock[],
  contentWidth: number,
  zoomFactor: number,
): ExclusionZone[] {
  const zones: ExclusionZone[] = []
  for (const block of blocks) {
    if (block.type === 'image') {
      const zone = calculateExclusionZone(block.value, contentWidth, zoomFactor)
      if (zone) zones.push(zone)
    }
  }
  return zones
}

function calculateParagraphInsets(
  blockTop: number,
  blockBottom: number,
  exclusions: ExclusionZone[],
  contentWidth: number,
): { marginLeft?: number; marginRight?: number } | null {
  let insetLeft = 0
  let insetRight = 0
  let hasInset = false

  for (const zone of exclusions) {
    if (zone.wrapMode === 'TOP_AND_BOTTOM') continue
    if (blockBottom <= zone.top || blockTop >= zone.bottom) continue

    const zoneWidth = zone.right - zone.left
    if (zone.side === 'left') {
      insetLeft = Math.max(insetLeft, zone.right)
      hasInset = true
    } else if (zone.side === 'right') {
      insetRight = Math.max(insetRight, contentWidth - zone.left)
      hasInset = true
    } else {
      if (insetLeft + insetRight + zoneWidth >= contentWidth * 0.8) continue
      insetLeft = Math.max(insetLeft, zone.right)
      hasInset = true
    }
  }

  if (!hasInset) return null
  if (insetLeft + insetRight >= contentWidth * 0.8) return null

  return {
    marginLeft: insetLeft > 0 ? insetLeft : undefined,
    marginRight: insetRight > 0 ? insetRight : undefined,
  }
}

function blockHeightHint(block: DocumentBlock, zoomFactor: number): number | null {
  switch (block.type) {
    case 'paragraph': {
      const hintedHeight = hwpToPx(block.value.layoutHeightHint)
      if (hintedHeight) {
        return Math.round(hintedHeight * zoomFactor)
      }

      if (block.value.lineSegmentCount && block.value.lineSegmentCount > 0) {
        const spacingType = block.value.style?.lineSpacingType ?? 'PERCENT'
        const spacing = block.value.style?.lineSpacing
        const fontSize = block.value.runs[0]?.style?.fontSize
        const baseFontPx = fontSize ? fontSize / 100 * (96 / 72) : 10 * (96 / 72) // pt to px

        let singleLineHeight: number | null = null
        if ((spacingType === 'PERCENT' || spacingType === '0') && spacing) {
          singleLineHeight = baseFontPx * (spacing / 100)
        } else if (spacing) {
          const lineHeightPx = hwpToPx(spacing)
          if (lineHeightPx) singleLineHeight = lineHeightPx
        }

        if (singleLineHeight) {
          return Math.round(singleLineHeight * block.value.lineSegmentCount * zoomFactor)
        }
      }

      return null
    }
    case 'table':
      return block.value.height ? Math.round((hwpToPx(block.value.height) ?? 0) * zoomFactor) : null
    case 'image':
      return block.value.treatAsChar
        ? block.value.height
          ? Math.round((hwpToPx(block.value.height) ?? 0) * zoomFactor)
          : null
        : floatingObjectFootprintHeight(block.value, zoomFactor)
    case 'footnote':
    case 'unsupported':
      return null
  }
}

function blockPageBreakBefore(block: DocumentBlock): boolean {
  return block.type === 'paragraph' ? block.value.pageBreakBefore : false
}

function blockKeepWithNext(block: DocumentBlock): boolean {
  return block.type === 'paragraph' ? block.value.style?.keepWithNext ?? false : false
}

function blockKeepLines(block: DocumentBlock): boolean {
  return block.type === 'paragraph' ? block.value.style?.keepLines ?? false : false
}

function paginationHeight(actualHeight: number, hintHeight: number | null): number {
  if (hintHeight == null || hintHeight <= 0) {
    return actualHeight
  }

  // Trust hint height more heavily — it comes from the document's own layout metrics
  // Blend: 70% hint, 30% actual to absorb font/rendering differences
  return Math.round(hintHeight * 0.7 + actualHeight * 0.3)
}

function overflowsRemaining(
  actualHeight: number,
  hintHeight: number | null,
  remainingHeight: number,
): boolean {
  // If hint says it fits, trust it — the document's own layout metrics are authoritative
  if (hintHeight != null && hintHeight > 0 && hintHeight <= remainingHeight + PAGINATION_TOLERANCE) {
    return false
  }

  if (actualHeight <= remainingHeight + PAGINATION_TOLERANCE) {
    return false
  }

  return paginationHeight(actualHeight, hintHeight) > remainingHeight + PAGINATION_TOLERANCE
}

function fitsFreshPage(item: PaginationItem, availableHeight: number): boolean {
  return !overflowsRemaining(item.actualHeight, item.hintHeight, availableHeight)
}

function isSingleCellContainerTable(
  block: DocumentBlock,
): block is Extract<DocumentBlock, { type: 'table' }> {
  return (
    block.type === 'table' &&
    block.value.rows.length === 1 &&
    block.value.rows[0]?.cells.length === 1 &&
    (block.value.rows[0]?.cells[0]?.blocks.length ?? 0) > 0
  )
}

function cloneTableFragment(
  table: TableValue,
  cellBlocks: DocumentBlock[],
): Extract<DocumentBlock, { type: 'table' }> {
  const sourceRow = table.rows[0]
  const sourceCell = sourceRow.cells[0]

  return {
    type: 'table',
    value: {
      ...table,
      height: null,
      rows: [
        {
          ...sourceRow,
          cells: [
            {
              ...sourceCell,
              height: null,
              blocks: cellBlocks,
            },
          ],
        },
      ],
    },
  }
}

function cloneTableRowsFragment(
  table: TableValue,
  rowStart: number,
  rowEndExclusive: number,
  headerRowCount: number,
): Extract<DocumentBlock, { type: 'table' }> {
  const repeatedHeaderRows =
    rowStart > 0 && headerRowCount > 0
      ? table.rows.slice(0, headerRowCount)
      : []
  const dataRows = table.rows.slice(rowStart, rowEndExclusive)

  return {
    type: 'table',
    value: {
      ...table,
      height: null,
      rows: [...repeatedHeaderRows, ...dataRows].map((row) => ({
        ...row,
        cells: row.cells.map((cell) => ({
          ...cell,
          blocks: [...cell.blocks],
        })),
      })),
    },
  }
}

function paragraphRunsToLines(paragraph: ParagraphValue): TextRun[][] {
  if (paragraph.runs.length === 0) {
    return [[]]
  }

  const lines: TextRun[][] = [[]]

  for (const run of paragraph.runs) {
    const parts = run.text.split('\n')
    parts.forEach((part, partIndex) => {
      if (part.length > 0) {
        lines[lines.length - 1].push({
          ...run,
          text: part,
        })
      }

      if (partIndex < parts.length - 1) {
        lines.push([])
      }
    })
  }

  return lines.length > 0 ? lines : [[]]
}

function paragraphLinesToRuns(lines: TextRun[][]): TextRun[] {
  const runs: TextRun[] = []

  lines.forEach((line, lineIndex) => {
    if (lineIndex > 0) {
      runs.push({
        text: '\n',
        style: null,
      })
    }

    line.forEach((run) => {
      runs.push({
        ...run,
      })
    })
  })

  return runs
}

function cloneParagraphFragment(
  paragraph: ParagraphValue,
  lineRuns: TextRun[][],
  fragmentIndex: number,
  fragmentCount: number,
): Extract<DocumentBlock, { type: 'paragraph' }> {
  const style = paragraph.style
    ? {
        ...paragraph.style,
        marginPrev: fragmentIndex === 0 ? paragraph.style.marginPrev : 0,
        marginNext: fragmentIndex === fragmentCount - 1 ? paragraph.style.marginNext : 0,
        keepWithNext:
          fragmentIndex === fragmentCount - 1 ? paragraph.style.keepWithNext : false,
      }
    : paragraph.style

  return {
    type: 'paragraph',
    value: {
      ...paragraph,
      marker: fragmentIndex === 0 ? paragraph.marker : null,
      runs: paragraphLinesToRuns(lineRuns),
      style,
      lineSegmentCount: lineRuns.length,
      layoutHeightHint: null,
      pageBreakBefore: fragmentIndex === 0 ? paragraph.pageBreakBefore : false,
    },
  }
}

function tableRowHeightHint(row: TableRow, zoomFactor: number): number | null {
  const cellHeights = row.cells
    .map((cell) => (cell.height ? Math.round((hwpToPx(cell.height) ?? 0) * zoomFactor) : 0))
    .filter((height) => height > 0)

  if (cellHeights.length === 0) {
    return null
  }

  return Math.max(...cellHeights)
}

function computeSafeRowBreakStarts(table: TableValue): Set<number> {
  const safeBreakStarts = new Set<number>()
  const activeSpans: number[] = []

  table.rows.forEach((row, rowIndex) => {
    let columnIndex = 0

    row.cells.forEach((cell) => {
      while ((activeSpans[columnIndex] ?? 0) > 0) {
        columnIndex += 1
      }

      const colSpan = Math.max(1, cell.colSpan ?? 1)
      const rowSpan = Math.max(1, cell.rowSpan ?? 1)

      if (rowSpan > 1) {
        for (let spanIndex = 0; spanIndex < colSpan; spanIndex += 1) {
          activeSpans[columnIndex + spanIndex] = rowSpan
        }
      }

      columnIndex += colSpan
    })

    const hasCarryOver = activeSpans.some((span) => span > 1)

    activeSpans.forEach((span, spanIndex) => {
      if (span > 0) {
        activeSpans[spanIndex] = span - 1
      }
    })

    if (!hasCarryOver) {
      safeBreakStarts.add(rowIndex + 1)
    }
  })

  safeBreakStarts.add(table.rows.length)
  return safeBreakStarts
}

type RowSpanGridCell = {
  ownerRow: number
  ownerCol: number
  rowSpan: number
  colSpan: number
}

function buildRowSpanGrid(table: TableValue): RowSpanGridCell[][] {
  const rowCount = table.rows.length
  let maxCols = 0
  for (const row of table.rows) {
    let cols = 0
    for (const cell of row.cells) {
      cols += Math.max(1, cell.colSpan ?? 1)
    }
    if (cols > maxCols) maxCols = cols
  }

  const grid: RowSpanGridCell[][] = Array.from({ length: rowCount }, () =>
    Array.from({ length: maxCols }, () => ({
      ownerRow: -1,
      ownerCol: -1,
      rowSpan: 1,
      colSpan: 1,
    })),
  )

  for (let r = 0; r < rowCount; r += 1) {
    let c = 0
    for (const cell of table.rows[r].cells) {
      while (c < maxCols && grid[r][c].ownerRow !== -1) {
        c += 1
      }
      if (c >= maxCols) break

      const cs = Math.max(1, cell.colSpan ?? 1)
      const rs = Math.max(1, cell.rowSpan ?? 1)

      for (let dr = 0; dr < rs && r + dr < rowCount; dr += 1) {
        for (let dc = 0; dc < cs && c + dc < maxCols; dc += 1) {
          grid[r + dr][c + dc] = { ownerRow: r, ownerCol: c, rowSpan: rs, colSpan: cs }
        }
      }

      c += cs
    }
  }

  return grid
}

function countCrossingSpans(grid: RowSpanGridCell[][], splitAfterRow: number): number {
  if (splitAfterRow < 0 || splitAfterRow >= grid.length) return 0
  const cols = grid[0]?.length ?? 0
  let crossings = 0
  for (let c = 0; c < cols; c += 1) {
    const cell = grid[splitAfterRow][c]
    if (cell.ownerRow === -1) continue
    if (cell.ownerRow + cell.rowSpan - 1 > splitAfterRow) {
      crossings += 1
    }
  }
  return crossings
}

function findBestForcedSplitRow(
  grid: RowSpanGridCell[][],
  headerRowCount: number,
): number | null {
  const rowCount = grid.length
  const startRow = Math.max(1, headerRowCount + 1)
  if (startRow >= rowCount) return null

  let bestRow = -1
  let bestCrossings = Infinity

  for (let r = startRow; r < rowCount; r += 1) {
    const crossings = countCrossingSpans(grid, r - 1)
    if (crossings < bestCrossings) {
      bestCrossings = crossings
      bestRow = r
    }
    if (crossings === 0) break
  }

  return bestRow > 0 && bestRow < rowCount ? bestRow : null
}

function cloneTableRowsFragmentWithSpanFix(
  table: TableValue,
  rowStart: number,
  rowEndExclusive: number,
  headerRowCount: number,
  grid: RowSpanGridCell[][],
): Extract<DocumentBlock, { type: 'table' }> {
  const repeatedHeaderRows =
    rowStart > 0 && headerRowCount > 0
      ? table.rows.slice(0, headerRowCount).map((row) => ({
          ...row,
          cells: row.cells.map((cell) => ({ ...cell, blocks: [...cell.blocks] })),
        }))
      : []

  const dataRows: TableRow[] = []

  for (let r = rowStart; r < rowEndExclusive; r += 1) {
    const originalRow = table.rows[r]
    const fixedCells: TableCell[] = []
    const cols = grid[0]?.length ?? 0
    const visited = new Set<string>()

    for (let c = 0; c < cols; c += 1) {
      const g = grid[r][c]
      if (g.ownerRow === -1) continue
      const key = `${g.ownerRow}-${g.ownerCol}`
      if (visited.has(key)) continue
      visited.add(key)

      if (g.ownerRow >= rowStart) {
        // Cell starts within this fragment
        const originalCell = findCellAt(table, g.ownerRow, g.ownerCol, grid)
        if (originalCell) {
          const clippedRowSpan = Math.min(g.rowSpan, rowEndExclusive - g.ownerRow)
          fixedCells.push({
            ...originalCell,
            rowSpan: clippedRowSpan > 1 ? clippedRowSpan : originalCell.rowSpan,
            blocks: [...originalCell.blocks],
          })
        }
      } else if (r === rowStart) {
        // Cell started before this fragment — create continuation cell
        const originalCell = findCellAt(table, g.ownerRow, g.ownerCol, grid)
        if (originalCell) {
          const remainingSpan = g.ownerRow + g.rowSpan - rowStart
          const clippedSpan = Math.min(remainingSpan, rowEndExclusive - rowStart)
          fixedCells.push({
            ...originalCell,
            rowSpan: clippedSpan > 1 ? clippedSpan : undefined,
            blocks: [...originalCell.blocks],
          })
        }
      }
    }

    if (fixedCells.length > 0) {
      dataRows.push({ cells: fixedCells })
    } else {
      dataRows.push({
        ...originalRow,
        cells: originalRow.cells.map((cell) => ({ ...cell, blocks: [...cell.blocks] })),
      })
    }
  }

  return {
    type: 'table',
    value: {
      ...table,
      height: null,
      rows: [...repeatedHeaderRows, ...dataRows],
    },
  }
}

function findCellAt(
  table: TableValue,
  targetRow: number,
  targetCol: number,
  grid: RowSpanGridCell[][],
): TableCell | null {
  const row = table.rows[targetRow]
  if (!row) return null

  let c = 0
  for (const cell of row.cells) {
    while (c < (grid[0]?.length ?? 0) && grid[targetRow][c]?.ownerRow !== -1 && grid[targetRow][c]?.ownerCol !== c) {
      c += 1
    }
    if (c === targetCol) return cell
    c += Math.max(1, cell.colSpan ?? 1)
  }

  return null
}

function resolveHeaderRowCount(table: TableValue, safeBreakStarts: Set<number>): number {
  if (!table.repeatHeader) {
    return 0
  }

  if (table.headerRowCount != null && table.headerRowCount > 0) {
    return table.headerRowCount
  }

  if (safeBreakStarts.has(1)) {
    return 1
  }

  for (let index = 2; index < table.rows.length; index += 1) {
    if (safeBreakStarts.has(index)) {
      return index
    }
  }

  return 0
}

function buildSingleCellSplitMeasurement(
  block: DocumentBlock,
  blockNode: HTMLElement,
  zoomFactor: number,
): SingleCellTableSplitMeasurement | null {
  if (!isSingleCellContainerTable(block)) {
    return null
  }

  const cellContent = blockNode.querySelector<HTMLElement>('[data-table-cell-content="true"]')
  const childNodes = Array.from(cellContent?.children ?? []).filter(
    (child): child is HTMLElement =>
      child instanceof HTMLElement && child.dataset.cellBlockIndex != null,
  )
  const sourceBlocks = block.value.rows[0].cells[0]?.blocks ?? []
  if (!cellContent || childNodes.length === 0 || childNodes.length !== sourceBlocks.length) {
    return null
  }

  const tableHintHeight = blockHeightHint(block, zoomFactor)
  const childItems = childNodes.map((childNode, index) => {
    const childBlock = sourceBlocks[index]
    return {
      block: childBlock,
      actualHeight: Math.max(1, Math.round(childNode.offsetHeight)),
      hintHeight: blockHeightHint(childBlock, zoomFactor),
    }
  })

  const childHintSum = childItems.reduce((sum, item) => sum + (item.hintHeight ?? item.actualHeight), 0)
  const chromeHeight = Math.max(0, Math.round(blockNode.offsetHeight - cellContent.offsetHeight))
  const chromeHintHeight = Math.max(
    0,
    Math.round((tableHintHeight ?? blockNode.offsetHeight) - childHintSum),
  )

  return {
    kind: 'singleCellTable',
    chromeHeight,
    chromeHintHeight,
    childItems,
  }
}

function buildTableRowSplitMeasurement(
  block: DocumentBlock,
  blockNode: HTMLElement,
  zoomFactor: number,
): TableRowSplitMeasurement | null {
  if (block.type !== 'table' || block.value.rows.length < 2) {
    return null
  }

  const tableElement = blockNode.querySelector<HTMLTableElement>('.table-block > table')
  const rowNodes = tableElement?.tBodies.item(0)
    ? Array.from(tableElement.tBodies.item(0)!.rows)
    : []

  if (rowNodes.length !== block.value.rows.length) {
    return null
  }

  const rowItems = block.value.rows.map((row, index) => ({
    row,
    actualHeight: Math.max(1, Math.round(rowNodes[index]?.offsetHeight ?? 0)),
    hintHeight: tableRowHeightHint(row, zoomFactor),
  }))

  const totalRowHintHeight = rowItems.reduce(
    (sum, rowItem) => sum + (rowItem.hintHeight ?? rowItem.actualHeight),
    0,
  )
  const tableHintHeight = blockHeightHint(block, zoomFactor)
  const chromeHeight = Math.max(
    0,
    Math.round(
      blockNode.offsetHeight - rowItems.reduce((sum, rowItem) => sum + rowItem.actualHeight, 0),
    ),
  )
  const chromeHintHeight = Math.max(
    0,
    Math.round((tableHintHeight ?? blockNode.offsetHeight) - totalRowHintHeight),
  )
  const safeBreakStarts = computeSafeRowBreakStarts(block.value)
  const headerRowCount = resolveHeaderRowCount(block.value, safeBreakStarts)

  return {
    kind: 'tableRows',
    chromeHeight,
    chromeHintHeight,
    headerRowCount,
    safeBreakStarts,
    rowItems,
  }
}

function buildParagraphSplitMeasurement(
  block: DocumentBlock,
  blockNode: HTMLElement,
  zoomFactor: number,
): ParagraphSplitMeasurement | null {
  if (block.type !== 'paragraph') {
    return null
  }

  const lineRuns = paragraphRunsToLines(block.value)
  if (lineRuns.length < 2) {
    return null
  }

  const topMarginHeight = block.value.style?.marginPrev
    ? Math.round((hwpToPx(block.value.style.marginPrev) ?? 0) * zoomFactor)
    : 0
  const bottomMarginHeight = block.value.style?.marginNext
    ? Math.round((hwpToPx(block.value.style.marginNext) ?? 0) * zoomFactor)
    : 0
  const contentActualHeight = Math.max(
    1,
    Math.round(blockNode.offsetHeight - topMarginHeight - bottomMarginHeight),
  )
  const hintHeight = blockHeightHint(block, zoomFactor) ?? blockNode.offsetHeight
  const contentHintHeight = Math.max(
    1,
    Math.round(hintHeight - topMarginHeight - bottomMarginHeight),
  )
  const lineNodes = Array.from(
    blockNode.querySelectorAll<HTMLElement>('[data-paragraph-line-index]'),
  )
  const lineActualHeights =
    lineNodes.length === lineRuns.length
      ? lineNodes.map((lineNode) => Math.max(1, Math.round(lineNode.offsetHeight)))
      : Array.from({ length: lineRuns.length }, () =>
          Math.max(1, Math.round(contentActualHeight / lineRuns.length)),
        )
  const actualHeightTotal = lineActualHeights.reduce((sum, height) => sum + height, 0)
  const lineHintHeights = lineActualHeights.map((height, index) => {
    if (actualHeightTotal <= 0) {
      return Math.max(1, Math.round(contentHintHeight / lineRuns.length))
    }

    const isLastLine = index === lineActualHeights.length - 1
    if (isLastLine) {
      const usedHint = lineActualHeights
        .slice(0, index)
        .reduce(
          (sum, currentHeight) =>
            sum + Math.max(1, Math.round((contentHintHeight * currentHeight) / actualHeightTotal)),
          0,
        )
      return Math.max(1, contentHintHeight - usedHint)
    }

    return Math.max(1, Math.round((contentHintHeight * height) / actualHeightTotal))
  })

  return {
    kind: 'paragraphLines',
    lineRuns,
    lineActualHeights,
    lineHintHeights,
    topMarginHeight,
    bottomMarginHeight,
  }
}

function splitMeasurementHintHeight(
  splitMeasurement: SplittableBlockMeasurement | null | undefined,
): number | null {
  if (!splitMeasurement) {
    return null
  }

  switch (splitMeasurement.kind) {
    case 'singleCellTable':
      return Math.max(
        1,
        Math.round(
          splitMeasurement.chromeHintHeight +
            splitMeasurement.childItems.reduce(
              (sum, childItem) => sum + (childItem.hintHeight ?? childItem.actualHeight),
              0,
            ),
        ),
      )
    case 'tableRows':
      return Math.max(
        1,
        Math.round(
          splitMeasurement.chromeHintHeight +
            splitMeasurement.rowItems.reduce(
              (sum, rowItem) => sum + (rowItem.hintHeight ?? rowItem.actualHeight),
              0,
            ),
        ),
      )
    case 'paragraphLines':
      return Math.max(
        1,
        Math.round(
          splitMeasurement.topMarginHeight +
            splitMeasurement.lineHintHeights.reduce((sum, lineHeight) => sum + lineHeight, 0) +
            splitMeasurement.bottomMarginHeight,
        ),
      )
  }
}

function measurePaginationItems(
  section: DocumentSection,
  blockNodes: HTMLElement[],
  zoomFactor: number,
): PaginationItem[] {
  return blockNodes.map((blockNode) => {
    const index = Number(blockNode.dataset.blockIndex ?? '0')
    const block = section.blocks[index]
    const splitMeasurement =
      buildTableRowSplitMeasurement(block, blockNode, zoomFactor) ??
      buildSingleCellSplitMeasurement(block, blockNode, zoomFactor) ??
      buildParagraphSplitMeasurement(block, blockNode, zoomFactor)
    const directHintHeight = blockHeightHint(block, zoomFactor)
    const measuredHintHeight = splitMeasurementHintHeight(splitMeasurement)

    return {
      key: `${section.id}-${index}`,
      block,
      actualHeight: Math.max(1, Math.round(blockNode.offsetHeight)),
      hintHeight: directHintHeight ?? measuredHintHeight,
      pageBreakBefore: blockPageBreakBefore(block),
      keepWithNext: blockKeepWithNext(block),
      keepLines: blockKeepLines(block),
      splitMeasurement,
    }
  })
}

function splitParagraphItem(
  item: PaginationItem,
  splitMeasurement: ParagraphSplitMeasurement,
  remainingHeight: number,
  availableHeight: number,
): PaginationItem[] | null {
  if (item.block.type !== 'paragraph' || splitMeasurement.lineRuns.length < 2) {
    return null
  }

  const paragraph = item.block.value
  const totalLines = splitMeasurement.lineRuns.length
  const minBreakLines = totalLines >= MIN_PARAGRAPH_FRAGMENT_LINES * 2
    ? MIN_PARAGRAPH_FRAGMENT_LINES
    : 1
  const rawFragments: Array<{
    lines: TextRun[][]
    actualHeight: number
    hintHeight: number
  }> = []
  let cursor = 0

  while (cursor < totalLines) {
    const pageAllowance = rawFragments.length === 0 ? remainingHeight : availableHeight
    if (pageAllowance <= PAGINATION_TOLERANCE) {
      return null
    }

    const fragmentIndex = rawFragments.length
    const remainingLines = totalLines - cursor
    const topMargin = fragmentIndex === 0 ? splitMeasurement.topMarginHeight : 0
    const remainingActualHeights = splitMeasurement.lineActualHeights.slice(cursor)
    const remainingHintHeights = splitMeasurement.lineHintHeights.slice(cursor)
    const averageRemainingActualHeight =
      remainingActualHeights.reduce((sum, height) => sum + height, 0) /
      remainingActualHeights.length
    const averageRemainingHintHeight =
      remainingHintHeights.reduce((sum, height) => sum + height, 0) /
      remainingHintHeights.length
    const maxFitLines = Math.min(
      remainingLines,
      Math.max(
        1,
        Math.floor(
          (pageAllowance - topMargin) /
            Math.max(
              1,
              paginationHeight(averageRemainingActualHeight, averageRemainingHintHeight),
            ),
        ),
      ),
    )

    let chosenLineCount = 0
    for (let candidate = maxFitLines; candidate >= 1; candidate -= 1) {
      const isLastFragment = candidate === remainingLines
      const linesAfter = remainingLines - candidate
      if (!isLastFragment) {
        if (candidate < minBreakLines || linesAfter < minBreakLines) {
          continue
        }
      } else if (fragmentIndex > 0 && candidate < minBreakLines) {
        continue
      }

      const actualHeight =
        topMargin +
        splitMeasurement.lineActualHeights
          .slice(cursor, cursor + candidate)
          .reduce((sum, height) => sum + height, 0) +
        (isLastFragment ? splitMeasurement.bottomMarginHeight : 0)
      const hintHeight =
        topMargin +
        splitMeasurement.lineHintHeights
          .slice(cursor, cursor + candidate)
          .reduce((sum, height) => sum + height, 0) +
        (isLastFragment ? splitMeasurement.bottomMarginHeight : 0)

      if (!overflowsRemaining(actualHeight, hintHeight, pageAllowance)) {
        chosenLineCount = candidate
        break
      }
    }

    if (chosenLineCount <= 0) {
      return null
    }

    const fragmentLines = splitMeasurement.lineRuns.slice(cursor, cursor + chosenLineCount)
    const isLastFragment = cursor + chosenLineCount === totalLines
    const actualHeight =
      topMargin +
      splitMeasurement.lineActualHeights
        .slice(cursor, cursor + chosenLineCount)
        .reduce((sum, height) => sum + height, 0) +
      (isLastFragment ? splitMeasurement.bottomMarginHeight : 0)
    const hintHeight =
      topMargin +
      splitMeasurement.lineHintHeights
        .slice(cursor, cursor + chosenLineCount)
        .reduce((sum, height) => sum + height, 0) +
      (isLastFragment ? splitMeasurement.bottomMarginHeight : 0)

    rawFragments.push({
      lines: fragmentLines,
      actualHeight: Math.round(actualHeight),
      hintHeight: Math.round(hintHeight),
    })

    cursor += chosenLineCount
  }

  if (rawFragments.length < 2) {
    return null
  }

  return rawFragments.map((fragment, index) => ({
    key: `${item.key}-paragraph-${index}`,
    block: cloneParagraphFragment(
      paragraph,
      fragment.lines,
      index,
      rawFragments.length,
    ),
    actualHeight: fragment.actualHeight,
    hintHeight: fragment.hintHeight,
    pageBreakBefore: index === 0 ? item.pageBreakBefore : false,
    keepWithNext: index === rawFragments.length - 1 ? item.keepWithNext : false,
    keepLines: index === rawFragments.length - 1 ? item.keepLines : false,
    splitMeasurement: null,
  }))
}

function splitSingleCellTableItem(
  item: PaginationItem,
  splitMeasurement: SingleCellTableSplitMeasurement,
  remainingHeight: number,
  availableHeight: number,
): PaginationItem[] | null {
  if (item.block.type !== 'table' || splitMeasurement.childItems.length < 2) {
    return null
  }

  const fragments: PaginationItem[] = []
  const childItems = splitMeasurement.childItems
  let cursor = 0

  while (cursor < childItems.length) {
    const pageAllowance = (fragments.length === 0 ? remainingHeight : availableHeight) - splitMeasurement.chromeHeight
    if (pageAllowance <= PAGINATION_TOLERANCE) {
      return null
    }

    const fragmentStart = cursor
    let usedHeight = 0
    let usedHintHeight = 0

    while (cursor < childItems.length) {
      const childItem = childItems[cursor]
      const nextUsedHeight = usedHeight + childItem.actualHeight
      const nextUsedHintHeight = usedHintHeight + (childItem.hintHeight ?? childItem.actualHeight)
      const overflow =
        cursor > fragmentStart &&
        overflowsRemaining(nextUsedHeight, nextUsedHintHeight, pageAllowance)

      if (overflow) {
        break
      }

      usedHeight = nextUsedHeight
      usedHintHeight = nextUsedHintHeight
      cursor += 1

      if (
        cursor === fragmentStart + 1 &&
        overflowsRemaining(usedHeight, usedHintHeight, pageAllowance)
      ) {
        break
      }
    }

    const fragmentChildren = childItems.slice(fragmentStart, cursor)
    if (fragmentChildren.length === 0) {
      return null
    }

    fragments.push({
      key: `${item.key}-fragment-${fragments.length}`,
      block: cloneTableFragment(
        item.block.value,
        fragmentChildren.map((child) => child.block),
      ),
      actualHeight:
        splitMeasurement.chromeHeight +
        fragmentChildren.reduce((sum, child) => sum + child.actualHeight, 0),
      hintHeight:
        splitMeasurement.chromeHintHeight +
        fragmentChildren.reduce((sum, child) => sum + (child.hintHeight ?? child.actualHeight), 0),
      pageBreakBefore: false,
      keepWithNext: false,
      keepLines: false,
      splitMeasurement: null,
    })
  }

  return fragments.length > 1 ? fragments : null
}

function buildAtomicRowGroups(splitMeasurement: TableRowSplitMeasurement): Array<{
  start: number
  end: number
  actualHeight: number
  hintHeight: number
}> {
  const boundaries = Array.from(splitMeasurement.safeBreakStarts)
    .filter((boundary) => boundary > 0)
    .sort((left, right) => left - right)

  const groups: Array<{
    start: number
    end: number
    actualHeight: number
    hintHeight: number
  }> = []

  let start = 0
  for (const boundary of boundaries) {
    if (boundary <= start) {
      continue
    }

    const rows = splitMeasurement.rowItems.slice(start, boundary)
    if (rows.length === 0) {
      continue
    }

    groups.push({
      start,
      end: boundary,
      actualHeight: rows.reduce((sum, row) => sum + row.actualHeight, 0),
      hintHeight: rows.reduce(
        (sum, row) => sum + (row.hintHeight ?? row.actualHeight),
        0,
      ),
    })
    start = boundary
  }

  return groups
}

function splitTableByRowsItem(
  item: PaginationItem,
  splitMeasurement: TableRowSplitMeasurement,
  remainingHeight: number,
  availableHeight: number,
): PaginationItem[] | null {
  if (item.block.type !== 'table' || splitMeasurement.rowItems.length < 2) {
    return null
  }

  const rowGroups = buildAtomicRowGroups(splitMeasurement)
  if (rowGroups.length < 2) {
    return forceSplitTableByRows(item, splitMeasurement, remainingHeight, availableHeight)
  }

  const fragments: PaginationItem[] = []
  let groupCursor = 0

  while (groupCursor < rowGroups.length) {
    const isFirstFragment = fragments.length === 0
    const pageAllowance = isFirstFragment ? remainingHeight : availableHeight
    let repeatHeader = !isFirstFragment && splitMeasurement.headerRowCount > 0
    let repeatedHeaderRows = repeatHeader
      ? splitMeasurement.rowItems.slice(0, splitMeasurement.headerRowCount)
      : []
    let usedActual =
      splitMeasurement.chromeHeight +
      repeatedHeaderRows.reduce((sum, rowItem) => sum + rowItem.actualHeight, 0)
    let usedHint =
      splitMeasurement.chromeHintHeight +
      repeatedHeaderRows.reduce(
        (sum, rowItem) => sum + (rowItem.hintHeight ?? rowItem.actualHeight),
        0,
      )

    const fragmentStartGroup = groupCursor

    if (
      repeatHeader &&
      groupCursor < rowGroups.length &&
      overflowsRemaining(
        usedActual + rowGroups[groupCursor].actualHeight,
        usedHint + rowGroups[groupCursor].hintHeight,
        pageAllowance,
      ) &&
      !overflowsRemaining(
        splitMeasurement.chromeHeight + rowGroups[groupCursor].actualHeight,
        splitMeasurement.chromeHintHeight + rowGroups[groupCursor].hintHeight,
        pageAllowance,
      )
    ) {
      repeatHeader = false
      repeatedHeaderRows = []
      usedActual = splitMeasurement.chromeHeight
      usedHint = splitMeasurement.chromeHintHeight
    }

    while (groupCursor < rowGroups.length) {
      const group = rowGroups[groupCursor]
      const nextActual = usedActual + group.actualHeight
      const nextHint = usedHint + group.hintHeight
      const overflow =
        groupCursor > fragmentStartGroup &&
        overflowsRemaining(nextActual, nextHint, pageAllowance)

      if (overflow) {
        break
      }

      usedActual = nextActual
      usedHint = nextHint
      groupCursor += 1

      if (
        groupCursor === fragmentStartGroup + 1 &&
        overflowsRemaining(usedActual, usedHint, pageAllowance)
      ) {
        break
      }
    }

    if (groupCursor <= fragmentStartGroup) {
      return null
    }

    const startRow = rowGroups[fragmentStartGroup].start
    const endRow = rowGroups[groupCursor - 1].end

    fragments.push({
      key: `${item.key}-rows-${fragments.length}`,
      block: cloneTableRowsFragment(
        item.block.value,
        startRow,
        endRow,
        repeatHeader ? splitMeasurement.headerRowCount : 0,
      ),
      actualHeight: Math.round(usedActual),
      hintHeight: Math.round(usedHint),
      pageBreakBefore: isFirstFragment ? item.pageBreakBefore : false,
      keepWithNext: false,
      keepLines: false,
      splitMeasurement: null,
    })
  }

  return fragments.length > 1 ? fragments : null
}

function forceSplitTableByRows(
  item: PaginationItem,
  splitMeasurement: TableRowSplitMeasurement,
  _remainingHeight: number,
  _availableHeight: number,
): PaginationItem[] | null {
  if (item.block.type !== 'table') return null

  const table = item.block.value
  const grid = buildRowSpanGrid(table)
  const headerRowCount = splitMeasurement.headerRowCount
  const forcedSplitRow = findBestForcedSplitRow(grid, headerRowCount)

  if (forcedSplitRow == null) return null

  const headerHeight = splitMeasurement.rowItems
    .slice(0, headerRowCount)
    .reduce((sum, ri) => sum + ri.actualHeight, 0)
  const headerHintHeight = splitMeasurement.rowItems
    .slice(0, headerRowCount)
    .reduce((sum, ri) => sum + (ri.hintHeight ?? ri.actualHeight), 0)

  const splitPoints = [forcedSplitRow]
  let cursor = forcedSplitRow
  while (cursor < table.rows.length) {
    const nextSplit = findBestForcedSplitRow(
      grid.slice(cursor).map((row) =>
        row.map((cell) => ({
          ...cell,
          ownerRow: cell.ownerRow - cursor,
        })),
      ),
      0,
    )
    if (nextSplit == null) break
    cursor += nextSplit
    if (cursor < table.rows.length) {
      splitPoints.push(cursor)
    }
  }

  const boundaries = [0, ...splitPoints, table.rows.length]
  const fragments: PaginationItem[] = []

  for (let i = 0; i < boundaries.length - 1; i += 1) {
    const rowStart = boundaries[i]
    const rowEnd = boundaries[i + 1]
    const isFirst = i === 0
    const useHeader = !isFirst && headerRowCount > 0

    const fragmentRows = splitMeasurement.rowItems.slice(rowStart, rowEnd)
    const dataActual = fragmentRows.reduce((sum, ri) => sum + ri.actualHeight, 0)
    const dataHint = fragmentRows.reduce(
      (sum, ri) => sum + (ri.hintHeight ?? ri.actualHeight),
      0,
    )

    const totalActual =
      splitMeasurement.chromeHeight + (useHeader ? headerHeight : 0) + dataActual
    const totalHint =
      splitMeasurement.chromeHintHeight + (useHeader ? headerHintHeight : 0) + dataHint

    fragments.push({
      key: `${item.key}-forced-${i}`,
      block: cloneTableRowsFragmentWithSpanFix(
        table,
        rowStart,
        rowEnd,
        useHeader ? headerRowCount : 0,
        grid,
      ),
      actualHeight: Math.round(totalActual),
      hintHeight: Math.round(totalHint),
      pageBreakBefore: isFirst ? item.pageBreakBefore : false,
      keepWithNext: false,
      keepLines: false,
      splitMeasurement: null,
    })
  }

  return fragments.length > 1 ? fragments : null
}

function paginateSectionBlocks(
  items: PaginationItem[],
  availableHeight: number,
): DocumentBlock[][] {
  if (items.length === 0) {
    return [[]]
  }

  const queue = [...items]
  const pages: DocumentBlock[][] = []
  let currentPage: DocumentBlock[] = []
  let remainingHeight = availableHeight

  const flushPage = () => {
    pages.push(currentPage)
    currentPage = []
    remainingHeight = availableHeight
  }

  while (queue.length > 0) {
    const item = queue.shift()
    if (!item) {
      break
    }

    if (item.pageBreakBefore && currentPage.length > 0) {
      flushPage()
    }

    if (item.keepWithNext && currentPage.length > 0) {
      const nextItem = queue[0]
      if (nextItem) {
        const pairActualHeight = item.actualHeight + nextItem.actualHeight
        const pairHintHeight =
          (item.hintHeight ?? item.actualHeight) +
          (nextItem.hintHeight ?? nextItem.actualHeight)
        if (overflowsRemaining(pairActualHeight, pairHintHeight, remainingHeight)) {
          flushPage()
        }
      }
    }

    if (
      item.keepLines &&
      currentPage.length > 0 &&
      overflowsRemaining(item.actualHeight, item.hintHeight, remainingHeight) &&
      fitsFreshPage(item, availableHeight)
    ) {
      queue.unshift({ ...item, pageBreakBefore: false })
      flushPage()
      continue
    }

    if (overflowsRemaining(item.actualHeight, item.hintHeight, remainingHeight)) {
      if (currentPage.length > 0 && fitsFreshPage(item, availableHeight)) {
        queue.unshift({ ...item, pageBreakBefore: false })
        flushPage()
        continue
      }

      if (item.splitMeasurement) {
        let fragments: PaginationItem[] | null = null

        switch (item.splitMeasurement.kind) {
          case 'singleCellTable':
            fragments = splitSingleCellTableItem(
              item,
              item.splitMeasurement,
              remainingHeight,
              availableHeight,
            )
            break
          case 'tableRows':
            fragments = splitTableByRowsItem(
              item,
              item.splitMeasurement,
              remainingHeight,
              availableHeight,
            )
            break
          case 'paragraphLines':
            fragments = splitParagraphItem(
              item,
              item.splitMeasurement,
              remainingHeight,
              availableHeight,
            )
            break
        }

        if (fragments) {
          const [firstFragment, ...restFragments] = fragments
          if (restFragments.length > 0) {
            queue.unshift(...restFragments)
          }

          if (overflowsRemaining(firstFragment.actualHeight, firstFragment.hintHeight, remainingHeight) && currentPage.length > 0) {
            flushPage()
          }

          currentPage.push(firstFragment.block)
          remainingHeight -= paginationHeight(firstFragment.actualHeight, firstFragment.hintHeight)
          if (remainingHeight <= PAGINATION_TOLERANCE) {
            flushPage()
          }
          continue
        }
      }

      if (currentPage.length > 0) {
        queue.unshift({ ...item, pageBreakBefore: false, keepWithNext: item.keepWithNext })
        flushPage()
        continue
      }
    }

    currentPage.push(item.block)
    remainingHeight -= paginationHeight(item.actualHeight, item.hintHeight)
    if (remainingHeight <= PAGINATION_TOLERANCE && queue.length > 0) {
      flushPage()
    }
  }

  if (currentPage.length > 0 || pages.length === 0) {
    pages.push(currentPage)
  }

  return pages
}

function selectHeaderFooter(
  candidates: HeaderFooter[],
  pageNumberInSection: number,
): HeaderFooter | null {
  if (candidates.length === 0) {
    return null
  }

  const parity = pageNumberInSection % 2 === 0 ? 'EVEN' : 'ODD'
  return (
    candidates.find((candidate) => candidate.applyPageType === parity) ??
    candidates.find((candidate) => candidate.applyPageType === 'BOTH' || !candidate.applyPageType) ??
    candidates[0]
  )
}

function renderParagraph(
  paragraph: ParagraphValue,
  keyPrefix: string,
  zoomFactor: number,
  context: PlaceholderContext,
  insetStyle?: CSSProperties,
  blockSearchMatches?: BlockSearchMatches,
  activeSearchMatchIndex?: number,
) {
  const lines = paragraphRunsToLines(paragraph)
  const mergedStyle = insetStyle
    ? { ...paragraphStyle(paragraph.style, zoomFactor), ...insetStyle }
    : paragraphStyle(paragraph.style, zoomFactor)

  return (
    <p
      key={keyPrefix}
      className="paragraph-block"
      style={mergedStyle}
    >
      {lines.map((line, lineIndex) => (
        <span
          key={`${keyPrefix}-line-${lineIndex}`}
          className="paragraph-line"
          data-paragraph-line-index={lineIndex}
        >
          {lineIndex === 0 && paragraph.marker ? (
            <span
              className="paragraph-marker"
              style={{
                ...paragraphMarkerStyle(paragraph.style, zoomFactor),
                ...textStyle(paragraph.marker.style, zoomFactor),
              }}
            >
              {replacePlaceholders(paragraph.marker.text, context)}
            </span>
          ) : null}
          {line.length > 0 ? (
            line.map((run, runIndex) => {
              const flatRunIndex = lines.slice(0, lineIndex).reduce((s, l) => s + l.length, 0) + runIndex
              const runHighlights = blockSearchMatches?.filter((m) => m.runIndex === flatRunIndex) ?? []
              return (
                <span key={`${keyPrefix}-line-${lineIndex}-run-${runIndex}`} style={textStyle(run.style, zoomFactor)}>
                  {runHighlights.length > 0
                    ? renderHighlightedText(run.text, runHighlights, activeSearchMatchIndex ?? -1, context)
                    : replacePlaceholders(run.text, context)}
                </span>
              )
            })
          ) : (
            <span className="empty-paragraph-marker">{'\u00A0'}</span>
          )}
        </span>
      ))}
    </p>
  )
}

function renderTableCell(
  cell: TableCell,
  keyPrefix: string,
  zoomFactor: number,
  context: PlaceholderContext,
  assets: AssetRef[],
  options?: { ignoreWidth?: boolean },
) {
  const CellTag = cell.isHeader ? 'th' : 'td'
  return (
    <CellTag
      key={keyPrefix}
      className={cell.isHeader ? 'table-cell table-cell-header' : 'table-cell'}
      colSpan={cell.colSpan ?? 1}
      rowSpan={cell.rowSpan ?? 1}
      style={tableCellStyle(cell, zoomFactor, options)}
    >
      {cell.style?.diagonal ? renderTableCellDiagonal(cell.style.diagonal, keyPrefix) : null}
      <div className="table-cell-content" data-table-cell-content="true">
        {cell.blocks.length > 0
          ? cell.blocks.map((block, index) =>
              (
                <div
                  key={`${keyPrefix}-block-wrap-${index}`}
                  className="table-cell-block"
                  data-cell-block-index={index}
                >
                  {renderBlock(block, `${keyPrefix}-block-${index}`, zoomFactor, context, assets)}
                </div>
              ),
            )
          : replacePlaceholders(cell.text, context)}
      </div>
    </CellTag>
  )
}

type SearchHighlight = { runIndex: number; start: number; length: number; matchGlobalIndex: number }

function renderHighlightedText(
  text: string,
  highlights: SearchHighlight[],
  activeMatchIndex: number,
  context: PlaceholderContext,
): React.ReactNode {
  if (highlights.length === 0) return replacePlaceholders(text, context)
  const sorted = [...highlights].sort((a, b) => a.start - b.start)
  const parts: React.ReactNode[] = []
  let cursor = 0
  for (const hl of sorted) {
    if (hl.start > cursor) {
      parts.push(replacePlaceholders(text.slice(cursor, hl.start), context))
    }
    const isActive = hl.matchGlobalIndex === activeMatchIndex
    parts.push(
      <mark
        key={`hl-${hl.matchGlobalIndex}`}
        className={isActive ? 'search-highlight search-highlight-active' : 'search-highlight'}
        data-search-match={hl.matchGlobalIndex}
      >
        {text.slice(hl.start, hl.start + hl.length)}
      </mark>,
    )
    cursor = hl.start + hl.length
  }
  if (cursor < text.length) {
    parts.push(replacePlaceholders(text.slice(cursor), context))
  }
  return parts
}

type BlockSearchMatches = Array<{ runIndex: number; start: number; length: number; matchGlobalIndex: number }>

function renderBlock(
  block: DocumentBlock,
  keyPrefix: string,
  zoomFactor: number,
  context: PlaceholderContext,
  assets: AssetRef[],
  insets?: { marginLeft?: number; marginRight?: number } | null,
  blockSearchMatches?: BlockSearchMatches,
  activeSearchMatchIndex?: number,
) {
  switch (block.type) {
    case 'paragraph': {
      const insetStyle: CSSProperties | undefined =
        insets && (insets.marginLeft || insets.marginRight)
          ? {
              marginLeft: insets.marginLeft ? `${insets.marginLeft}px` : undefined,
              marginRight: insets.marginRight ? `${insets.marginRight}px` : undefined,
            }
          : undefined
      return renderParagraph(block.value, keyPrefix, zoomFactor, context, insetStyle, blockSearchMatches, activeSearchMatchIndex)
    }
    case 'table': {
      const fitToContainer = shouldFitTableToContainer(block.value)
      return (
        <div key={keyPrefix} className="table-block">
          <table style={tableStyle(block.value, zoomFactor)}>
            <tbody>
              {block.value.rows.map((row, rowIndex) => (
                <tr
                  key={`${keyPrefix}-row-${rowIndex}`}
                  data-table-row-index={rowIndex}
                >
                  {row.cells.map((cell, cellIndex) => (
                    renderTableCell(
                      cell,
                      `${keyPrefix}-cell-${rowIndex}-${cellIndex}`,
                      zoomFactor,
                      context,
                      assets,
                      fitToContainer ? { ignoreWidth: true } : undefined,
                    )
                  ))}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )
    }
    case 'image': {
      const asset = resolveAsset(assets, block.value.assetId)
      const canRenderImage = asset?.mediaType.startsWith('image/') && Boolean(asset.dataUri)
      const imageContent = (
        <div
          className={block.value.treatAsChar ? 'image-block image-inline' : 'image-block image-floating'}
          style={objectPlacementStyle(block.value, zoomFactor)}
        >
          {canRenderImage ? (
            <img
              src={asset?.dataUri ?? undefined}
              alt={block.value.altText ?? block.value.caption ?? block.value.kind}
            />
          ) : (
            <>
              <span className="block-label">{block.value.kind.toUpperCase()}</span>
              <strong>{block.value.assetId}</strong>
            </>
          )}
          {block.value.caption ? <p>{replacePlaceholders(block.value.caption, context)}</p> : null}
          {!block.value.caption && block.value.altText ? <p>{replacePlaceholders(block.value.altText, context)}</p> : null}
        </div>
      )

      if (!block.value.treatAsChar) {
        return (
          <div
            key={keyPrefix}
            className="image-floating-anchor"
            style={floatingObjectReserveStyle(block.value, zoomFactor)}
          >
            {imageContent}
          </div>
        )
      }

      return (
        <div key={keyPrefix}>
          {imageContent}
        </div>
      )
    }
    case 'footnote':
      return (
        <aside key={keyPrefix} className="footnote-block">
          <span className="footnote-number">{block.value.number ?? ''}</span>
          <div className="footnote-content">
            {block.value.blocks.map((child, i) =>
              renderBlock(child, `${keyPrefix}-fn-${i}`, zoomFactor, context, assets),
            )}
          </div>
        </aside>
      )
    case 'unsupported':
      return (
        <div key={keyPrefix} className="unsupported-block">
          <span className="block-label warning">Unsupported</span>
          <strong>{block.value.kind}</strong>
          {block.value.reason ? <p>{block.value.reason}</p> : null}
        </div>
      )
  }
}

function LazyPage({
  children,
  enabled,
  style,
  ...rest
}: {
  children: React.ReactNode
  enabled: boolean
  style?: CSSProperties
  className?: string
  'data-page-number'?: number
}) {
  const ref = useRef<HTMLElement | null>(null)
  const [visible, setVisible] = useState(!enabled)

  useEffect(() => {
    if (!enabled) { setVisible(true); return }
    const el = ref.current
    if (!el) return
    const observer = new IntersectionObserver(
      ([entry]) => { if (entry.isIntersecting) setVisible(true) },
      { rootMargin: '200px' },
    )
    observer.observe(el)
    return () => observer.disconnect()
  }, [enabled])

  if (!visible) {
    return (
      <article ref={ref} style={{ ...style, overflow: 'hidden' }} {...rest}>
        <div style={{ minHeight: style?.height ?? '1123px', display: 'flex', alignItems: 'center', justifyContent: 'center', color: '#94a3b8', fontSize: '0.9rem' }}>
          로딩 중...
        </div>
      </article>
    )
  }

  return <article ref={ref} style={style} {...rest}>{children}</article>
}

function App() {
  const [opening, setOpening] = useState(false)
  const [openError, setOpenError] = useState<string | null>(null)
  const [loadedDocument, setLoadedDocument] = useState<LoadedDocument | null>(null)
  const [renderedPages, setRenderedPages] = useState<RenderedPage[]>([])
  const [zoomPreset, setZoomPreset] = useState<ZoomPreset>('fitWidth')
  const [workspaceWidth, setWorkspaceWidth] = useState<number | null>(null)
  const [searchOpen, setSearchOpen] = useState(false)
  const [searchQuery, setSearchQuery] = useState('')
  const [searchMatchIndex, setSearchMatchIndex] = useState(0)
  const searchInputRef = useRef<HTMLInputElement | null>(null)
  const [recentFiles, setRecentFiles] = useState<RecentFile[]>(loadRecentFiles)
  const [dragOver, setDragOver] = useState(false)
  const [darkMode, setDarkMode] = useState(false)
  const deferredDocument = useDeferredValue(loadedDocument)
  const pageWorkspaceRef = useRef<HTMLDivElement | null>(null)
  const measureRefs = useRef<Record<number, HTMLDivElement | null>>({})
  const tauriAvailable = '__TAURI_INTERNALS__' in window

  useEffect(() => {
    const pageWorkspace = pageWorkspaceRef.current
    if (!pageWorkspace || typeof ResizeObserver === 'undefined') {
      return
    }

    const resizeObserver = new ResizeObserver((entries) => {
      const entry = entries[0]
      if (entry) {
        setWorkspaceWidth(entry.contentRect.width)
      }
    })

    resizeObserver.observe(pageWorkspace)
    return () => resizeObserver.disconnect()
  }, [deferredDocument])

  useLayoutEffect(() => {
    if (!deferredDocument) {
      setRenderedPages([])
      return
    }

    const pages: RenderedPage[] = []
    let nextDisplayPageNumber = deferredDocument.document.sections[0]?.pageStartNumber ?? 1

    for (const section of deferredDocument.document.sections) {
      const zoomFactor = resolveZoomFactor(section.pageLayout, zoomPreset, workspaceWidth)
      const measureRoot = measureRefs.current[section.id]
      const bodyRoot = measureRoot?.querySelector<HTMLElement>('[data-region="body"]')
      const headerRoot = measureRoot?.querySelector<HTMLElement>('[data-region="header"]')
      const footerRoot = measureRoot?.querySelector<HTMLElement>('[data-region="footer"]')
      const blockNodes = bodyRoot
        ? Array.from(bodyRoot.querySelectorAll<HTMLElement>('[data-block-index]'))
        : []

      let groups: DocumentBlock[][] = []
      if (blockNodes.length > 0) {
        const availableHeight = Math.max(
          120,
          pageInnerHeight(section.pageLayout, zoomFactor) -
            (headerRoot?.offsetHeight ?? 0) -
            (footerRoot?.offsetHeight ?? 0),
        )
        const paginationItems = measurePaginationItems(section, blockNodes, zoomFactor)
        groups = paginateSectionBlocks(paginationItems, availableHeight)
      }

      if (groups.length === 0) {
        groups = [section.blocks]
      }

      if (section.pageStartNumber != null) {
        nextDisplayPageNumber = section.pageStartNumber
      }

      groups.forEach((group, groupIndex) => {
        pages.push({
          key: `${section.id}-${groupIndex}`,
          sectionId: section.id,
          displayPageNumber: nextDisplayPageNumber + groupIndex,
          sectionPageNumber: groupIndex + 1,
          pageLayout: section.pageLayout,
          blocks: group,
          header: selectHeaderFooter(section.headers, groupIndex + 1),
          footer: selectHeaderFooter(section.footers, groupIndex + 1),
        })
      })

      nextDisplayPageNumber += groups.length
    }

    setRenderedPages(pages)
  }, [deferredDocument, workspaceWidth, zoomPreset, renderedPages.length])

  async function handleOpenDocument() {
    if (!tauriAvailable) {
      setOpenError(
        '실제 문서 로딩은 현재 Tauri desktop shell에서 지원합니다. `pnpm dev`로 데스크톱 앱을 실행해 주세요.',
      )
      return
    }

    setOpening(true)
    setOpenError(null)

    try {
      const document = await openWithNativeDialog()
      if (!document) {
        return
      }

      startTransition(() => {
        setLoadedDocument(document)
        saveRecentFile(document.fileName)
        setRecentFiles(loadRecentFiles())
      })
    } catch (error) {
      setOpenError(error instanceof Error ? error.message : String(error))
    } finally {
      setOpening(false)
    }
  }

  async function handleFileDrop(file: File) {
    if (!tauriAvailable) return
    const ext = file.name.split('.').pop()?.toLowerCase()
    if (ext !== 'hwp' && ext !== 'hwpx') {
      setOpenError('지원하지 않는 파일 형식입니다. .hwp 또는 .hwpx 파일을 사용하세요.')
      return
    }
    setOpening(true)
    setOpenError(null)
    try {
      const bytes = new Uint8Array(await file.arrayBuffer())
      const document = await openWithBytes(file.name, bytes)
      if (document) {
        startTransition(() => {
          setLoadedDocument(document)
          saveRecentFile(document.fileName)
          setRecentFiles(loadRecentFiles())
        })
      }
    } catch (error) {
      setOpenError(error instanceof Error ? error.message : String(error))
    } finally {
      setOpening(false)
    }
  }

  const activeDocument = deferredDocument
  const documentAssets = activeDocument?.document.assets ?? []
  const title = activeDocument?.document.metadata.title ?? activeDocument?.fileName
  const totalPages = renderedPages.length > 0 ? renderedPages.length : activeDocument?.document.sections.length ?? 0

  // --- Keyboard shortcuts ---
  const scrollToPage = useCallback((pageNum: number) => {
    const el = pageWorkspaceRef.current?.querySelector(`[data-page-number="${pageNum}"]`)
    if (el) el.scrollIntoView({ behavior: 'smooth', block: 'start' })
  }, [])

  useEffect(() => {
    const handleKeyDown = (e: KeyboardEvent) => {
      if ((e.ctrlKey || e.metaKey) && e.key === 'f') {
        e.preventDefault()
        if (deferredDocument) {
          setSearchOpen(true)
          setTimeout(() => searchInputRef.current?.focus(), 0)
        }
      }
      if (e.key === 'Escape' && searchOpen) {
        setSearchOpen(false)
        setSearchQuery('')
        setSearchMatchIndex(0)
      }
      // Page navigation
      if (!searchOpen && deferredDocument && !(e.target instanceof HTMLInputElement)) {
        if (e.key === 'PageDown' || (e.key === ' ' && !e.shiftKey)) {
          e.preventDefault()
          pageWorkspaceRef.current?.scrollBy({ top: window.innerHeight * 0.8, behavior: 'smooth' })
        }
        if (e.key === 'PageUp' || (e.key === ' ' && e.shiftKey)) {
          e.preventDefault()
          pageWorkspaceRef.current?.scrollBy({ top: -window.innerHeight * 0.8, behavior: 'smooth' })
        }
        if (e.key === 'Home') {
          e.preventDefault()
          pageWorkspaceRef.current?.scrollTo({ top: 0, behavior: 'smooth' })
        }
        if (e.key === 'End') {
          e.preventDefault()
          pageWorkspaceRef.current?.scrollTo({ top: pageWorkspaceRef.current.scrollHeight, behavior: 'smooth' })
        }
        if ((e.ctrlKey || e.metaKey) && e.key === 'g') {
          e.preventDefault()
          const input = prompt(`페이지 번호 (1-${totalPages})`)
          if (input) {
            const num = parseInt(input, 10)
            if (num >= 1 && num <= totalPages) scrollToPage(num)
          }
        }
      }
    }
    window.addEventListener('keydown', handleKeyDown)
    return () => window.removeEventListener('keydown', handleKeyDown)
  }, [deferredDocument, searchOpen, totalPages, scrollToPage])

  const searchMatches = useMemo(() => {
    if (!searchQuery || searchQuery.length < 2 || renderedPages.length === 0) return []
    const query = searchQuery.toLowerCase()
    const matches: Array<{ pageIndex: number; blockIndex: number; runIndex: number; start: number; length: number }> = []
    renderedPages.forEach((page, pageIndex) => {
      page.blocks.forEach((block, blockIndex) => {
        if (block.type === 'paragraph') {
          block.value.runs.forEach((run, runIndex) => {
            const text = run.text.toLowerCase()
            let pos = 0
            while (pos < text.length) {
              const idx = text.indexOf(query, pos)
              if (idx === -1) break
              matches.push({ pageIndex, blockIndex, runIndex, start: idx, length: query.length })
              pos = idx + 1
            }
          })
        }
      })
    })
    return matches
  }, [searchQuery, renderedPages])

  useEffect(() => {
    if (searchMatches.length > 0 && searchMatchIndex >= 0) {
      const el = document.querySelector(`[data-search-match="${searchMatchIndex}"]`)
      if (el) el.scrollIntoView({ behavior: 'smooth', block: 'center' })
    }
  }, [searchMatchIndex, searchMatches.length])

  const [sidebarOpen, setSidebarOpen] = useState(false)
  const [sidebarTab, setSidebarTab] = useState<'toc' | 'thumbnails'>('toc')

  const tocEntries = useMemo(() => {
    if (!renderedPages.length) return []
    const entries: Array<{ text: string; level: number; pageNumber: number }> = []
    renderedPages.forEach((page) => {
      page.blocks.forEach((block) => {
        if (block.type === 'paragraph' && block.value.style?.headingType === 'OUTLINE') {
          const level = block.value.style?.headingLevel ?? 0
          const text = block.value.runs.map((r) => r.text).join('').trim()
          if (text) entries.push({ text, level, pageNumber: page.displayPageNumber })
        }
      })
    })
    return entries
  }, [renderedPages])

  return (
    <main
      className={`app-shell${dragOver ? ' drag-over' : ''}${darkMode ? ' dark-mode' : ''}`}
      onDragOver={(e) => { e.preventDefault(); setDragOver(true) }}
      onDragLeave={() => setDragOver(false)}
      onDrop={(e) => {
        e.preventDefault()
        setDragOver(false)
        const file = e.dataTransfer.files[0]
        if (file) void handleFileDrop(file)
      }}
    >
      <header className="topbar">
        <div className="brand-lockup">
          <img className="brand-mark" src="/max-viewer-logo.svg" alt="MAX Viewer logo" />
          <div className="title-wrap">
            <p className="app-kicker">HWP / HWPX Viewer</p>
            <h1>MAX Viewer</h1>
            <p className="app-subtitle">한컴 문서를 페이지 형태로 읽는 뷰어</p>
          </div>
        </div>

        <div className="toolbar">
          {activeDocument ? (
            <div className="toolbar-group">
              <span className="meta-pill">{activeDocument.fileName}</span>
              <span className="meta-pill page-nav">
                쪽 <input
                  className="page-jump-input"
                  type="number"
                  min={1}
                  max={totalPages}
                  defaultValue={1}
                  onKeyDown={(e) => {
                    if (e.key === 'Enter') {
                      const num = parseInt((e.target as HTMLInputElement).value, 10)
                      if (num >= 1 && num <= totalPages) scrollToPage(num)
                    }
                  }}
                /> / {totalPages}
              </span>
              <div className="zoom-controls" role="group" aria-label="배율 선택">
                <button
                  type="button"
                  className={`secondary-button ${zoomPreset === 'fitWidth' ? 'active' : ''}`}
                  onClick={() => setZoomPreset('fitWidth')}
                >
                  폭 맞춤
                </button>
                <button
                  type="button"
                  className={`secondary-button ${zoomPreset === '100' ? 'active' : ''}`}
                  onClick={() => setZoomPreset('100')}
                >
                  100%
                </button>
                <button
                  type="button"
                  className={`secondary-button ${zoomPreset === '125' ? 'active' : ''}`}
                  onClick={() => setZoomPreset('125')}
                >
                  125%
                </button>
                <button
                  type="button"
                  className={`secondary-button ${zoomPreset === '150' ? 'active' : ''}`}
                  onClick={() => setZoomPreset('150')}
                >
                  150%
                </button>
              </div>
            </div>
          ) : null}
          {activeDocument ? (
            <>
              <button
                className="secondary-button"
                type="button"
                onClick={() => window.print()}
              >
                인쇄
              </button>
              <button
                className="secondary-button"
                type="button"
                onClick={() => setDarkMode((d) => !d)}
              >
                {darkMode ? '라이트' : '다크'}
              </button>
            </>
          ) : null}
          <button
            className="primary-button"
            type="button"
            onClick={() => void handleOpenDocument()}
            disabled={opening}
          >
            {opening ? '파일 여는 중...' : '파일 열기'}
          </button>
        </div>
      </header>

      {searchOpen && activeDocument ? (
        <div className="search-bar">
          <input
            ref={searchInputRef}
            className="search-input"
            type="text"
            placeholder="검색어 입력 (2자 이상)"
            value={searchQuery}
            onChange={(e) => { setSearchQuery(e.target.value); setSearchMatchIndex(0) }}
            onKeyDown={(e) => {
              if (e.key === 'Enter') {
                e.preventDefault()
                if (e.shiftKey) {
                  setSearchMatchIndex((prev) => (prev - 1 + searchMatches.length) % Math.max(1, searchMatches.length))
                } else {
                  setSearchMatchIndex((prev) => (prev + 1) % Math.max(1, searchMatches.length))
                }
              }
              if (e.key === 'Escape') {
                setSearchOpen(false)
                setSearchQuery('')
                setSearchMatchIndex(0)
              }
            }}
          />
          {searchQuery.length >= 2 ? (
            <span className="search-count">
              {searchMatches.length > 0
                ? `${searchMatchIndex + 1} / ${searchMatches.length}`
                : '결과 없음'}
            </span>
          ) : null}
          <button
            className="secondary-button"
            type="button"
            onClick={() => setSearchMatchIndex((prev) => (prev - 1 + searchMatches.length) % Math.max(1, searchMatches.length))}
            disabled={searchMatches.length === 0}
          >
            ▲
          </button>
          <button
            className="secondary-button"
            type="button"
            onClick={() => setSearchMatchIndex((prev) => (prev + 1) % Math.max(1, searchMatches.length))}
            disabled={searchMatches.length === 0}
          >
            ▼
          </button>
          <button
            className="secondary-button"
            type="button"
            onClick={() => { setSearchOpen(false); setSearchQuery(''); setSearchMatchIndex(0) }}
          >
            닫기
          </button>
        </div>
      ) : null}

      {openError ? <p className="error-banner">{openError}</p> : null}

      {activeDocument ? (
        <section className="document-stage">
          {activeDocument ? (
            <button
              className="secondary-button sidebar-toggle"
              type="button"
              onClick={() => setSidebarOpen((v) => !v)}
            >
              {sidebarOpen ? '패널 닫기' : '목차/썸네일'}
            </button>
          ) : null}

          {sidebarOpen ? (
            <aside className="document-sidebar">
              <div className="sidebar-tabs">
                <button
                  className={`sidebar-tab ${sidebarTab === 'toc' ? 'active' : ''}`}
                  type="button"
                  onClick={() => setSidebarTab('toc')}
                >
                  목차
                </button>
                <button
                  className={`sidebar-tab ${sidebarTab === 'thumbnails' ? 'active' : ''}`}
                  type="button"
                  onClick={() => setSidebarTab('thumbnails')}
                >
                  썸네일
                </button>
              </div>
              {sidebarTab === 'toc' ? (
                <nav className="toc-panel">
                  {tocEntries.length > 0 ? (
                    <ul className="toc-list">
                      {tocEntries.map((entry, i) => (
                        <li
                          key={i}
                          className="toc-item"
                          style={{ paddingLeft: `${entry.level * 1}rem` }}
                        >
                          <button
                            className="toc-link"
                            type="button"
                            onClick={() => scrollToPage(entry.pageNumber)}
                          >
                            {entry.text}
                          </button>
                          <span className="toc-page">{entry.pageNumber}</span>
                        </li>
                      ))}
                    </ul>
                  ) : (
                    <p className="toc-empty">문서에 개요 제목이 없습니다.</p>
                  )}
                </nav>
              ) : (
                <div className="thumbnail-panel">
                  {renderedPages.map((page) => (
                    <button
                      key={page.key}
                      className="thumbnail-item"
                      type="button"
                      onClick={() => scrollToPage(page.displayPageNumber)}
                    >
                      <span className="thumbnail-label">쪽 {page.displayPageNumber}</span>
                    </button>
                  ))}
                </div>
              )}
            </aside>
          ) : null}

          <div className="summary-row">
            <div className="summary-card">
              <span className="summary-label">문서</span>
              <strong>{title}</strong>
            </div>
            <div className="summary-card compact">
              <span className="summary-label">형식</span>
              <strong>{formatLabel(activeDocument.document.format ?? activeDocument.diagnostics.format)}</strong>
            </div>
            <div className="summary-card compact">
              <span className="summary-label">페이지</span>
              <strong>{totalPages}</strong>
            </div>
            <div className="summary-card compact">
              <span className="summary-label">문단</span>
              <strong>{activeDocument.layout.paragraphCount}</strong>
            </div>
          </div>

          {activeDocument.diagnostics.isEncrypted ? (
            <div className="info-banner encrypted-notice">
              <p>이 문서는 암호화되어 있어 본문을 완전히 표시할 수 없습니다.</p>
              <p>암호 해제 기능은 향후 업데이트에서 지원할 예정입니다.</p>
            </div>
          ) : null}

          <div className="measure-layer" aria-hidden="true">
            {activeDocument.document.sections.map((section) => {
              const zoomFactor = resolveZoomFactor(section.pageLayout, zoomPreset, workspaceWidth)
              const measureContext: PlaceholderContext = { pageNumber: 1, totalPages: totalPages || 1 }

              return (
                <div
                  key={`measure-${section.id}`}
                  ref={(element) => {
                    measureRefs.current[section.id] = element
                  }}
                  className="paper-page paper-measure"
                  style={pageMeasureStyle(section.pageLayout, zoomFactor)}
                >
                  {selectHeaderFooter(section.headers, 1) ? (
                    <div data-region="header" className="page-region page-header">
                      {selectHeaderFooter(section.headers, 1)?.blocks.map((block, index) =>
                        renderBlock(
                          block,
                          `measure-header-${section.id}-${index}`,
                          zoomFactor,
                          measureContext,
                          documentAssets,
                        ),
                      )}
                    </div>
                  ) : null}
                  <div data-region="body" className="page-content">
                    {section.blocks.map((block, index) => (
                      <div
                        key={`measure-block-${section.id}-${index}`}
                        data-block-index={index}
                        data-page-break-before={blockPageBreakBefore(block)}
                        className="measure-block"
                      >
                        {renderBlock(
                          block,
                          `measure-body-${section.id}-${index}`,
                          zoomFactor,
                          measureContext,
                          documentAssets,
                        )}
                      </div>
                    ))}
                  </div>
                  {selectHeaderFooter(section.footers, 1) ? (
                    <div data-region="footer" className="page-region page-footer">
                      {selectHeaderFooter(section.footers, 1)?.blocks.map((block, index) =>
                        renderBlock(
                          block,
                          `measure-footer-${section.id}-${index}`,
                          zoomFactor,
                          measureContext,
                          documentAssets,
                        ),
                      )}
                    </div>
                  ) : null}
                </div>
              )
            })}
          </div>

          <div ref={pageWorkspaceRef} className="page-workspace">
            <div className="page-stack">
              {renderedPages.map((page, pageIndex) => {
                const zoomFactor = resolveZoomFactor(page.pageLayout, zoomPreset, workspaceWidth)
                const context: PlaceholderContext = {
                  pageNumber: page.displayPageNumber,
                  totalPages,
                }

                return (
                  <LazyPage
                    key={page.key}
                    className="paper-page"
                    data-page-number={page.displayPageNumber}
                    style={pageStyle(page.pageLayout, zoomFactor)}
                    enabled={renderedPages.length > 10}
                  >
                    <div className="page-caption">쪽 {page.displayPageNumber}</div>
                    {page.header ? (
                      <div className="page-region page-header">
                        {page.header.blocks.map((block, index) =>
                          renderBlock(
                            block,
                            `page-header-${page.key}-${index}`,
                            zoomFactor,
                            context,
                            documentAssets,
                          ),
                        )}
                      </div>
                    ) : null}
                    <div className="page-content page-body">
                      {page.blocks.length > 0 ? (
                        (() => {
                          const cw = pageContentWidth(page.pageLayout, zoomFactor)
                          const exclusions = collectExclusionZones(page.blocks, cw, zoomFactor)
                          let cumulativeHeight = 0
                          return page.blocks.map((block, index) => {
                            const estimatedHeight = block.type === 'paragraph'
                              ? Math.round((hwpToPx(block.value.layoutHeightHint) ?? 24) * zoomFactor)
                              : block.type === 'image' ? Math.round((hwpToPx(block.value.height) ?? 48) * zoomFactor)
                              : 48
                            const blockTop = cumulativeHeight
                            const blockBottom = cumulativeHeight + estimatedHeight
                            cumulativeHeight = blockBottom
                            const insets = exclusions.length > 0
                              ? calculateParagraphInsets(blockTop, blockBottom, exclusions, cw)
                              : null
                            const blockMatches: BlockSearchMatches = searchMatches
                              .filter((m) => m.pageIndex === pageIndex && m.blockIndex === index)
                              .map((m) => ({ runIndex: m.runIndex, start: m.start, length: m.length, matchGlobalIndex: searchMatches.indexOf(m) }))
                            return renderBlock(
                              block,
                              `page-body-${page.key}-${index}`,
                              zoomFactor,
                              context,
                              documentAssets,
                              insets,
                              blockMatches.length > 0 ? blockMatches : undefined,
                              searchMatchIndex,
                            )
                          })
                        })()
                      ) : (
                        <div className="document-empty">
                          <strong>표시할 본문이 없습니다.</strong>
                          <p>{activeDocument.plainText || '문서에서 표시 가능한 텍스트를 찾지 못했습니다.'}</p>
                        </div>
                      )}
                    </div>
                    {page.footer ? (
                      <div className="page-region page-footer">
                        {page.footer.blocks.map((block, index) =>
                          renderBlock(
                            block,
                            `page-footer-${page.key}-${index}`,
                            zoomFactor,
                            context,
                            documentAssets,
                          ),
                        )}
                      </div>
                    ) : null}
                  </LazyPage>
                )
              })}
            </div>
          </div>
        </section>
      ) : (
        <section className="empty-stage">
          <div className="empty-card">
            <p className="app-kicker">시작하기</p>
            <h2>문서를 열어 한컴 문서처럼 확인하세요</h2>
            <p>
              `.hwp` 또는 `.hwpx` 파일을 열면 페이지 윤곽과 본문 내용을 이 화면에 바로 표시합니다.
            </p>
            <div className="empty-actions">
              <button
                className="primary-button"
                type="button"
                onClick={() => void handleOpenDocument()}
                disabled={opening}
              >
                {opening ? '파일 여는 중...' : '파일 열기'}
              </button>
              <span className="empty-hint">지원 형식: .hwp, .hwpx — 파일을 끌어다 놓을 수도 있습니다</span>
            </div>
            {recentFiles.length > 0 ? (
              <div className="recent-files">
                <p className="recent-label">최근 문서</p>
                <ul className="recent-list">
                  {recentFiles.map((file, i) => (
                    <li key={i} className="recent-item">
                      <span className="recent-name">{file.name}</span>
                      <span className="recent-date">{new Date(file.openedAt).toLocaleDateString()}</span>
                    </li>
                  ))}
                </ul>
              </div>
            ) : null}
          </div>
        </section>
      )}
    </main>
  )
}

export default App
