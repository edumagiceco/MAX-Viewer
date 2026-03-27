import {
  startTransition,
  useDeferredValue,
  useEffect,
  useLayoutEffect,
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
  markerAlign?: string | null
  markerWidthAdjust?: number | null
  markerTextOffsetType?: string | null
  markerTextOffset?: number | null
}

type ParagraphValue = {
  marker?: TextRun | null
  runs: TextRun[]
  style?: ParagraphStyle | null
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
  isHeader: boolean
}

type TableRow = {
  cells: TableCell[]
}

type TableValue = {
  rows: TableRow[]
  width?: number | null
  height?: number | null
  cellSpacing?: number | null
  repeatHeader: boolean
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
  caption?: string | null
}

type UnsupportedValue = {
  kind: string
  reason?: string | null
}

type DocumentBlock =
  | { type: 'paragraph'; value: ParagraphValue }
  | { type: 'table'; value: TableValue }
  | { type: 'image'; value: ImageValue }
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

const DEFAULT_PAGE_WIDTH = 794
const DEFAULT_PAGE_HEIGHT = 1123

async function openWithNativeDialog(): Promise<LoadedDocument | null> {
  const { invoke } = await import('@tauri-apps/api/core')
  return await invoke<LoadedDocument | null>('pick_and_open_document')
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

  if (style.lineSpacingType === 'PERCENT') {
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
  return hwpToPx(pageLayout?.width) ?? DEFAULT_PAGE_WIDTH
}

function basePageHeight(pageLayout?: PageLayout | null): number {
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

  return {
    width: `${width}px`,
    height: `${height}px`,
    paddingTop: `${paddingTop}px`,
    paddingRight: `${paddingRight}px`,
    paddingBottom: `${paddingBottom}px`,
    paddingLeft: `${paddingLeft}px`,
    flexShrink: 0,
  }
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
    fontFamily: style?.fontFamily
      ? `"${style.fontFamily}", "Malgun Gothic", "Apple SD Gothic Neo", "Noto Sans KR", sans-serif`
      : undefined,
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

function tableCellStyle(cell: TableCell, zoomFactor: number): CSSProperties {
  return {
    width: cell.width ? `${Math.round((hwpToPx(cell.width) ?? 0) * zoomFactor)}px` : undefined,
    height: cell.height ? `${Math.round((hwpToPx(cell.height) ?? 0) * zoomFactor)}px` : undefined,
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
}

function resolveAsset(assets: AssetRef[], assetId: string): AssetRef | undefined {
  return (
    assets.find((asset) => asset.id === assetId) ??
    assets.find((asset) => asset.sourcePath === assetId) ??
    assets.find((asset) => asset.id.split('/').pop() === assetId) ??
    assets.find((asset) => asset.id.split('/').pop()?.replace(/\.[^.]+$/, '') === assetId)
  )
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

  const style: CSSProperties = {
    position: 'absolute',
    zIndex: image.zOrder ?? 1,
    width: width ? `${width}px` : undefined,
    height: height ? `${height}px` : undefined,
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

  return style
}

function blockPageBreakBefore(block: DocumentBlock): boolean {
  return block.type === 'paragraph' ? block.value.pageBreakBefore : false
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
) {
  return (
    <p
      key={keyPrefix}
      className="paragraph-block"
      style={paragraphStyle(paragraph.style, zoomFactor)}
    >
      {paragraph.marker ? (
        <span
          className="paragraph-marker"
          style={{ ...paragraphMarkerStyle(paragraph.style, zoomFactor), ...textStyle(paragraph.marker.style, zoomFactor) }}
        >
          {replacePlaceholders(paragraph.marker.text, context)}
        </span>
      ) : null}
      {paragraph.runs.length > 0 ? (
        paragraph.runs.map((run, runIndex) => (
          <span key={`${keyPrefix}-run-${runIndex}`} style={textStyle(run.style, zoomFactor)}>
            {replacePlaceholders(run.text, context)}
          </span>
        ))
      ) : (
        <span className="empty-paragraph-marker">{'\u00A0'}</span>
      )}
    </p>
  )
}

function renderTableCell(
  cell: TableCell,
  keyPrefix: string,
  zoomFactor: number,
  context: PlaceholderContext,
  assets: AssetRef[],
) {
  const CellTag = cell.isHeader ? 'th' : 'td'
  return (
    <CellTag
      key={keyPrefix}
      className={cell.isHeader ? 'table-cell table-cell-header' : 'table-cell'}
      colSpan={cell.colSpan ?? 1}
      rowSpan={cell.rowSpan ?? 1}
      style={tableCellStyle(cell, zoomFactor)}
    >
      <div className="table-cell-content">
        {cell.blocks.length > 0
          ? cell.blocks.map((block, index) =>
              renderBlock(block, `${keyPrefix}-block-${index}`, zoomFactor, context, assets),
            )
          : replacePlaceholders(cell.text, context)}
      </div>
    </CellTag>
  )
}

function renderBlock(
  block: DocumentBlock,
  keyPrefix: string,
  zoomFactor: number,
  context: PlaceholderContext,
  assets: AssetRef[],
) {
  switch (block.type) {
    case 'paragraph':
      return renderParagraph(block.value, keyPrefix, zoomFactor, context)
    case 'table':
      return (
        <div key={keyPrefix} className="table-block">
          <table
            style={{
              width: block.value.width
                ? `${Math.round((hwpToPx(block.value.width) ?? 0) * zoomFactor)}px`
                : undefined,
            }}
          >
            <tbody>
              {block.value.rows.map((row, rowIndex) => (
                <tr key={`${keyPrefix}-row-${rowIndex}`}>
                  {row.cells.map((cell, cellIndex) => (
                    renderTableCell(
                      cell,
                      `${keyPrefix}-cell-${rowIndex}-${cellIndex}`,
                      zoomFactor,
                      context,
                      assets,
                    )
                  ))}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )
    case 'image': {
      const asset = resolveAsset(assets, block.value.assetId)
      const canRenderImage = asset?.mediaType.startsWith('image/') && Boolean(asset.dataUri)
      return (
        <div
          key={keyPrefix}
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
    }
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

function App() {
  const [opening, setOpening] = useState(false)
  const [openError, setOpenError] = useState<string | null>(null)
  const [loadedDocument, setLoadedDocument] = useState<LoadedDocument | null>(null)
  const [renderedPages, setRenderedPages] = useState<RenderedPage[]>([])
  const [zoomPreset, setZoomPreset] = useState<ZoomPreset>('fitWidth')
  const [workspaceWidth, setWorkspaceWidth] = useState<number | null>(null)
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

      let groups: number[][] = []
      if (blockNodes.length > 0) {
        const availableHeight = Math.max(
          120,
          pageInnerHeight(section.pageLayout, zoomFactor) -
            (headerRoot?.offsetHeight ?? 0) -
            (footerRoot?.offsetHeight ?? 0),
        )

        let currentPage: number[] = []
        let currentPageTop = blockNodes[0]?.offsetTop ?? 0

        for (const blockNode of blockNodes) {
          const blockIndex = Number(blockNode.dataset.blockIndex ?? '0')
          const block = section.blocks[blockIndex]
          const blockBottom = blockNode.offsetTop + blockNode.offsetHeight
          const shouldBreak =
            currentPage.length > 0 &&
            (blockPageBreakBefore(block) || blockBottom - currentPageTop > availableHeight)

          if (shouldBreak) {
            groups.push(currentPage)
            currentPage = [blockIndex]
            currentPageTop = blockNode.offsetTop
          } else {
            if (currentPage.length === 0) {
              currentPageTop = blockNode.offsetTop
            }
            currentPage.push(blockIndex)
          }
        }

        if (currentPage.length > 0) {
          groups.push(currentPage)
        }
      }

      if (groups.length === 0) {
        groups = [section.blocks.map((_, index) => index)]
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
          blocks: group.map((index) => section.blocks[index]),
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
      })
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

  return (
    <main className="app-shell">
      <header className="topbar">
        <div className="title-wrap">
          <p className="app-kicker">HWP / HWPX Viewer</p>
          <h1>MAX-Viewer</h1>
          <p className="app-subtitle">한컴 문서를 페이지 형태로 읽는 뷰어</p>
        </div>

        <div className="toolbar">
          {activeDocument ? (
            <div className="toolbar-group">
              <span className="meta-pill">{activeDocument.fileName}</span>
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

      {openError ? <p className="error-banner">{openError}</p> : null}

      {activeDocument ? (
        <section className="document-stage">
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
            <p className="info-banner">암호화된 문서라 일부 내용은 표시되지 않을 수 있습니다.</p>
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
              {renderedPages.map((page) => {
                const zoomFactor = resolveZoomFactor(page.pageLayout, zoomPreset, workspaceWidth)
                const context: PlaceholderContext = {
                  pageNumber: page.displayPageNumber,
                  totalPages,
                }

                return (
                  <article
                    key={page.key}
                    className="paper-page"
                    style={pageStyle(page.pageLayout, zoomFactor)}
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
                        page.blocks.map((block, index) =>
                          renderBlock(
                            block,
                            `page-body-${page.key}-${index}`,
                            zoomFactor,
                            context,
                            documentAssets,
                          ),
                        )
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
                  </article>
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
              <span className="empty-hint">지원 형식: .hwp, .hwpx</span>
            </div>
          </div>
        </section>
      )}
    </main>
  )
}

export default App
