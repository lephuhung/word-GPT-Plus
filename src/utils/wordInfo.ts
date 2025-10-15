function twipsToInches(twips: number): number {
  return Math.round((twips / 1440) * 100) / 100
}

function twipsToCm(twips: number): number {
  return Math.round((twips / 566.929) * 100) / 100
}

function twipsToPoints(twips: number): number {
  return Math.round((twips / 20) * 100) / 100
}

function parseMarginsFromOoxml(ooxml: string): { top?: number; right?: number; bottom?: number; left?: number } {
  // Find the last <w:pgMar .../> which contains margin info in twips
  const matches = [...ooxml.matchAll(/<w:pgMar[^>]*?>/g)]
  if (!matches.length) return {}
  const last = matches[matches.length - 1][0]
  const extract = (name: string): number | undefined => {
    const m = last.match(new RegExp(`${name}="(\\d+)"`))
    return m ? parseInt(m[1], 10) : undefined
  }
  return {
    top: extract('w:top'),
    right: extract('w:right'),
    bottom: extract('w:bottom'),
    left: extract('w:left')
  }
}

function parsePageSizeFromOoxml(
  ooxml: string
): { width?: number; height?: number; orient?: 'portrait' | 'landscape' } {
  // Find the last <w:pgSz .../> which contains page size info in twips
  const matches = [...ooxml.matchAll(/<w:pgSz[^>]*?>/g)]
  if (!matches.length) return {}
  const last = matches[matches.length - 1][0]
  const extractNum = (name: string): number | undefined => {
    const m = last.match(new RegExp(`${name}="(\\d+)"`))
    return m ? parseInt(m[1], 10) : undefined
  }
  const extractStr = (name: string): string | undefined => {
    const m = last.match(new RegExp(`${name}="([^"]+)"`))
    return m ? m[1] : undefined
  }
  const orient = extractStr('w:orient') as 'portrait' | 'landscape' | undefined
  return {
    width: extractNum('w:w'),
    height: extractNum('w:h'),
    orient
  }
}

function parseLineSpacingFromOoxml(
  ooxml: string
): { line?: number; lineRule?: string } | undefined {
  // Try to find the first paragraph spacing definition
  const pPrIndex = ooxml.indexOf('<w:pPr')
  let searchStr = ooxml
  if (pPrIndex >= 0) {
    searchStr = ooxml.slice(pPrIndex)
  }
  const m = searchStr.match(/<w:spacing[^>]*?>/)
  if (!m) return undefined
  const tag = m[0]
  const num = (attr: string): number | undefined => {
    const r = tag.match(new RegExp(`${attr}="(\\d+)"`))
    return r ? parseInt(r[1], 10) : undefined
  }
  const str = (attr: string): string | undefined => {
    const r = tag.match(new RegExp(`${attr}="([^"]+)"`))
    return r ? r[1] : undefined
  }
  return { line: num('w:line'), lineRule: str('w:lineRule') }
}

export interface WordDocumentInfo {
  marginsTwips?: { top?: number; right?: number; bottom?: number; left?: number }
  marginsInches?: { top?: number; right?: number; bottom?: number; left?: number }
  marginsCm?: { top?: number; right?: number; bottom?: number; left?: number }
  font?: { name?: string; size?: number; bold?: boolean; italic?: boolean }
  pageSizeTwips?: { width?: number; height?: number; orient?: 'portrait' | 'landscape' }
  pageSizeCm?: { width?: number; height?: number; orient?: 'portrait' | 'landscape' }
  paperKind?: 'A4' | 'Letter' | 'Legal' | 'Other'
  lineSpacingPoints?: number
  lineSpacingRule?: string
  lineSpacingMultiple?: number
  wordCount?: number
}

export async function getWordDocumentInfo(): Promise<WordDocumentInfo> {
  return Word.run(async context => {
    const body = context.document.body
    const ooxml = body.getOoxml()

    // Try selection font first
    const range = context.document.getSelection()
    range.font.load('name,size,bold,italic')

    // Also try first paragraph font as fallback
    const firstParagraph = body.paragraphs.getFirst()
    firstParagraph.font.load('name,size,bold,italic')

    // Load entire document text for word count
    const wholeRange = body.getRange()
    wholeRange.load('text')

    await context.sync()

    const xml = ooxml.value || ''
    const twips = parseMarginsFromOoxml(xml)
    const page = parsePageSizeFromOoxml(xml)
    const spacing = parseLineSpacingFromOoxml(xml)

    const info: WordDocumentInfo = {
      marginsTwips: twips,
      marginsInches: twips
        ? {
            top: twips.top !== undefined ? twipsToInches(twips.top) : undefined,
            right: twips.right !== undefined ? twipsToInches(twips.right) : undefined,
            bottom: twips.bottom !== undefined ? twipsToInches(twips.bottom) : undefined,
            left: twips.left !== undefined ? twipsToInches(twips.left) : undefined
          }
        : undefined,
      marginsCm: twips
        ? {
            top: twips.top !== undefined ? twipsToCm(twips.top) : undefined,
            right: twips.right !== undefined ? twipsToCm(twips.right) : undefined,
            bottom: twips.bottom !== undefined ? twipsToCm(twips.bottom) : undefined,
            left: twips.left !== undefined ? twipsToCm(twips.left) : undefined
          }
        : undefined,
      font: {
        name: range.font.name || firstParagraph.font.name,
        size: range.font.size || firstParagraph.font.size,
        bold: !!(range.font.bold ?? firstParagraph.font.bold),
        italic: !!(range.font.italic ?? firstParagraph.font.italic)
      },
      pageSizeTwips: page,
      pageSizeCm: page && page.width !== undefined && page.height !== undefined
        ? { width: twipsToCm(page.width), height: twipsToCm(page.height), orient: page.orient }
        : undefined,
      paperKind: (() => {
        if (!page || page.width === undefined || page.height === undefined) return undefined
        const w = twipsToCm(page.width)
        const h = twipsToCm(page.height)
        const min = Math.min(w, h)
        const max = Math.max(w, h)
        const approx = (a: number, b: number, tol = 0.5) => Math.abs(a - b) <= tol
        if (approx(min, 21.0) && approx(max, 29.7)) return 'A4'
        if (approx(min, 21.59) && approx(max, 27.94)) return 'Letter'
        if (approx(min, 21.59) && approx(max, 35.56)) return 'Legal'
        return 'Other'
      })(),
      lineSpacingPoints: spacing?.line !== undefined ? twipsToPoints(spacing.line) : undefined,
      lineSpacingRule: spacing?.lineRule,
      lineSpacingMultiple:
        spacing?.line !== undefined && (range.font.size || firstParagraph.font.size)
          ? Math.round((spacing.line / ((range.font.size || firstParagraph.font.size) * 20)) * 100) / 100
          : undefined,
      wordCount: (() => {
        const text = wholeRange.text || ''
        const tokens = text
          .trim()
          .split(/\s+/)
          .filter(t => /[A-Za-zÀ-ỹ0-9]/.test(t))
        return tokens.length
      })()
    }

    return info
  })
}

export function formatWordInfoVi(info: WordDocumentInfo): string {
  const mCm = info.marginsCm || {}
  const f = info.font || {}

  const toStr = (obj: any, unit: string) =>
    `Trên: ${obj.top ?? 'N/A'}${unit}, Phải: ${obj.right ?? 'N/A'}${unit}, Dưới: ${obj.bottom ?? 'N/A'}${unit}, Trái: ${obj.left ?? 'N/A'}${unit}`

  const lines: string[] = []
  if (info.pageSizeCm || info.paperKind) {
    const ps = info.pageSizeCm
    const w = ps?.width !== undefined ? `${ps.width}cm` : 'N/A'
    const h = ps?.height !== undefined ? `${ps.height}cm` : 'N/A'
    const orient = ps?.orient || 'portrait'
    lines.push(`Kích thước giấy: ${info.paperKind ?? 'Không xác định'} (${w} x ${h}), Hướng: ${orient}`)
  }
  if (info.marginsCm) {
    lines.push(`Khoảng cách lề (cm): ${toStr(mCm, 'cm')}`)
  }
  lines.push(
    `Phông chữ: ${f.name ?? 'N/A'}, Cỡ: ${f.size ?? 'N/A'}pt, Đậm: ${f.bold ? 'Có' : 'Không'}, Nghiêng: ${f.italic ? 'Có' : 'Không'}`
  )
  if (info.lineSpacingPoints !== undefined) {
    const multi = info.lineSpacingMultiple !== undefined ? ` (~ ${info.lineSpacingMultiple}x)` : ''
    lines.push(`Giãn dòng: ${info.lineSpacingPoints}pt${multi}`)
  }
  if (info.wordCount !== undefined) {
    lines.push(`Số chữ: ${info.wordCount}`)
  }

  return lines.join('\n')
}