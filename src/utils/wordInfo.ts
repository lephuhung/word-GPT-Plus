function twipsToInches(twips: number): number {
  return Math.round((twips / 1440) * 100) / 100
}

function twipsToCm(twips: number): number {
  return Math.round((twips / 566.929) * 100) / 100
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

export interface WordDocumentInfo {
  marginsTwips?: { top?: number; right?: number; bottom?: number; left?: number }
  marginsInches?: { top?: number; right?: number; bottom?: number; left?: number }
  marginsCm?: { top?: number; right?: number; bottom?: number; left?: number }
  font?: { name?: string; size?: number; bold?: boolean; italic?: boolean }
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

    await context.sync()

    const xml = ooxml.value || ''
    const twips = parseMarginsFromOoxml(xml)

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
      }
    }

    return info
  })
}

export function formatWordInfoVi(info: WordDocumentInfo): string {
  const mIn = info.marginsInches || {}
  const mCm = info.marginsCm || {}
  const mTw = info.marginsTwips || {}
  const f = info.font || {}

  const toStr = (obj: any, unit: string) =>
    `Trên: ${obj.top ?? 'N/A'}${unit}, Phải: ${obj.right ?? 'N/A'}${unit}, Dưới: ${obj.bottom ?? 'N/A'}${unit}, Trái: ${obj.left ?? 'N/A'}${unit}`

  const lines: string[] = []
  if (info.marginsTwips) {
    lines.push(`Khoảng cách lề (twips): ${toStr(mTw, '')}`)
  }
  if (info.marginsInches) {
    lines.push(`Khoảng cách lề (inch): ${toStr(mIn, 'in')}`)
  }
  if (info.marginsCm) {
    lines.push(`Khoảng cách lề (cm): ${toStr(mCm, 'cm')}`)
  }
  lines.push(
    `Phông chữ: ${f.name ?? 'N/A'}, Cỡ: ${f.size ?? 'N/A'}pt, Đậm: ${f.bold ? 'Có' : 'Không'}, Nghiêng: ${f.italic ? 'Có' : 'Không'}`
  )

  return lines.join('\n')
}