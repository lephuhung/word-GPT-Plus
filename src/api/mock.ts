export type FormatStyle =
  | 'bold'
  | 'italic'
  | 'heading1'
  | 'heading2'
  | 'heading3'
  | 'heading4'
  | 'heading5'
  | 'heading6'
  | 'code'
  | 'quote'

export type ListType = 'bullet' | 'number'

export interface FormatBlockParagraph {
  type: 'paragraph'
  text: string
  styles?: FormatStyle[]
}

export interface FormatBlockHeading {
  type: 'heading'
  text: string
  level: 1 | 2 | 3 | 4 | 5 | 6
}

export interface FormatBlockQuote {
  type: 'quote'
  text: string
}

export interface FormatBlockCode {
  type: 'code'
  text: string
}

export interface FormatBlockListItem {
  text: string
}

export interface FormatBlockList {
  type: 'list'
  listType: ListType
  items: FormatBlockListItem[]
}

export type FormatBlock =
  | FormatBlockParagraph
  | FormatBlockHeading
  | FormatBlockQuote
  | FormatBlockCode
  | FormatBlockList

export interface FormatDefaults {
  fontName?: string
  fontSize?: number
  spaceAfter?: number
  spaceBefore?: number
}

export interface FormatConfig {
  title?: string
  defaults?: FormatDefaults
  blocks: FormatBlock[]
}

export async function getFormatConfig(): Promise<FormatConfig> {
  // Mock cấu hình format toàn trang để bắt đầu soạn thảo
  return {
    title: 'Tiêu đề tài liệu',
    defaults: {
      fontName: 'Times New Roman',
      fontSize: 12,
      spaceAfter: 8,
      spaceBefore: 0,
    },
    blocks: [
      { type: 'heading', text: 'Giới thiệu', level: 1 },
      {
        type: 'paragraph',
        text:
          'Đây là đoạn mở đầu. Bạn có thể bắt đầu soạn thảo nội dung ở đây.',
      },
      { type: 'heading', text: 'Nội dung chính', level: 2 },
      {
        type: 'list',
        listType: 'bullet',
        items: [
          { text: 'Ý chính thứ nhất' },
          { text: 'Ý chính thứ hai' },
          { text: 'Ý chính thứ ba' },
        ],
      },
      { type: 'heading', text: 'Ví dụ', level: 2 },
      { type: 'code', text: 'console.log("Hello World");' },
      { type: 'quote', text: 'Một trích dẫn nhấn mạnh nội dung quan trọng.' },
      {
        type: 'paragraph',
        styles: ['bold'],
        text: 'Kết luận: Định dạng đã được áp dụng theo mẫu.',
      },
    ],
  }
}

export default {
  getFormatConfig,
}