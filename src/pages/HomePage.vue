<template>
  <div class="homepage">
    <!-- Header with branding -->
    <div class="header">
      <div class="brand">
        <Zap class="brand-icon" />
        <h1 class="brand-title">Word GPT+</h1>
      </div>
      <div class="actions">
        <button class="settings-btn user-btn" @click="userPage" :disabled="loading">
          <User class="icon" />
          <span class="username-text">{{ username }}</span>
        </button>
        <button class="settings-btn" @click="settings" :disabled="loading">
          <Settings class="icon" />
        </button>
      </div>
    </div>
    <!-- Main Actions removed per request -->
    <!-- Word Info -->
    <!-- <div class="collapsible-section">
      <div class="collapsible-header" @click="wordInfoExpanded = !wordInfoExpanded">
        <div class="collapsible-title">
          <Info class="collapsible-icon" />
          <span>Word Info</span>
        </div>
        <ChevronDown class="collapse-icon" :class="{ expanded: wordInfoExpanded }" />
      </div>
      <div :class="['collapsible-content', { expanded: wordInfoExpanded }]">
        <div class="settings-grid">
          <div class="setting-item">
            <textarea class="result-textarea" v-model="wordInfo" readonly></textarea>
          </div>
        </div>
      </div>
    </div> -->
    <!-- Tùy chọn định dạng: chuyển sang Popup -->
    <!-- Word Info section removed per request -->

    <!-- Progress Indicator -->
    <div v-if="loading" class="progress-section">
      <div class="progress-indicator">
        <Loader2 class="spinning-icon" />
        <span class="progress-text">AI is processing...</span>
      </div>
    </div>


    <!-- Khung chat toàn trang -->
    <div class="section chat-section">
     
      <div class="chat-window">
      <template v-for="(msg, idx) in historyDialog">
        <div
          :class="['chat-msg', msg.role === 'assistant' ? 'assistant' : 'user', msg.notice ? 'notice' : '']">
          <div class="chat-meta">{{ msg.meta || (msg.role === 'assistant' ? 'Word GPT' : 'Bạn') }}</div>
          <pre class="chat-content">{{ msg.content }}</pre>
          <div v-if="msg.role === 'assistant' && !msg.notice && !msg.noActions" class="chat-actions internal">
            <button class="chat-action-btn replace-btn" @click="insertFromMessage(idx, 'replace')" :disabled="loading" :title="$t('replace')" aria-label="Thay thế">
              <svg class="action-icon" viewBox="0 0 16 16" fill="currentColor">
                <path d="M8 3a5 5 0 1 0 4.546 2.914.5.5 0 0 1 .908-.417A6 6 0 1 1 8 2v1z"/>
                <path d="M8 4.466V.534a.25.25 0 0 1 .41-.192l2.36 1.966c.12.1.12.284 0 .384L8.41 4.658A.25.25 0 0 1 8 4.466z"/>
              </svg>
              <span class="btn-label">{{ $t('replace') }}</span>
            </button>
            <button class="chat-action-btn append-btn" @click="insertFromMessage(idx, 'append')" :disabled="loading" :title="$t('append')" aria-label="Thêm vào">
              <svg class="action-icon" viewBox="0 0 16 16" fill="currentColor">
                <path d="M8 2a.5.5 0 0 1 .5.5v5h5a.5.5 0 0 1 0 1h-5v5a.5.5 0 0 1-1 0v-5h-5a.5.5 0 0 1 0-1h5v-5A.5.5 0 0 1 8 2z"/>
              </svg>
              <span class="btn-label">{{ $t('append') }}</span>
            </button>
            <button
              class="chat-action-btn undo-btn"
              @click="performUndo"
              :disabled="loading || !(undoState.enabled && undoState.msgIdx === idx)"
              :title="'Hoàn tác'"
              aria-label="Hoàn tác"
            >
              <svg class="action-icon" viewBox="0 0 16 16" fill="currentColor">
                <path d="M3.5 1a.5.5 0 0 1 .5.5V4h8a4 4 0 1 1 0 8H8a.5.5 0 0 1 0-1h4a3 3 0 1 0 0-6H4v2.5a.5.5 0 0 1-1 0v-6A.5.5 0 0 1 3.5 1z"/>
              </svg>
              <span class="btn-label">Hoàn tác</span>
            </button>
          </div>
        </div>
      </template>
        <div v-if="historyDialog.length === 0" class="chat-empty">Chưa có tin nhắn. Hãy nhập yêu cầu ở dưới.</div>
      </div>
      <!-- Composer tối giản: chỉ còn ô nhập -->
      <div class="composer">
        <div class="composer-container">
          <button class="upload-btn" @click="triggerFileUpload" :disabled="loading" title="Upload PDF">
            <svg class="upload-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
              <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
              <polyline points="14,2 14,8 20,8"/>
              <line x1="16" y1="13" x2="8" y2="13"/>
              <line x1="16" y1="17" x2="8" y2="17"/>
              <line x1="10" y1="9" x2="8" y2="9"/>
            </svg>
          </button>
          <input
            v-model="directInput"
            class="composer-input"
            placeholder="Nhập yêu cầu (bot sẽ trả lời trong chat)"
            @keydown.enter.prevent="sendDirectInput"
          />
          <button class="send-btn" @click="sendDirectInput" :disabled="loading || !directInput.trim()" title="Gửi tin nhắn">
            <svg class="send-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
              <line x1="22" y1="2" x2="11" y2="13"/>
              <polygon points="22,2 15,22 11,13 2,9 22,2"/>
            </svg>
          </button>
          <input
            ref="fileInput"
            type="file"
            accept=".pdf"
            @change="handleFileUpload"
            style="display: none;"
          />
        </div>
      </div>
      <!-- Hiển thị số từ đã chọn khi người dùng bôi đen -->
      <!-- <div class="selection-count" v-if="selectedWordCount > 0">Đã chọn: {{ selectedWordCount }} từ</div> -->
      <!-- Hàng nút nhanh cho AI tools (4 nút nhỏ) -->
      <div class="quick-tools-row">
        <button class="quick-tool-btn btn-translate" @click="performAction('translate')" :disabled="loading">{{ $t('translate') }}</button>
        <button class="quick-tool-btn btn-summary" @click="performAction('summary')" :disabled="loading">{{ $t('summary') }}</button>
        <!-- <button class="quick-tool-btn btn-polish" @click="performAction('polish')" :disabled="loading">{{ $t('polish') }}</button> -->
        <button class="quick-tool-btn btn-grammar" @click="performAction('grammar')" :disabled="loading">{{ $t('grammar') }}</button>
        <!-- Nút Hoàn tác đã được chuyển vào mỗi ô chat, gỡ khỏi quick-tools -->
        <!-- <button class="quick-tool-btn full-width btn-format" @click="openFormatDialog = true" :disabled="loading">Định dạng văn bản</button> -->
      </div>

      <!-- Ẩn phần hiển thị kết quả AI theo yêu cầu -->
    </div>

    <!-- Dialogs -->
    <el-dialog v-model="openFormatDialog" width="90vw" class="format-dialog" title="Định dạng Word">
      <div style="display:flex; flex-direction:column; gap:8px;">
        <el-checkbox v-model="formatOptions.addHeader">Thêm tiêu ngữ</el-checkbox>
        <el-checkbox v-model="formatOptions.a4Size">Thiết lập A4 + Times New Roman 14pt</el-checkbox>
        <el-checkbox v-model="formatOptions.justifyMargins">Căn đều (3 -2 -2 - 2 cm)</el-checkbox>
        <el-checkbox v-model="formatOptions.lineSpacing">Giãn dòng ~1.15, bỏ khoảng trước/sau</el-checkbox>
      </div>
      <template #footer>
        <el-button @click="openFormatDialog = false">Đóng</el-button>
        <el-button type="primary" @click="applyFormatOptions; openFormatDialog = false">Áp dụng</el-button>
      </template>
    </el-dialog>
  </div>
</template>

<script lang="ts" setup>
import { ElMessage } from 'element-plus'
import { Settings, Loader2, Zap, ChevronDown } from 'lucide-vue-next'
import { User } from 'lucide-vue-next'
import { onBeforeMount, ref, watch, computed } from 'vue'
import { useRouter } from 'vue-router'
import { useI18n } from 'vue-i18n'

import API from '@/api'
// Word info feature removed per request

import { buildInPrompt } from '@/utils/constant'
import { promptDbInstance } from '@/store/promtStore'

import { checkAuth } from '@/utils/common'
import { localStorageKey } from '@/utils/enum'
import useSettingForm from '@/utils/settingForm'
import { settingPreset } from '@/utils/settingPreset'

import HomePageDialog from '@/components/HomePageDialog.vue'
import HomePageAddDialog from '@/components/HomePageAddDialog.vue'

const SYSTEM_PROMPT = ` 
 Bạn là một trợ lý AI điều khiển Microsoft Word thông qua Office.js (Word JavaScript API). 
 
 Nhiệm vụ của bạn: 
 1️⃣ Đọc yêu cầu người dùng. 
 2️⃣ Phân loại yêu cầu thành **hành động Word** (thay, định dạng, chèn, truy vấn...) hoặc **trao đổi ngôn ngữ tự nhiên** (dịch, tóm tắt, hỏi đáp...). 
 3️⃣ Trả về JSON duy nhất, KHÔNG có văn bản ngoài JSON. 
 
 --- 
 
 ### Cấu trúc phản hồi: 
 Nếu là hành động Word: 
 \`\`\`json 
 { 
   "action": "<loại hành động>", 
   "parameters": { ... } 
 } 
 \`\`\` 
 
 Nếu là trò chuyện / dịch / hỏi đáp: 
 \`\`\`json 
 { 
   "action": "chat", 
   "response": "<câu trả lời hoặc nội dung văn bản>" 
 } 
 \`\`\` 
 
 --- 
 
 ### Các loại hành động hợp lệ thương thích Office.js: 
 
 #### 1. "replace" – thay thế văn bản 
 Tìm và thay thế văn bản trong toàn tài liệu. 
 \`\`\`json 
 { 
   "action": "replace", 
   "parameters": { "find": "cũ", "replace": "mới", "matchCase": false } 
 } 
 \`\`\` 
 
 #### 2. "format" – định dạng đoạn chọn hoặc toàn văn bản 
 Các thuộc tính cho phép: 
 - fontName 
 - fontSize 
 - bold 
 - italic 
 - underline 
 - color 
 - alignment (left, center, right, justify) 
 - lineSpacing 
 - paragraphSpacingBefore / paragraphSpacingAfter 
 \`\`\`json 
 { 
   "action": "format", 
   "parameters": { "fontName": "Times New Roman", "fontSize": 14, "alignment": "center" } 
 } 
 \`\`\` 
 
 #### 3. "insert" – chèn văn bản 
 \`\`\`json 
 { 
   "action": "insert", 
   "parameters": { "text": "Xin chào", "position": "end" } 
 } 
 \`\`\` 
 position có thể là: "start", "end", "beforeSelection", "afterSelection" 
 
 #### 4. "query" – truy vấn thông tin tài liệu 
 Cho phép các loại: 
 - pageSize (A4, Letter, v.v.) 
 - orientation (portrait/landscape) 
 - fontInfo (font và cỡ chữ đang chọn) 
 - wordCount 
 \`\`\`json 
 { 
   "action": "query", 
   "parameters": { "queryType": "pageSize" } 
 } 
 \`\`\` 
 
 #### 5. "highlight" – tô màu nền cho đoạn văn hoặc từ khóa 
 \`\`\`json 
 { 
   "action": "highlight", 
   "parameters": { "find": "quan trọng", "color": "yellow" } 
 } 
 \`\`\` 
 
 #### 6. "comment" – thêm nhận xét vào đoạn chọn 
 \`\`\`json 
 { 
   "action": "comment", 
   "parameters": { "text": "Cần xem lại đoạn này" } 
 } 
 \`\`\` 
 
 #### 7. "chat" – xử lý tự nhiên 
 Ví dụ: dịch, tóm tắt, phân tích, hỏi đáp... 
 \`\`\`json 
 { 
   "action": "chat", 
   "response": "<nội dung>" 
 } 
 \`\`\` 
 
 --- 
 
 ### Quy tắc bắt buộc: 
 - Không tạo định dạng hoặc thao tác Word mà Office.js không hỗ trợ. 
 - Nếu người dùng yêu cầu một tác vụ vượt quá giới hạn (ví dụ: “tạo biểu đồ nâng cao” hoặc “vẽ đường viền tùy chỉnh”), hãy phản hồi bằng "action": "chat" và mô tả kết quả thay thế. 
 - Luôn đảm bảo JSON hợp lệ, không có văn bản ngoài. 
 - Nếu có nhiều thao tác cần thiết, có thể trả về: 
 \`\`\`json 
 { "actions": [ { ... }, { ... } ] } 
 \`\`\` 
 
 --- 
 
 ### Ví dụ: 
 - "Đặt font Arial, size 12, căn giữa": 
 → {"action":"format","parameters":{"fontName":"Arial","fontSize":12,"alignment":"center"}} 
 
 - "Thay 'Hà Nội' bằng 'TP.HCM'": 
 → {"action":"replace","parameters":{"find":"Hà Nội","replace":"TP.HCM"}} 
 
 - "Khổ giấy hiện tại là gì?": 
 → {"action":"query","parameters":{"queryType":"pageSize"}} 
 
 - "Dịch đoạn này sang tiếng Anh": 
 → {"action":"chat","response":"<văn bản dịch>"} 
 
 - "Tô vàng từ 'quan trọng'": 
 → {"action":"highlight","parameters":{"find":"quan trọng","color":"yellow"}} 
 `

const { t } = useI18n()

const { settingForm } = useSettingForm()

// Username hiển thị cạnh icon User
const username = ref('')
onBeforeMount(() => {
  const saved = typeof localStorage !== 'undefined' ? localStorage.getItem('authUsername') : ''
  username.value = saved || (import.meta.env?.DEV ? 'demo' : 'User')
})

// Collapsible section states
const aiToolsExpanded = ref(true)
const promptsExpanded = ref(false)

// system prompt
const systemPrompt = ref('')
const systemPromptSelected = ref('')
const systemPromptList = ref<IStringKeyMap[]>([])
const addSystemPromptVisible = ref(false)
const addSystemPromptAlias = ref('')
const addSystemPromptValue = ref('')
const removeSystemPromptVisible = ref(false)
const removeSystemPromptValue = ref<any[]>([])

// user prompt
const prompt = ref('')
const promptSelected = ref('')
const promptList = ref<IStringKeyMap[]>([])
const addPromptVisible = ref(false)
const addPromptAlias = ref('')
const addPromptValue = ref('')
const removePromptVisible = ref(false)
const removePromptValue = ref<any[]>([])

// Direct prompt input
const directInput = ref('')
// File upload functionality
const fileInput = ref<HTMLInputElement>()

// Composer states (đã bỏ nút + và micro)
// Bỏ Quick response theo yêu cầu

// File upload functions
function triggerFileUpload() {
  fileInput.value?.click()
}

function handleFileUpload(event: Event) {
  const target = event.target as HTMLInputElement
  const file = target.files?.[0]
  
  if (!file) return
  
  if (file.type !== 'application/pdf') {
    ElMessage.warning('Chỉ hỗ trợ file PDF')
    return
  }
  
  // TODO: Implement PDF processing logic here
  ElMessage.success(`Đã chọn file: ${file.name}`)
  console.log('Selected PDF file:', file)
  
  // Reset input
  target.value = ''
}

// Đếm số từ trong vùng văn bản đang bôi đen của Word
const selectedWordCount = ref(0)
async function updateSelectionCount() {
  try {
    await Word.run(async (ctx) => {
      const sel = ctx.document.getSelection()
      sel.load('text')
      await ctx.sync()
      const text = String(sel.text || '').trim()
      selectedWordCount.value = text ? (text.split(/\s+/).filter(Boolean).length) : 0
    })
  } catch (e) {
    // Không có vùng chọn hoặc Office.js chưa sẵn sàng
    selectedWordCount.value = 0
  }
}
function registerSelectionHandler() {
  try {
    // Cập nhật ngay khi vào, và đăng ký lắng nghe sự kiện thay đổi vùng chọn
    updateSelectionCount()
    Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, () => {
      updateSelectionCount()
    })
  } catch (e) {
    console.warn('Không thể đăng ký handler DocumentSelectionChanged:', e)
  }
}

// Cập nhật/loại bỏ thông báo nhỏ trong khung chat khi người dùng bôi đen văn bản
function updateSelectionNotice(newCount?: number) {
  const count = typeof newCount === 'number' ? newCount : selectedWordCount.value
  if (count > 0) {
    const message: any = {
      role: 'assistant',
      content: `Đã chọn: ${count} từ`,
      meta: 'Người dùng',
      notice: true
    }
    if (selectionNoticeIndex.value !== null && historyDialog.value[selectionNoticeIndex.value!]) {
      historyDialog.value[selectionNoticeIndex.value!] = message
    } else {
      historyDialog.value.push(message)
      selectionNoticeIndex.value = historyDialog.value.length - 1
    }
  } else {
    if (selectionNoticeIndex.value !== null) {
      historyDialog.value.splice(selectionNoticeIndex.value!, 1)
      selectionNoticeIndex.value = null
    }
  }
}

// Theo dõi thay đổi selectedWordCount để cập nhật thông báo nhỏ
watch(selectedWordCount, (n) => {
  updateSelectionNotice(n)
})

// Xây dựng danh sách messages gửi cho model (bỏ qua các notice)
function buildMessagesForModel() {
  return historyDialog.value
    .filter((m: any) => !m.notice)
    .map((m: any) => ({ role: m.role, content: (m.hiddenContent ?? m.content) }))
}

// result
const result = ref('res')
const displayResult = computed(() => {
  const text = String(result.value || '')
  if (text.length <= 200) return text
  return text.slice(0, 200) + '...'
})
// Word info removed
const loading = ref(false)
const router = useRouter()
const historyDialog = ref<any[]>([])
// Vị trí thông báo nhỏ (notice) trong mảng historyDialog
const selectionNoticeIndex = ref<number | null>(null)

const jsonIssue = ref(false)
const errorIssue = ref(false)

// Lưu system prompt gần nhất để gửi kèm khi gọi API (không hiển thị lên UI)
const lastSystemPrompt = ref<string>('')

// handleGetWordInfo removed

// insert type
const insertType = ref<insertTypes>('replace')
const useWordFormatting = ref(true)
// Tuỳ chọn định dạng Word & hộp thoại
const formatOptions = ref({
  addHeader: false,
  a4Size: true,
  justifyMargins: true,
  lineSpacing: true
})
// Hộp thoại popup cho tuỳ chọn định dạng
const openFormatDialog = ref(false)
// Trạng thái Undo cho hành động Thay thế/Thêm
const undoState = ref<{ enabled: boolean; ccId: number | null; type: 'replace' | 'append' | null; originalText: string | null; msgIdx: number | null }>({
  enabled: false,
  ccId: null,
  type: null,
  originalText: null,
  msgIdx: null
})

// Dynamic model selection based on current API
// Reactive options for Ollama models
const ollamaModelOptions = ref<Array<{ label: string; value: string }>>(settingPreset.ollamaModelSelect.optionList || [])

const currentModelOptions = computed(() => ollamaModelOptions.value)



const currentModelSelect = computed({
  get() {
    return settingForm.value.ollamaModelSelect
  },
  set(value) {
    settingForm.value.ollamaModelSelect = value
  }
})



const addWatch = () => {
  watch(
    () => settingForm.value.replyLanguage,
    () => {
      localStorage.setItem('replyLanguage', settingForm.value.replyLanguage)
    }
  )
}

async function initData() {
  insertType.value =
    (localStorage.getItem(localStorageKey.insertType) as insertTypes) ||
    'replace'
  useWordFormatting.value =
    localStorage.getItem(localStorageKey.useWordFormatting) === 'true'
  systemPrompt.value =
    localStorage.getItem(localStorageKey.defaultSystemPrompt) ||
    'Act like a personal assistant.'
  prompt.value = localStorage.getItem(localStorageKey.defaultPrompt) || ''
  lastSystemPrompt.value = SYSTEM_PROMPT
}

function handleInsertTypeChange(val: insertTypes) {
  insertType.value = val
  localStorage.setItem(localStorageKey.insertType, val)
}

function handleWordFormattingChange(val: boolean) {
  useWordFormatting.value = val
  localStorage.setItem(localStorageKey.useWordFormatting, String(val))
}

// Load Ollama models dynamically when Ollama API is selected
async function loadOllamaModels() {
  try {
    const endpoint = (settingForm.value.ollamaEndpoint || '').trim()
    if (!endpoint) {
      ElMessage.warning('Vui lòng cấu hình Ollama endpoint trong Settings')
      return
    }
    const models = await API.ollama.listModels(endpoint)
    if (models && models.length > 0) {
      // Update reactive options for HomePage
      ollamaModelOptions.value = models
      // Also update preset for SettingsPage usage (non-reactive there)
      settingPreset.ollamaModelSelect.optionList = models
      // Ensure current selection is valid
      const current = settingForm.value.ollamaModelSelect
      const exists = models.some(m => m.value === current)
      if (!exists) {
        settingForm.value.ollamaModelSelect = models[0].value
      }
      // Không hiển thị thông báo thành công để tránh gây nhiễu
    } else {
      ElMessage.warning('Không tìm thấy model từ endpoint Ollama')
    }
  } catch (e) {
    console.error('Load Ollama models failed:', e)
    ElMessage.warning('Nạp models thất bại. Kiểm tra kết nối/endpoint')
  }
}

// Load Ollama models on mount (app is Ollama-only)

// Tạo nội dung hiển thị rút gọn cho khung chat (tiếng Việt)
function buildDisplayMessage(
  taskType: keyof typeof buildInPrompt | 'custom',
  selectedText: string
) {
  const labels: Record<string, string> = {
    translate: 'Dịch đoạn văn bản',
    summary: 'Tóm tắt đoạn văn bản',
    grammar: 'Sửa ngữ pháp đoạn văn bản',
    polish: 'Chỉnh sửa đoạn văn bản',
    custom: 'Yêu cầu'
  }
  const base = labels[String(taskType)] || 'Yêu cầu'
  const preview = String(selectedText || '')
    .trim()
    .replace(/\s+/g, ' ')
  const maxLen = 140
  const shortened = preview.length > maxLen ? preview.slice(0, maxLen) + '…' : preview
  return preview ? `${base}: ${shortened}` : base
}

async function template(taskType: keyof typeof buildInPrompt | 'custom') {
  loading.value = true
  let systemMessage
  let userMessage = ''
  const getSeletedText = async () => {
    return Word.run(async context => {
      const range = context.document.getSelection()
      range.load('text')
      await context.sync()
      return range.text
    })
  }
  const selectedText = await getSeletedText()
  if (taskType === 'custom') {
    if (systemPrompt.value.includes('{language}')) {
      systemMessage = systemPrompt.value.replace(
        '{language}',
        settingForm.value.replyLanguage
      )
    } else {
      systemMessage = systemPrompt.value
    }
    if (userMessage.includes('{text}')) {
      userMessage = userMessage.replace('{text}', selectedText)
    } else {
      userMessage = `Reply in ${settingForm.value.replyLanguage} ${prompt.value} ${selectedText}`
    }
  } else {
    systemMessage = SYSTEM_PROMPT
    userMessage = buildInPrompt[taskType].user(
      selectedText,
      settingForm.value.replyLanguage
    )
  }
  if (settingForm.value.ollamaEndpoint) {
    // Ghi nhớ system prompt cho các lần gọi tiếp theo
    lastSystemPrompt.value = systemMessage

    // Hiển thị nội dung rút gọn lên UI và giữ nội dung đầy đủ cho model
    const displayMessage = buildDisplayMessage(taskType, selectedText)
    const userPayload = {
      role: 'user',
      content: displayMessage,
      hiddenContent: userMessage
    }
    historyDialog.value.push(userPayload)
    // Thêm thông báo đã nhận bao nhiêu từ (nếu có vùng chọn)
    const countSelected = String(selectedText || '').trim()
      ? String(selectedText).trim().split(/\s+/).filter(Boolean).length
      : 0
    // Cập nhật thông báo lựa chọn (duy trì 1 notice duy nhất)
    updateSelectionNotice(countSelected)
    // Xây dựng payload gửi tới model: chèn system prompt ở đầu, sau đó là lịch sử chat (đã lọc notice)
    const messagesForModel = [
      { role: 'system', content: lastSystemPrompt.value },
      ...buildMessagesForModel()
    ]

    await API.ollama.createChatCompletionStream({
      ollamaEndpoint: settingForm.value.ollamaEndpoint,
      ollamaModel:
        settingForm.value.ollamaCustomModel ||
        settingForm.value.ollamaModelSelect,
      messages: messagesForModel,
      result,
      historyDialog,
      errorIssue,
      loading,
      temperature: settingForm.value.ollamaTemperature
    })
    await handleApiResponseWithPrompt(String(result.value || ''))
  } else {
    ElMessage.error('Set Ollama endpoint first')
    return
  }
  if (errorIssue.value === true) {
    errorIssue.value = false
    ElMessage.error('Something is wrong')
    return
  }
  
}

function checkApiKey() {
  const auth = {
    type: 'ollama',
    ollamaEndpoint: settingForm.value.ollamaEndpoint
  } as any
  if (!checkAuth(auth)) {
    ElMessage.error('Set Olla endpoint first')
    return false
  }
  return true
}

async function performAction(action: keyof typeof buildInPrompt) {
  if (!checkApiKey()) return
  template(action)
}

function settings() {
  router.push('/settings')
}

function userPage() {
  router.push('/user')
}

function StartChat() {
  if (!checkApiKey()) return
  template('custom')
}

async function continueChat() {
  if (!checkApiKey()) return
  loading.value = true
  try {
    historyDialog.value.push({
      role: 'user',
      content: 'continue'
    })
    // Nếu có system prompt trước đó, thêm vào đầu payload gửi model
    const messagesForModel = lastSystemPrompt.value
      ? [{ role: 'system', content: lastSystemPrompt.value }, ...buildMessagesForModel()]
      : buildMessagesForModel()

    await API.ollama.createChatCompletionStream({
      ollamaEndpoint: settingForm.value.ollamaEndpoint,
      ollamaModel:
        settingForm.value.ollamaCustomModel ||
        settingForm.value.ollamaModelSelect,
      messages: messagesForModel,
      result,
      historyDialog,
      errorIssue,
      loading,
      temperature: settingForm.value.ollamaTemperature
    })
    await handleApiResponseWithPrompt(String(result.value || ''))
  } catch (error) {
    result.value = String(error)
    errorIssue.value = true
    console.error(error)
  }
  if (errorIssue.value === true) {
    errorIssue.value = false
    ElMessage.error('Something is wrong')
    return
  }
  
}

// Send direct prompt to AI API
async function sendDirectInput() {
  if (!checkApiKey()) return
  // Lấy nội dung gõ trực tiếp và ghép thêm văn bản đang được bôi đen (nếu có)
  let content = directInput.value.trim()
  let selectedText = ''
  try {
    await Word.run(async (ctx) => {
      const sel = ctx.document.getSelection()
      sel.load('text')
      await ctx.sync()
      selectedText = (sel.text || '')
    })
    if (selectedText && selectedText.trim().length > 0) {
      content = `${content}\n\n${selectedText.trim()}`
    }
  } catch (err) {
    console.warn('Không thể lấy vùng chọn từ Word:', err)
  }
  if (!content) {
    ElMessage.warning('Vui lòng nhập yêu cầu')
    return
  }
  loading.value = true
  try {
    historyDialog.value.push({
      role: 'user',
      content
    })
    // Clean input field after sending
    directInput.value = ''
    // Thêm thông báo đã nhận bao nhiêu từ (chỉ tính phần văn bản bôi đen)
    const countSelectedOnly = selectedText && selectedText.trim().length > 0
      ? selectedText.trim().split(/\s+/).filter(Boolean).length
      : 0
    // Cập nhật thông báo lựa chọn (duy trì 1 notice duy nhất)
    updateSelectionNotice(countSelectedOnly)
    await API.ollama.createChatCompletionStream({
      ollamaEndpoint: settingForm.value.ollamaEndpoint,
      ollamaModel:
        settingForm.value.ollamaCustomModel ||
        settingForm.value.ollamaModelSelect,
      messages: (
        lastSystemPrompt.value
          ? [{ role: 'system', content: lastSystemPrompt.value }, ...buildMessagesForModel()]
          : buildMessagesForModel()
      ),
      result,
      historyDialog,
      errorIssue,
      loading,
      temperature: settingForm.value.ollamaTemperature
    })
    await handleApiResponseWithPrompt(String(result.value || ''))
  } catch (error) {
    result.value = String(error)
    errorIssue.value = true
    console.error(error)
  }
  if (errorIssue.value === true) {
    errorIssue.value = false
    ElMessage.error('Something is wrong')
    return
  }
  
}

// Manual insert from a specific assistant message
async function insertFromMessage(index: number, type: insertTypes) {
  try {
    const msg = historyDialog.value[index]
    if (!msg || msg.role !== 'assistant') {
      ElMessage.warning('Chỉ chèn nội dung từ trả lời của Bot')
      return
    }
    const text = String(msg.content || '').trim()
    if (!text) {
      ElMessage.warning('Nội dung trống')
      return
    }
    const tempResult = ref(text)
    const tempInsertType = ref<insertTypes>(type)

    if (type === 'replace') {
      await Word.run(async (ctx) => {
        const sel = ctx.document.getSelection()
        sel.load('text')
        await ctx.sync()
        const originalText = sel.text || ''
        const cc = sel.insertContentControl()
        cc.appearance = 'Hidden'
        cc.load('id')
        await ctx.sync()
        cc.insertText(tempResult.value, 'Replace')
        await ctx.sync()
        undoState.value = {
          enabled: true,
          ccId: cc.id,
          type: 'replace',
          originalText,
          msgIdx: index
        }
      })
    } else if (type === 'append') {
      await Word.run(async (ctx) => {
        const sel = ctx.document.getSelection()
        const para = sel.insertParagraph(tempResult.value, 'After')
        const cc = para.insertContentControl()
        cc.appearance = 'Hidden'
        cc.load('id')
        await ctx.sync()
        undoState.value = {
          enabled: true,
          ccId: cc.id,
          type: 'append',
          originalText: null,
          msgIdx: index
        }
      })
    } else {
      if (useWordFormatting.value) {
        API.common.insertFormattedResult(tempResult, tempInsertType)
      } else {
        API.common.insertResult(tempResult, tempInsertType)
      }
    }
  } catch (e) {
    console.error(e)
  }
}

async function performUndo() {
  try {
    if (!undoState.value.enabled || !undoState.value.ccId) return
    await Word.run(async (ctx) => {
      const cc = ctx.document.contentControls.getByIdOrNullObject(Number(undoState.value.ccId))
      cc.load('isNullObject')
      await ctx.sync()
      if (cc.isNullObject) {
        ElMessage.warning('Không tìm thấy vùng cần hoàn tác')
        undoState.value.enabled = false
        return
      }
      if (undoState.value.type === 'replace') {
        cc.insertText(String(undoState.value.originalText || ''), 'Replace')
        // Xóa content control nhưng giữ lại nội dung đã khôi phục
        cc.delete(true)
      } else if (undoState.value.type === 'append') {
        // Xóa cả content control và nội dung được thêm
        cc.delete(false)
      }
      await ctx.sync()
    })
  } catch (e) {
    console.error('Hoàn tác thất bại:', e)
    ElMessage.error('Hoàn tác thất bại')
  } finally {
    undoState.value.enabled = false
    undoState.value.ccId = null
    undoState.value.type = null
    undoState.value.originalText = null
    undoState.value.msgIdx = null
  }
}

// JSON handler for system prompt actions
async function handleApiResponseWithPrompt(rawText: string) {
  const parsed = parseJsonFromText(rawText)
  if (!parsed) return
  const last = historyDialog.value[historyDialog.value.length - 1]
  if (last && last.role === 'assistant' && String(last.content) === String(rawText)) {
    historyDialog.value.pop()
  }
  const steps = Array.isArray((parsed as any).actions) ? (parsed as any).actions : [parsed]
  for (const s of steps) {
    await executeWordAction(s)
  }
}

function parseJsonFromText(text: string): any | null {
  try {
    const fence = text.match(/```json\s*([\s\S]*?)```/i) || text.match(/```\s*([\s\S]*?)```/i)
    if (fence && fence[1]) {
      return JSON.parse(fence[1])
    }
    const start = text.indexOf('{')
    const end = text.lastIndexOf('}')
    if (start >= 0 && end > start) {
      const candidate = text.slice(start, end + 1)
      return JSON.parse(candidate)
    }
  } catch (e) {
    console.warn('Không thể parse JSON từ phản hồi:', e)
  }
  return null
}

async function executeWordAction(actionObj: any) {
  try {
    const action = String(actionObj?.action || '').trim()
    const p = actionObj?.parameters || {}
    if (action === 'chat') {
      const responseText = String(actionObj?.response || p?.text || '')
      if (responseText) {
        historyDialog.value.push({ role: 'assistant', content: responseText })
      }
      return
    }
    if (action === 'insert') {
      const text = String(p?.text || '')
      const position = String(p?.position || 'afterSelection')
      await Word.run(async ctx => {
        let target = ctx.document.getSelection()
        if (position === 'start') target = ctx.document.body.getRange('Start')
        else if (position === 'end') target = ctx.document.body.getRange('End')
        else if (position === 'beforeSelection') target = ctx.document.getSelection().getRange('Before')
        else if (position === 'afterSelection') target = ctx.document.getSelection().getRange('After')
        target.insertText(text, 'After')
        await ctx.sync()
      })
      historyDialog.value.push({ role: 'assistant', content: 'Đã chèn văn bản' })
      return
    }
    if (action === 'replace') {
      const find = String(p?.find || '')
      const replace = String(p?.replace || '')
      const matchCase = !!p?.matchCase
      await Word.run(async ctx => {
        const results = ctx.document.body.search(find, { matchCase })
        results.load('items')
        await ctx.sync()
        for (const r of results.items) {
          r.insertText(replace, 'Replace')
        }
        await ctx.sync()
      })
      historyDialog.value.push({ role: 'assistant', content: 'Đã thay thế văn bản', noActions: true })
      return
    }
    if (action === 'format') {
      await Word.run(async ctx => {
        const sel = ctx.document.getSelection()
        if (p.fontName) sel.font.name = String(p.fontName)
        if (p.fontSize) sel.font.size = Number(p.fontSize)
        if (typeof p.bold === 'boolean') sel.font.bold = p.bold
        if (typeof p.italic === 'boolean') sel.font.italic = p.italic
        if (typeof p.underline === 'boolean') sel.font.underline = p.underline ? 'Single' : 'None'
        if (p.color) sel.font.color = String(p.color)
        // Use Office.js enum for alignment to avoid InvalidArgument
        const alignMap: Record<string, any> = {
          left: Word.Alignment.left,
          center: Word.Alignment.centered,
          right: Word.Alignment.right,
          justify: Word.Alignment.justified
        }
        if (p.alignment && alignMap[String(p.alignment)]) {
          const paras = sel.paragraphs
          paras.load('items')
          await ctx.sync()
          for (const para of paras.items) {
            para.alignment = alignMap[String(p.alignment)]
          }
        }
        await ctx.sync()
      })
      historyDialog.value.push({ role: 'assistant', content: 'Đã định dạng văn bản', noActions: true })
      return
    }
    if (action === 'highlight') {
      const find = String(p?.find || '')
      const color = String(p?.color || 'yellow')
      await Word.run(async ctx => {
        const results = ctx.document.body.search(find, { matchCase: !!p?.matchCase })
        results.load('items')
        await ctx.sync()
        for (const r of results.items) {
          r.font.highlightColor = color
        }
        await ctx.sync()
      })
      historyDialog.value.push({ role: 'assistant', content: 'Đã tô màu nổi bật', noActions: true })
      return
    }
    if (action === 'comment') {
      const text = String(p?.text || '')
      await Word.run(async ctx => {
        const sel = ctx.document.getSelection() as any
        if (sel && typeof sel.insertComment === 'function') {
          sel.insertComment(text)
        } else if ((ctx.document as any).comments && typeof (ctx.document as any).comments.add === 'function') {
          ;(ctx.document as any).comments.add(sel, text)
        } else {
          sel.insertText(`[Nhận xét] ${text}`, 'Before')
        }
        await ctx.sync()
      })
      historyDialog.value.push({ role: 'assistant', content: 'Đã thêm nhận xét' })
      return
    }
    if (action === 'query') {
      const qt = String(p?.queryType || '')
      if (qt === 'fontInfo') {
        await Word.run(async ctx => {
          const sel = ctx.document.getSelection()
          sel.font.load('name,size,bold,italic,color')
          await ctx.sync()
          const info = `Phông chữ: ${sel.font.name}, Cỡ: ${sel.font.size}pt, Đậm: ${sel.font.bold ? 'Có' : 'Không'}, Nghiêng: ${sel.font.italic ? 'Có' : 'Không'}`
          historyDialog.value.push({ role: 'assistant', content: info, noActions: true })
        })
        return
      }
      if (qt === 'wordCount') {
        await Word.run(async ctx => {
          const body = ctx.document.body
          const rng = body.getRange('Start')
          const end = body.getRange('End')
          const total = rng.expandTo(end)
          total.load('text')
          await ctx.sync()
          const count = String(total.text || '').trim().split(/\s+/).filter(Boolean).length
          historyDialog.value.push({ role: 'assistant', content: `Số chữ: ${count}`, noActions: true })
        })
        return
      }
      if (qt === 'pageSize' || qt === 'orientation') {
        await Word.run(async ctx => {
          const ooxml = ctx.document.body.getOoxml()
          await ctx.sync()
          const xml = ooxml.value || ''
          const pgSz = xml.match(/<w:pgSz[^>]*w:w="(\d+)"[^>]*w:h="(\d+)"[^>]*?(w:orient="([^"]+)")?/)
          if (pgSz) {
            const wTw = parseInt(pgSz[1], 10)
            const hTw = parseInt(pgSz[2], 10)
            const orient = pgSz[4] || 'portrait'
            const cm = (tw: number) => Math.round((tw / 566.929) * 100) / 100
            const info = `Khổ giấy: ${cm(wTw)}cm x ${cm(hTw)}cm, Hướng: ${orient}`
            historyDialog.value.push({ role: 'assistant', content: info, noActions: true })
          } else {
            historyDialog.value.push({ role: 'assistant', content: 'Không lấy được khổ giấy', noActions: true })
          }
        })
        return
      }
      historyDialog.value.push({ role: 'assistant', content: 'Không hỗ trợ loại truy vấn này', noActions: true })
      return
    }
    historyDialog.value.push({ role: 'assistant', content: `Hành động chưa hỗ trợ: ${action}` })
  } catch (e) {
    console.error('Thực thi hành động thất bại:', e)
    ElMessage.error('Thực thi hành động thất bại')
  }
}
onBeforeMount(() => {
  addWatch()
  initData()
  loadOllamaModels()
  registerSelectionHandler()
})
</script>

<style scoped src="./HomePage.css"></style>
