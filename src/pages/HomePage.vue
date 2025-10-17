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
          <div v-if="msg.role === 'assistant' && !msg.notice" class="chat-actions internal">
            <button class="chat-action-btn replace-btn" @click="insertFromMessage(idx, 'replace')" :disabled="loading">
              <svg class="action-icon" viewBox="0 0 16 16" fill="currentColor">
                <path d="M8 3a5 5 0 1 0 4.546 2.914.5.5 0 0 1 .908-.417A6 6 0 1 1 8 2v1z"/>
                <path d="M8 4.466V.534a.25.25 0 0 1 .41-.192l2.36 1.966c.12.1.12.284 0 .384L8.41 4.658A.25.25 0 0 1 8 4.466z"/>
              </svg>
              {{ $t('replace') }}
            </button>
            <button class="chat-action-btn append-btn" @click="insertFromMessage(idx, 'append')" :disabled="loading">
              <svg class="action-icon" viewBox="0 0 16 16" fill="currentColor">
                <path d="M8 2a.5.5 0 0 1 .5.5v5h5a.5.5 0 0 1 0 1h-5v5a.5.5 0 0 1-1 0v-5h-5a.5.5 0 0 1 0-1h5v-5A.5.5 0 0 1 8 2z"/>
              </svg>
              {{ $t('append') }}
            </button>
          </div>
        </div>
      </template>
        <div v-if="historyDialog.length === 0" class="chat-empty">Chưa có tin nhắn. Hãy nhập yêu cầu ở dưới.</div>
      </div>
      <!-- Composer tối giản: chỉ còn ô nhập -->
      <div class="composer">
        <input
          v-model="directInput"
          class="composer-input"
          placeholder="Nhập yêu cầu (bot sẽ trả lời trong chat)"
          @keydown.enter.prevent="sendDirectInput"
        />
      </div>
      <!-- Hiển thị số từ đã chọn khi người dùng bôi đen -->
      <!-- <div class="selection-count" v-if="selectedWordCount > 0">Đã chọn: {{ selectedWordCount }} từ</div> -->
      <!-- Hàng nút nhanh cho AI tools (4 nút nhỏ) -->
      <div class="quick-tools-row">
        <button class="quick-tool-btn btn-translate" @click="performAction('translate')" :disabled="loading">{{ $t('translate') }}</button>
        <button class="quick-tool-btn btn-summary" @click="performAction('summary')" :disabled="loading">{{ $t('summary') }}</button>
        <button class="quick-tool-btn btn-polish" @click="performAction('polish')" :disabled="loading">{{ $t('polish') }}</button>
        <button class="quick-tool-btn btn-grammar" @click="performAction('grammar')" :disabled="loading">{{ $t('grammar') }}</button>
        <button class="quick-tool-btn full-width btn-format" @click="openFormatDialog = true" :disabled="loading">Định dạng văn bản</button>
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
        <el-button type="primary" @click="applyFormatOptions(); openFormatDialog = false">Áp dụng</el-button>
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
// Composer states (đã bỏ nút + và micro)
// Bỏ Quick response theo yêu cầu

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
    systemMessage = buildInPrompt[taskType].system(
      settingForm.value.replyLanguage
    )
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
    ElMessage.error('Set Ollama endpoint first')
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
function insertFromMessage(index: number, type: insertTypes) {
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
    if (useWordFormatting.value) {
      API.common.insertFormattedResult(tempResult, tempInsertType)
    } else {
      API.common.insertResult(tempResult, tempInsertType)
    }
  } catch (e) {
    errorIssue.value = true
    ElMessage.error('Không thể chèn nội dung')
    console.error(e)
  }
}

// Quick response removed per request

// Suggestions under greeting
function handleQuickSuggest(type: 'summary' | 'expand') {
  if (type === 'summary') {
    directInput.value = 'Create a summary of this page.'
  } else {
    directInput.value = 'Expand on this topic.'
  }
  sendDirectInput()
}

// Placeholder for voice input
// Bỏ tính năng voice input theo yêu cầu

// (Hàm handleGetWordInfo phục vụ collapse đã khai báo phía trên)

// Áp dụng định dạng theo tuỳ chọn hộp thoại
async function applyFormatOptions() {
  try {
    loading.value = true
    await Word.run(async (context) => {
      const body = context.document.body
      const cmToPoints = (cm: number) => Math.round(cm * 28.346)

      // Tuỳ chọn: Thêm tiêu đề
      if (formatOptions.value.addHeader) {
        const table = body.insertTable(2, 2, Word.InsertLocation.start, [
          [
            'UBND TỈNH HÀ TĨNH\nSỞ VĂN HÓA, THỂ THAO VÀ DU LỊCH',
            'CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM\nĐộc lập - Tự do - Hạnh phúc',
          ],
          [
            'Số: 2044 / SVHTTDL-VP\nVề việc hướng dẫn thực hiện công tác phổ biến, giáo dục pháp luật quý III năm 2025',
            'Hà Tĩnh, ngày  tháng  năm 20',
          ],
        ])
        table.alignment = 'Centered'
        const rows = table.rows
        rows.load('items')
        await context.sync()
        const [row1, row2] = rows.items
        row1.cells.load('items')
        row2.cells.load('items')
        await context.sync()
        const [left1, right1] = row1.cells.items
        const [left2, right2] = row2.cells.items

        left1.body.paragraphs.load('items')
        right1.body.paragraphs.load('items')
        left2.body.paragraphs.load('items')
        right2.body.paragraphs.load('items')
        await context.sync()

        for (const p of left1.body.paragraphs.items) {
          p.font.name = 'Times New Roman'
          p.font.size = 14
          p.font.bold = true
          p.alignment = 'Centered'
        }
        for (const p of right1.body.paragraphs.items) {
          p.font.name = 'Times New Roman'
          p.font.size = 14
          p.font.bold = true
          p.alignment = 'Centered'
        }
        const parasRight1 = right1.body.paragraphs
        parasRight1.load('items')
        await context.sync()
        if (parasRight1.items.length > 1) {
          parasRight1.items[1].font.underline = 'Single'
        }
        for (const p of left2.body.paragraphs.items) {
          p.font.name = 'Times New Roman'
          p.font.size = 14
          p.alignment = 'Left'
        }
        const parasLeft2 = left2.body.paragraphs
        parasLeft2.load('items')
        await context.sync()
        if (parasLeft2.items.length > 0) {
          parasLeft2.items[0].font.bold = true
        }
        if (parasLeft2.items.length > 1) {
          parasLeft2.items[1].font.italic = true
        }
        for (const p of right2.body.paragraphs.items) {
          p.font.name = 'Times New Roman'
          p.font.size = 14
          p.font.italic = true
          p.alignment = 'Centered'
        }
      }

      // Bỏ qua thiết lập khổ giấy A4 trực tiếp: Office.js không hỗ trợ
      // Tuỳ chọn: Thiết lập A4 (áp dụng font mặc định Times New Roman 14pt)
      if (formatOptions.value.a4Size) {
        const whole = body.getRange()
        whole.font.name = 'Times New Roman'
        whole.font.size = 14
      }

      // Tuỳ chọn: Căn lề & thụt lề, Tuỳ chọn: Giãn dòng
      const paragraphs = body.paragraphs
      paragraphs.load('items')
      await context.sync()
      const parentTables = paragraphs.items.map((p) => (p as any).parentTableOrNullObject)
      await context.sync()

      for (let i = 0; i < paragraphs.items.length; i++) {
        const p = paragraphs.items[i]
        const parentTable = parentTables[i]
        if (parentTable && !parentTable.isNullObject) continue // bỏ qua tiêu ngữ

        if (formatOptions.value.justifyMargins) {
          p.alignment = 'Justified'
          p.leftIndent = cmToPoints(2)
          p.rightIndent = cmToPoints(3)
          p.firstLineIndent = cmToPoints(1)
        }
        if (formatOptions.value.lineSpacing) {
          p.lineSpacing = 13.8 // ~1.15 cho 14pt
          p.spaceBefore = 0
          p.spaceAfter = 0
        }
      }

      await context.sync()
    })

    ElMessage.success('Áp dụng định dạng theo tuỳ chọn thành công!')
  } catch (e) {
    console.error('Apply format options failed:', e)
    ElMessage.error('Áp dụng định dạng thất bại: ' + String(e))
  } finally {
    loading.value = false
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
