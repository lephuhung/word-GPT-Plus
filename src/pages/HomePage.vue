<template>
  <div class="homepage">
    <!-- Header with branding -->
    <div class="header">
      <div class="brand">
        <Zap class="brand-icon" />
        <h1 class="brand-title">Word GPT+</h1>
      </div>
      <div class="actions">
        <button 
          class="settings-btn user-btn"
          @click="userPage"
          :disabled="loading"
        >
          <User class="icon" />
          <span class="username-text">{{ username }}</span>
        </button>
        <button 
          class="settings-btn"
          @click="settings"
          :disabled="loading"
        >
          <Settings class="icon" />
        </button>
      </div>
    </div>

    <!-- Word Info -->


    <!-- Main Actions -->
    <div class="main-actions">
      <button 
        class="primary-btn"
        @click="handleGetWordInfo"
        :disabled="loading"
      >
        <FileText class="icon" />
        Lấy thông tin Word
      </button>
      <button 
        class="secondary-btn"
        @click="formatOptionsExpanded = !formatOptionsExpanded"
        :disabled="loading"
      >
        <ArrowRight class="icon" />
        Định dạng văn bản
      </button>
    </div>

    <!-- Format Options - Collapsible -->
    <div class="collapsible-section format-collapse">
      <div class="collapsible-header" @click="formatOptionsExpanded = !formatOptionsExpanded">
        <div class="collapsible-title">
          <Settings class="collapsible-icon" />
           <h2 class="section-title">Tùy chọn định dạng</h2>
        </div>
        <ChevronDown class="collapse-icon" :class="{ expanded: formatOptionsExpanded }" />
      </div>
      <div class="collapsible-content" :class="{ expanded: formatOptionsExpanded }">
        <div class="settings-grid">
          <div class="setting-item">
            <el-checkbox v-model="formatOptions.addHeader">Thêm tiêu đề</el-checkbox>
            <el-checkbox v-model="formatOptions.justifyMargins">Căn lề (Justify + thụt lề)</el-checkbox>
            <el-checkbox v-model="formatOptions.lineSpacing">Giãn dòng (~1.15)</el-checkbox>
            <el-checkbox v-model="formatOptions.a4Size">Thiết lập A4 (font Times New Roman 14pt)</el-checkbox>
          </div>
          <div class="section-actions">
            <el-button type="primary" :disabled="loading" @click.stop="applyFormatOptions">Áp dụng</el-button>
          </div>
        </div>
      </div>
    </div>
        <div class="section">
      <div class="section-header">
        <FileText class="section-icon" />
        <h2 class="section-title">Thông tin Word</h2>
      </div>
      <textarea 
        v-if="wordInfo"
        :value="wordInfo"
        class="result-textarea"
        placeholder="Thông tin Word"
        rows="8"
        readonly
      ></textarea>
    </div>

    <!-- Progress Indicator -->
    <div v-if="loading" class="progress-section">
      <div class="progress-indicator">
        <Loader2 class="spinning-icon" />
        <span class="progress-text">AI is processing...</span>
      </div>
    </div>


    <!-- AI Tools - Collapsible -->
    <div class="collapsible-section">
      <div class="collapsible-header" @click="aiToolsExpanded = !aiToolsExpanded">
        <div class="collapsible-title">
          <Sparkles class="collapsible-icon" />
          {{ $t('aiTools') }}
        </div>
        <ChevronDown class="collapse-icon" :class="{ expanded: aiToolsExpanded }" />
      </div>
      <div class="collapsible-content" :class="{ expanded: aiToolsExpanded }">
        <div class="action-grid">
          <button 
            v-for="item in actionList"
            :key="item"
            class="action-btn"
            @click="performAction(item)"
            :disabled="loading"
          >
            {{ $t(item) }}
          </button>
        </div>

        <!-- Direct Prompt Input -->
        <div class="prompt-section">
          <div class="prompt-header">
            <label class="prompt-label">Gõ yêu cầu trực tiếp</label>
            <div class="prompt-actions">
              <button class="icon-btn add-btn" @click="sendDirectInput" :disabled="loading || !directInput.trim()">
                <Play class="small-icon" />
              </button>
            </div>
          </div>
          <textarea 
            v-model="directInput"
            class="prompt-textarea"
            placeholder="Nhập yêu cầu gửi lên AI"
            rows="2"
          ></textarea>
        </div>
      </div>
    </div>

    <!-- Result -->
    <div class="section">
      <div class="section-header">
        <FileText class="section-icon" />
        <h2 class="section-title">{{ $t('result') }}</h2> 
      </div>
      <textarea 
        :value="displayResult"
        class="result-textarea"
        :placeholder="$t('result')"
        rows="6"
        readonly
      ></textarea>
    </div>

    <!-- Prompts - Collapsible -->
    <div class="collapsible-section">
      <div class="collapsible-header" @click="promptsExpanded = !promptsExpanded">
        <div class="collapsible-title">
          <MessageSquare class="collapsible-icon" />
          {{ $t('prompts') }}
        </div>
        <ChevronDown class="collapse-icon" :class="{ expanded: promptsExpanded }" />
      </div>
      <div class="collapsible-content" :class="{ expanded: promptsExpanded }">
        <!-- System Prompt -->
        <div class="prompt-section">
          <div class="prompt-header">
            <label class="prompt-label">{{ $t('homeSystem') }}</label>
            <div class="prompt-actions">
              <button class="icon-btn add-btn" @click="addSystemPromptVisible = true">
                <Plus class="small-icon" />
              </button>
              <button class="icon-btn remove-btn" @click="removeSystemPromptVisible = true">
                <Minus class="small-icon" />
              </button>
            </div>
          </div>
          <div class="select-wrapper">
            <select v-model="systemPromptSelected" class="select-input" @change="(e) => handleSystemPromptChange((e.target as HTMLSelectElement)?.value ?? '')">
              <option value="" disabled>{{ $t('homeSystemDescription') }}</option>
              <option v-for="item in systemPromptList" :key="item.value" :value="item.value">
                {{ item.key }}
              </option>
            </select>
            <ChevronDown class="select-icon" />
          </div>
          <textarea 
            v-model="systemPrompt"
            class="prompt-textarea"
            :placeholder="$t('homeSystemDescription')"
            rows="2"
            @blur="handleSystemPromptChange(systemPrompt)"
          ></textarea>
        </div>

        <!-- User Prompt -->
        <div class="prompt-section">
          <div class="prompt-header">
            <label class="prompt-label">{{ $t('homePrompt') }}</label>
            <div class="prompt-actions">
              <button class="icon-btn add-btn" @click="addPromptVisible = true">
                <Plus class="small-icon" />
              </button>
              <button class="icon-btn remove-btn" @click="removePromptVisible = true">
                <Minus class="small-icon" />
              </button>
            </div>
          </div>
          <div class="select-wrapper">
            <select v-model="promptSelected" class="select-input" @change="(e) => handlePromptChange((e.target as HTMLSelectElement)?.value ?? '')">
              <option value="" disabled>{{ $t('homePromptDescription') }}</option>
              <option v-for="item in promptList" :key="item.value" :value="item.value">
                {{ item.key }}
              </option>
            </select>
            <ChevronDown class="select-icon" />
          </div>
          <textarea 
            v-model="prompt"
            class="prompt-textarea"
            :placeholder="$t('homePromptDescription')"
            rows="2"
            @blur="handlePromptChange(prompt)"
          ></textarea>
        </div>
      </div>
    </div>

    <!-- Dialogs -->
    <HomePageAddDialog
      v-model:addVisible="addSystemPromptVisible"
      v-model:addAlias="addSystemPromptAlias"
      v-model:addValue="addSystemPromptValue"
      title="addSystemPrompt"
      alias-label="addSystemPromptAlias"
      alias-placeholder="addSystemPromptAliasDescription"
      prompt-label="homeSystem"
      prompt-placeholder="addSystemPromptDescription"
      @add="addSystemPrompt"
    />
    <HomePageAddDialog
      v-model:addVisible="addPromptVisible"
      v-model:addAlias="addPromptAlias"
      v-model:addValue="addPromptValue"
      title="addPrompt"
      alias-label="addPromptAlias"
      alias-placeholder="addPromptAliasDescription"
      prompt-label="homePrompt"
      prompt-placeholder="homePromptDescription"
      @add="addPrompt"
    />
    <HomePageDialog
      v-model:removeVisible="removeSystemPromptVisible"
      v-model:removeValue="removeSystemPromptValue"
      title="removeSystemPrompt"
      :option-list="systemPromptList"
      @remove="removeSystemPrompt"
    />
    <HomePageDialog
      v-model:removeVisible="removePromptVisible"
      v-model:removeValue="removePromptValue"
      title="removePrompt"
      :option-list="promptList"
      @remove="removePrompt"
    />

    <!-- (Đã thay dialog bằng collapse ở trên) -->
  </div>
</template>

<script lang="ts" setup>
import { ElMessage } from 'element-plus'
import {
  Plus,
  Minus,
  Settings,
  Sparkles,
  FileText,
  Play,
  ArrowRight,
  Loader2,
  ChevronDown,
  MessageSquare,
  Zap,
  RefreshCw
} from 'lucide-vue-next'
import { User } from 'lucide-vue-next'
import { onBeforeMount, ref, watch, computed } from 'vue'
import { useRouter } from 'vue-router'
import { useI18n } from 'vue-i18n'

import API from '@/api'
import { getWordDocumentInfo, formatWordInfoVi } from '@/utils/wordInfo'

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

// result
const result = ref('res')
const displayResult = computed(() => {
  const text = String(result.value || '')
  if (text.length <= 200) return text
  return text.slice(0, 200) + '...'
})
// word info
const wordInfo = ref('')
const loading = ref(false)
const router = useRouter()
const historyDialog = ref<any[]>([])

const jsonIssue = ref(false)
const errorIssue = ref(false)

// Format options collapsible state
const formatOptionsExpanded = ref(false)
const formatOptions = ref({
  addHeader: true,
  justifyMargins: true,
  lineSpacing: true,
  a4Size: true
})

async function handleGetWordInfo(): Promise<void> {
  try {
    loading.value = true
    // Xoá nội dung cũ để ẩn textbox cho đến khi có kết quả mới
    wordInfo.value = ''
    const info = await getWordDocumentInfo()
    const raw = formatWordInfoVi(info)
    const text = raw
      .split('\n')
      .map(line => {
        const l = line.trim()
        if (!l) return ''
        return /^[\-•\*]/.test(l) ? l : '- ' + l
      })
      .filter(Boolean)
      .join('\n')
    wordInfo.value = text
  } catch (e) {
    console.error('Get Word info failed:', e)
    wordInfo.value = 'Không thể lấy thông tin Word: ' + String(e)
  } finally {
    loading.value = false
  }
}

// insert type
const insertType = ref<insertTypes>('replace')
const useWordFormatting = ref(true)
const insertTypeList = ['replace', 'append', 'newLine', 'NoAction'].map(
  item => ({
    label: t(item),
    value: item
  })
)

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

async function getSystemPromptList() {
  const table = promptDbInstance.table('systemPrompt')
  const list = (await table.toArray()) as unknown as IStringKeyMap[]
  systemPromptList.value = list
}

async function addSystemPrompt() {
  const table = promptDbInstance.table('systemPrompt')
  await table.put({
    key: addSystemPromptAlias.value,
    value: addSystemPromptValue.value
  })
  addSystemPromptVisible.value = false
  getSystemPromptList()
}

async function removeSystemPrompt() {
  removeSystemPromptVisible.value = false
  const table = promptDbInstance.table('systemPrompt')
  for (const value of removeSystemPromptValue.value) {
    await table.delete(value)
  }
  removeSystemPromptValue.value = []
  getSystemPromptList()
}

async function removePrompt() {
  removePromptVisible.value = false
  const table = promptDbInstance.table('userPrompt')
  for (const value of removePromptValue.value) {
    await table.delete(value)
  }
  removePromptValue.value = []
  getPromptList()
}

function handleSystemPromptChange(val: string) {
  systemPrompt.value = val
  localStorage.setItem(localStorageKey.defaultSystemPrompt, val)
}

async function getPromptList() {
  const table = promptDbInstance.table('userPrompt')
  const list = (await table.toArray()) as unknown as IStringKeyMap[]
  promptList.value = list
}

async function addPrompt() {
  const table = promptDbInstance.table('userPrompt')
  await table.put({
    key: addPromptAlias.value,
    value: addPromptValue.value
  })
  addPromptVisible.value = false
  getPromptList()
}

function handlePromptChange(val: string) {
  prompt.value = val
  localStorage.setItem(localStorageKey.defaultPrompt, val)
}

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
  await getSystemPromptList()
  if (systemPromptList.value.find(item => item.value === systemPrompt.value)) {
    systemPromptSelected.value = systemPrompt.value
  }
  prompt.value = localStorage.getItem(localStorageKey.defaultPrompt) || ''
  await getPromptList()
  if (promptList.value.find(item => item.value === prompt.value)) {
    promptSelected.value = prompt.value
  }
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
    historyDialog.value = [
      {
        role: 'user',
        content: systemMessage + '\n' + userMessage
      }
    ]
    await API.ollama.createChatCompletionStream({
      ollamaEndpoint: settingForm.value.ollamaEndpoint,
      ollamaModel:
        settingForm.value.ollamaCustomModel ||
        settingForm.value.ollamaModelSelect,
      messages: historyDialog.value,
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
  if (!jsonIssue.value) {
    if (useWordFormatting.value) {
      API.common.insertFormattedResult(result, insertType)
    } else {
      API.common.insertResult(result, insertType)
    }
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

const actionList = Object.keys(buildInPrompt) as (keyof typeof buildInPrompt)[]

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
    await API.ollama.createChatCompletionStream({
      ollamaEndpoint: settingForm.value.ollamaEndpoint,
      ollamaModel:
        settingForm.value.ollamaCustomModel ||
        settingForm.value.ollamaModelSelect,
      messages: historyDialog.value,
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
  if (useWordFormatting.value) {
    API.common.insertFormattedResult(result, insertType)
  } else {
    API.common.insertResult(result, insertType)
  }
}

// Send direct prompt to AI API
async function sendDirectInput() {
  if (!checkApiKey()) return
  const content = directInput.value.trim()
  if (!content) {
    ElMessage.warning('Vui lòng nhập yêu cầu')
    return
  }
  loading.value = true
  try {
    historyDialog.value = [
      {
        role: 'user',
        content
      }
    ]
    await API.ollama.createChatCompletionStream({
      ollamaEndpoint: settingForm.value.ollamaEndpoint,
      ollamaModel:
        settingForm.value.ollamaCustomModel ||
        settingForm.value.ollamaModelSelect,
      messages: historyDialog.value,
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
  if (useWordFormatting.value) {
    API.common.insertFormattedResult(result, insertType)
  } else {
    API.common.insertResult(result, insertType)
  }
}

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
})
</script>

<style scoped src="./HomePage.css"></style>


