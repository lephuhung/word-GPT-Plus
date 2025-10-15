<template>
  <div class="settings-page">
    <!-- Header -->
    <div class="header">
      <div class="brand">
        <Settings class="brand-icon" />
        <h1 class="brand-title">{{ $t('settings') }}</h1>
      </div>
      <button 
        class="back-btn"
        @click="backToHome"
      >
        <ArrowLeft class="icon" />
        {{ $t('backToHome') }}
      </button>
    </div>

    <!-- Quick Settings Section -->
    <div class="section">
      <div class="section-header">
        <Sliders class="section-icon" />
        <h2 class="section-title">{{ $t('quickSettings') }}</h2>
      </div>
      <div class="settings-content">
        <div class="settings-row">
          <div class="setting-item">
            <label class="setting-label">{{ $t('localLanguageLabel') }}</label>
            <div class="select-wrapper">
              <select 
                v-model="settingForm.localLanguage"
                class="select-input"
              >
                <option value="" disabled>{{ getPlaceholder('localLanguage') }}</option>
                <option
                  v-for="item in settingPreset.localLanguage.optionList"
                  :key="item.value"
                  :value="item.value"
                >
                  {{ item.label }}
                </option>
              </select>
              <ChevronDown class="select-icon" />
            </div>
          </div>
          <div class="setting-item">
            <label class="setting-label">{{ $t('replyLanguageLabel') }}</label>
            <div class="select-wrapper">
              <select 
                v-model="settingForm.replyLanguage"
                class="select-input"
              >
                <option value="" disabled>{{ getPlaceholder('replyLanguage') }}</option>
                <option
                  v-for="item in settingPreset.replyLanguage.optionList"
                  :key="item.value"
                  :value="item.value"
                >
                  {{ item.label }}
                </option>
              </select>
              <ChevronDown class="select-icon" />
            </div>
          </div>
        </div>
        <!-- Ollama Configuration (moved into Quick Settings) -->
        <div class="settings-row">
          <div class="setting-item">
            <label class="setting-label">{{ $t(getLabel('ollamaEndpoint')) }}</label>
            <input
              v-model="settingForm.ollamaEndpoint"
              class="text-input"
              type="text"
              :placeholder="$t(getPlaceholder('ollamaEndpoint'))"
            />
          </div>
          <div class="setting-item">
            <label class="setting-label">{{ $t(getLabel('ollamaModelSelect')) }}</label>
            <div class="select-wrapper">
              <select 
                v-model="settingForm.ollamaModelSelect"
                class="select-input"
              >
                <option value="" disabled>{{ $t(getPlaceholder('ollamaModelSelect')) }}</option>
                <option
                  v-for="option in settingPreset.ollamaModelSelect.optionList"
                  :key="option.value"
                  :value="option.value"
                >
                  {{ option.label }}
                </option>
              </select>
              <ChevronDown class="select-icon" />
            </div>
          </div>
        </div>
        <div class="setting-item full-width">
          <label class="setting-label">{{ $t(getLabel('ollamaTemperature')) }}</label>
          <input
            v-model.number="settingForm.ollamaTemperature"
            class="number-input"
            type="number"
            min="0"
            max="2"
            step="0.1"
            :placeholder="$t(getPlaceholder('ollamaTemperature'))"
          />
        </div>

        <!-- Insert type -->
        <div class="setting-item full-width">
          <label class="setting-label">{{ $t('insertTypeLabel') }}</label>
          <div class="select-wrapper">
            <select 
              v-model="insertType"
              class="select-input"
              @change="(e) => handleInsertTypeChange((e.target as HTMLSelectElement)?.value as insertTypes)"
            >
              <option v-for="item in insertTypeList" :key="item.value" :value="item.value">
                {{ item.label }}
              </option>
            </select>
            <ChevronDown class="select-icon" />
          </div>
        </div>
      </div>
    </div>
    <!-- Prompts Section -->
    <div class="section">
      <div class="section-header">
        <MessageSquare class="section-icon" />
        <h2 class="section-title">{{ $t('prompts') }}</h2>
      </div>
      <div class="settings-content">
        <!-- System Prompt -->
        <div class="prompt-section">
          <div class="prompt-header">
            <label class="prompt-label">{{ $t('homeSystem') }}</label>
            <div class="prompt-actions">
              <button class="icon-btn" @click="addSystemPromptVisible = true">
                <Plus class="small-icon" />
              </button>
              <button class="icon-btn" @click="removeSystemPromptVisible = true">
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
              <button class="icon-btn" @click="addPromptVisible = true">
                <Plus class="small-icon" />
              </button>
              <button class="icon-btn" @click="removePromptVisible = true">
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

    <!-- Dialogs for prompt management -->
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
  </div>
</template>

<script lang="ts" setup>
import { onBeforeMount, watch, ref } from 'vue'
import { useRouter } from 'vue-router'
import { Settings, ArrowLeft, Sliders, ChevronDown, MessageSquare, Plus, Minus } from 'lucide-vue-next'
import { ElMessage } from 'element-plus'
import API from '@/api'

import { getLabel, getPlaceholder } from '@/utils/common'
import { SettingNames, settingPreset } from '@/utils/settingPreset'
import useSettingForm from '@/utils/settingForm'
import { localStorageKey } from '@/utils/enum'
import { promptDbInstance } from '@/store/promtStore'
import HomePageDialog from '@/components/HomePageDialog.vue'
import HomePageAddDialog from '@/components/HomePageAddDialog.vue'

const router = useRouter()
const { settingForm, settingFormKeys } = useSettingForm()

const getApiInputSettings = (platform: string) => {
  return Object.keys(settingForm.value).filter(
    key =>
      key.startsWith(platform) &&
      settingPreset[key as SettingNames].type === 'input'
  )
}

const getApiNumSettings = (platform: string) => {
  return Object.keys(settingForm.value).filter(
    key =>
      key.startsWith(platform) &&
      settingPreset[key as SettingNames].type === 'inputNum'
  )
}

const getApiSelectSettings = (platform: string) => {
  return Object.keys(settingForm.value).filter(
    key =>
      key.startsWith(platform) &&
      settingPreset[key as SettingNames].type === 'select'
  )
}

const addWatch = () => {
  settingFormKeys.forEach(key => {
    watch(
      () => settingForm.value[key],
      () => {
        if (settingPreset[key].saveFunc) {
          settingPreset[key].saveFunc!(settingForm.value[key])
          return
        }
        localStorage.setItem(
          settingPreset[key].saveKey || key,
          settingForm.value[key]
        )
      }
    )
  })
}

onBeforeMount(() => {
  addWatch()
  // Luôn mặc định dùng Ollama và tự nạp models nếu có endpoint
  if (settingForm.value.api !== 'ollama') {
    settingForm.value.api = 'ollama'
  }
  if ((settingForm.value.ollamaEndpoint || '').trim()) {
    loadOllamaModels()
  }
  getSystemPromptList()
  getPromptList()
})

function backToHome() {
  router.push('/')
}

// Insert Type control
const insertType = ref<insertTypes>((localStorage.getItem(localStorageKey.insertType) as insertTypes) || 'replace')
const insertTypeList = ['replace', 'append', 'newLine', 'NoAction'].map(item => ({ label: item, value: item }))
function handleInsertTypeChange(val: insertTypes) {
  insertType.value = val
  localStorage.setItem(localStorageKey.insertType, val)
}

// Load Ollama models
async function loadOllamaModels() {
  try {
    const endpoint = (settingForm.value.ollamaEndpoint || '').trim()
    if (!endpoint) {
      ElMessage.warning('Vui lòng cấu hình Ollama endpoint trước')
      return
    }
    const models = await API.ollama.listModels(endpoint)
    if (models && models.length > 0) {
      settingPreset.ollamaModelSelect.optionList = models
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

watch(
  () => settingForm.value.ollamaEndpoint,
  (val, oldVal) => {
    if (val && val !== oldVal) {
      loadOllamaModels()
    }
  }
)

// Prompt management state and handlers
const systemPrompt = ref<string>(localStorage.getItem(localStorageKey.defaultSystemPrompt) || '')
const systemPromptSelected = ref('')
const systemPromptList = ref<IStringKeyMap[]>([])
const addSystemPromptVisible = ref(false)
const addSystemPromptAlias = ref('')
const addSystemPromptValue = ref('')
const removeSystemPromptVisible = ref(false)
const removeSystemPromptValue = ref<any[]>([])

const prompt = ref<string>(localStorage.getItem(localStorageKey.defaultPrompt) || '')
const promptSelected = ref('')
const promptList = ref<IStringKeyMap[]>([])
const addPromptVisible = ref(false)
const addPromptAlias = ref('')
const addPromptValue = ref('')
const removePromptVisible = ref(false)
const removePromptValue = ref<any[]>([])

async function getSystemPromptList() {
  const table = promptDbInstance.table('systemPrompt')
  const list = (await table.toArray()) as unknown as IStringKeyMap[]
  systemPromptList.value = list
  if (systemPrompt.value && list.find(item => item.value === systemPrompt.value)) {
    systemPromptSelected.value = systemPrompt.value
  }
}

async function addSystemPrompt() {
  const table = promptDbInstance.table('systemPrompt')
  await table.put({ key: addSystemPromptAlias.value, value: addSystemPromptValue.value })
  addSystemPromptVisible.value = false
  addSystemPromptAlias.value = ''
  addSystemPromptValue.value = ''
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

function handleSystemPromptChange(val: string) {
  systemPrompt.value = val
  localStorage.setItem(localStorageKey.defaultSystemPrompt, val)
}

async function getPromptList() {
  const table = promptDbInstance.table('userPrompt')
  const list = (await table.toArray()) as unknown as IStringKeyMap[]
  promptList.value = list
  if (prompt.value && list.find(item => item.value === prompt.value)) {
    promptSelected.value = prompt.value
  }
}

async function addPrompt() {
  const table = promptDbInstance.table('userPrompt')
  await table.put({ key: addPromptAlias.value, value: addPromptValue.value })
  addPromptVisible.value = false
  addPromptAlias.value = ''
  addPromptValue.value = ''
  getPromptList()
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

function handlePromptChange(val: string) {
  prompt.value = val
  localStorage.setItem(localStorageKey.defaultPrompt, val)
}
</script>

<style scoped>
/* Global styles */
* {
  box-sizing: border-box;
}

.settings-page {
  width: 100%;
  max-width: 600px;
  margin: 0 auto;
  padding: 12px;
  font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, 'Helvetica Neue', Arial, sans-serif;
  background: #fafbfc;
  min-height: 100vh;
  color: #21252b;
  font-size: 12px;
  line-height: 1.3;
}

/* Header */
.header {
  display: flex;
  align-items: center;
  justify-content: space-between;
  margin-bottom: 16px;
  padding: 12px 0;
  border-bottom: 1px solid #e1e4e8;
}

.brand {
  display: flex;
  align-items: center;
  gap: 8px;
}

.brand-icon {
  width: 20px;
  height: 20px;
  color: #0366d6;
}

.brand-title {
  font-size: 18px;
  font-weight: 600;
  margin: 0;
  color: #24292e;
}

.back-btn {
  display: flex;
  align-items: center;
  gap: 6px;
  height: 32px;
  padding: 0 12px;
  border: 1px solid #d0d7de;
  background: white;
  border-radius: 6px;
  cursor: pointer;
  transition: all 0.15s ease;
  color: #24292e;
  font-size: 12px;
  font-weight: 500;
}

.back-btn:hover {
  background: #f6f8fa;
  border-color: #0366d6;
  color: #0366d6;
}

.back-btn .icon {
  width: 14px;
  height: 14px;
}

/* Sections */
.section {
  margin-bottom: 16px;
  background: white;
  border: 1px solid #e1e4e8;
  border-radius: 6px;
  overflow: hidden;
}

.section-header {
  display: flex;
  align-items: center;
  gap: 8px;
  padding: 12px 16px;
  background: #f6f8fa;
  border-bottom: 1px solid #e1e4e8;
}

.section-icon {
  width: 14px;
  height: 14px;
  color: #586069;
}

.section-title {
  font-size: 14px;
  font-weight: 600;
  margin: 0;
  color: #24292e;
}

/* Settings Content */
.settings-content {
  padding: 16px;
  display: flex;
  flex-direction: column;
  gap: 16px;
}

.settings-row {
  display: grid;
  grid-template-columns: 1fr 1fr;
  gap: 16px;
}

.setting-item {
  display: flex;
  flex-direction: column;
  gap: 6px;
}

.setting-item.full-width {
  grid-column: 1 / -1;
}

.setting-label {
  font-size: 12px;
  font-weight: 500;
  color: #24292e;
}

/* Prompt section styles */
.prompt-section { margin-top: 8px; }
.prompt-header { display: flex; align-items: center; justify-content: space-between; margin-bottom: 6px; }
.prompt-label { font-size: 12px; font-weight: 600; color: #24292e; }
.prompt-actions { display: flex; gap: 6px; }
.icon-btn { display: inline-flex; align-items: center; justify-content: center; height: 28px; padding: 0 8px; border: 1px solid #d0d7de; background: white; border-radius: 6px; cursor: pointer; }
.small-icon { width: 14px; height: 14px; }
.prompt-textarea { width: 100%; min-height: 64px; padding: 8px; border: 1px solid #d0d7de; border-radius: 4px; background: white; font-size: 12px; color: #24292e; }

/* Input Styles */
.text-input,
.number-input {
  width: 100%;
  height: 32px;
  padding: 0 8px;
  border: 1px solid #d0d7de;
  border-radius: 4px;
  background: white;
  font-size: 12px;
  color: #24292e;
  transition: all 0.15s ease;
}

.text-input:focus,
.number-input:focus {
  outline: none;
  border-color: #0366d6;
  box-shadow: 0 0 0 2px rgba(3, 102, 214, 0.1);
}

.text-input:disabled,
.number-input:disabled {
  background: #f6f8fa;
  cursor: not-allowed;
  opacity: 0.7;
}

/* Select Styles */
.select-wrapper {
  position: relative;
}

.select-input {
  width: 100%;
  height: 32px;
  padding: 0 28px 0 8px;
  border: 1px solid #d0d7de;
  border-radius: 4px;
  background: white;
  font-size: 12px;
  color: #24292e;
  cursor: pointer;
  appearance: none;
  transition: all 0.15s ease;
}

.select-input:focus {
  outline: none;
  border-color: #0366d6;
  box-shadow: 0 0 0 2px rgba(3, 102, 214, 0.1);
}

.select-input:disabled {
  background: #f6f8fa;
  cursor: not-allowed;
  opacity: 0.7;
}

.select-icon {
  position: absolute;
  right: 8px;
  top: 50%;
  transform: translateY(-50%);
  width: 14px;
  height: 14px;
  color: #586069;
  pointer-events: none;
}

/* Number Input Specific Styles */
.number-input {
  appearance: textfield;
  -webkit-appearance: textfield;
  -moz-appearance: textfield;
}

.number-input::-webkit-outer-spin-button,
.number-input::-webkit-inner-spin-button {
  -webkit-appearance: none;
  margin: 0;
}

/* Responsive Design */
@media (max-width: 768px) {
  .settings-page {
    max-width: 100%;
    padding: 8px;
  }

  .settings-row {
    grid-template-columns: 1fr;
    gap: 12px;
  }

  .header {
    flex-direction: column;
    gap: 12px;
    text-align: center;
  }

  .back-btn {
    width: 100%;
    justify-content: center;
  }

  .settings-content {
    padding: 12px;
  }
}

@media (max-width: 480px) {
  .settings-page {
    padding: 6px;
  }

  .brand-title {
    font-size: 16px;
  }

  .section-title {
    font-size: 12px;
  }

  .settings-content {
    padding: 10px;
    gap: 12px;
  }
}

/* Dark mode support */
@media (prefers-color-scheme: dark) {
  .settings-page {
    background: #1c1f23;
    color: #e1e4e8;
  }
  
  .header {
    border-color: #30363d;
  }
  
  .brand-title {
    color: #f0f6fc;
  }
  
  .section {
    background: #21262d;
    border-color: #30363d;
  }
  
  .section-header {
    background: #161b22;
    border-color: #30363d;
  }
  
  .section-title {
    color: #f0f6fc;
  }
  
  .back-btn {
    background: #21262d;
    border-color: #30363d;
    color: #e1e4e8;
  }
  
  .back-btn:hover {
    background: #30363d;
    border-color: #58a6ff;
    color: #58a6ff;
  }
  
  .text-input,
  .number-input,
  .select-input {
    background: #0d1117;
    border-color: #30363d;
    color: #e1e4e8;
  }
  .prompt-textarea { background: #0d1117; border-color: #30363d; color: #e1e4e8; }

  .text-input:disabled,
  .number-input:disabled,
  .select-input:disabled {
    background: #161b22;
  }
  
  .setting-label {
    color: #f0f6fc;
  }
}
</style>
