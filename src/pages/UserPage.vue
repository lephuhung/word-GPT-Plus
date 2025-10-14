<template>
  <div class="user-page">
    <!-- Header -->
    <div class="header">
      <div class="brand">
        <Settings class="brand-icon" />
        <h1 class="brand-title">User</h1>
      </div>
      <button 
        class="back-btn"
        @click="backToHome"
      >
        <ArrowLeft class="icon" />
        {{ $t('backToHome') || 'Back to Home' }}
      </button>
    </div>

    <!-- User Info -->
    <div class="section">
      <div class="section-header">
        <Sliders class="section-icon" />
        <h2 class="section-title">Thông tin người dùng</h2>
      </div>
      <div class="settings-content">
        <div class="setting-item full-width">
          <label class="setting-label">JWT Token</label>
          <input class="text-input" :value="maskedToken" readonly />
        </div>
        <div class="actions-row">
          <button class="back-btn" @click="copyToken" :disabled="!token">
            Copy Token
          </button>
          <button class="back-btn" @click="logout">
            Logout
          </button>
        </div>
      </div>
    </div>
  </div>
</template>

<script lang="ts" setup>
import { ref, onMounted } from 'vue'
import { useRouter } from 'vue-router'
import { Settings, ArrowLeft, Sliders } from 'lucide-vue-next'

const router = useRouter()

const token = ref<string>('')
const maskedToken = ref<string>('')

onMounted(() => {
  const t = localStorage.getItem('authToken') || ''
  token.value = t
  maskedToken.value = t ? `${t.slice(0, 8)}...${t.slice(-6)}` : '(chưa đăng nhập)'
})

function backToHome() {
  router.push('/')
}

async function copyToken() {
  try {
    await navigator.clipboard.writeText(token.value)
  } catch {}
}

function logout() {
  localStorage.removeItem('authToken')
  router.push('/login')
}
</script>

<style scoped>
* { box-sizing: border-box; }
.user-page { width: 100%; max-width: 600px; margin: 0 auto; padding: 12px; background: #fafbfc; min-height: 100vh; color: #21252b; font-size: 12px; line-height: 1.3; }
.header { display: flex; align-items: center; justify-content: space-between; margin-bottom: 16px; padding: 12px 0; border-bottom: 1px solid #e1e4e8; }
.brand { display: flex; align-items: center; gap: 8px; }
.brand-icon { width: 20px; height: 20px; color: #0366d6; }
.brand-title { font-size: 18px; font-weight: 600; margin: 0; color: #24292e; }
.back-btn { display: flex; align-items: center; gap: 6px; height: 32px; padding: 0 12px; border: 1px solid #d0d7de; background: white; border-radius: 6px; cursor: pointer; transition: all 0.15s ease; color: #24292e; font-size: 12px; font-weight: 500; }
.back-btn:hover { background: #f6f8fa; border-color: #0366d6; color: #0366d6; }
.back-btn .icon { width: 14px; height: 14px; }
.section { margin-bottom: 16px; background: white; border: 1px solid #e1e4e8; border-radius: 6px; overflow: hidden; }
.section-header { display: flex; align-items: center; gap: 8px; padding: 12px 16px; background: #f6f8fa; border-bottom: 1px solid #e1e4e8; }
.section-icon { width: 14px; height: 14px; color: #586069; }
.section-title { font-size: 14px; font-weight: 600; margin: 0; color: #24292e; }
.settings-content { padding: 16px; display: flex; flex-direction: column; gap: 16px; }
.setting-item { display: flex; flex-direction: column; gap: 6px; }
.setting-item.full-width { grid-column: 1 / -1; }
.setting-label { font-size: 12px; font-weight: 500; color: #24292e; }
.text-input { width: 100%; height: 32px; padding: 0 8px; border: 1px solid #d0d7de; border-radius: 4px; background: #f6f8fa; font-size: 12px; color: #24292e; }
.actions-row { display: flex; gap: 8px; }
@media (prefers-color-scheme: dark) {
  .user-page { background: #1c1f23; color: #e1e4e8; }
  .header { border-color: #30363d; }
  .brand-title { color: #f0f6fc; }
  .section { background: #21262d; border-color: #30363d; }
  .section-header { background: #161b22; border-color: #30363d; }
  .section-title { color: #f0f6fc; }
  .back-btn { background: #21262d; border-color: #30363d; color: #e1e4e8; }
  .back-btn:hover { background: #30363d; border-color: #58a6ff; color: #58a6ff; }
  .text-input { background: #0d1117; border-color: #30363d; color: #e1e4e8; }
}
</style>