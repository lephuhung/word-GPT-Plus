<template>
  <div class="login-page">
    <!-- Header -->
    <div class="header">
      <div class="brand">
        <Zap class="brand-icon" />
        <h1 class="brand-title">Word GPT+</h1>
      </div>
    </div>

    <!-- Sign In Section -->
    <div class="section">
      <div class="section-header">
        <LogIn class="section-icon" />
        <h2 class="section-title">Sign In</h2>
      </div>
      <div class="settings-content">
        <form @submit.prevent="onSubmit">
          <div class="setting-item">
            <label class="setting-label" for="username">Username</label>
            <input
              id="username"
              v-model.trim="username"
              type="text"
              class="text-input"
              placeholder="Enter username"
              required
            />
          </div>

          <div class="setting-item">
            <label class="setting-label" for="password">Password</label>
            <input
              id="password"
              v-model="password"
              type="password"
              class="text-input"
              placeholder="Enter password"
              required
            />
          </div>

          <button class="primary-btn" type="submit" :disabled="loading">
            {{ loading ? 'Signing in...' : 'Sign In' }}
          </button>
          <p v-if="error" class="error">{{ error }}</p>
          <p v-if="isDev" class="hint">Use demo/demo123 to login (mock)</p>
        </form>
      </div>
    </div>
  </div>
 </template>

<script lang="ts" setup>
import { ref } from 'vue'
import { useRouter } from 'vue-router'
import { LogIn, Zap } from 'lucide-vue-next'
import { login } from '@/api/auth'

const isDev = !!import.meta.env?.DEV

const username = ref(isDev ? 'demo' : '')
const password = ref(isDev ? 'demo123' : '')
const loading = ref(false)
const error = ref('')

const router = useRouter()

async function onSubmit() {
  error.value = ''
  if (!username.value || !password.value) {
    error.value = 'Please enter username and password'
    return
  }
  loading.value = true
  const res = await login({ username: username.value, password: password.value })
  loading.value = false
  if (!res.success) {
    error.value = res.message || 'Invalid credentials'
    return
  }
  if (res.token) {
    localStorage.setItem('authToken', res.token)
  }
  router.push('/')
}
</script>

<style scoped>
/* Global styles */
* {
  box-sizing: border-box;
}

.login-page {
  width: 100%;
  max-width: 320px; /* Match HomePage compact width */
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
  margin-bottom: 12px;
  padding: 8px 0;
  border-bottom: 1px solid #e1e4e8;
}

.brand {
  display: flex;
  align-items: center;
  gap: 6px;
}

.brand-icon {
  width: 16px;
  height: 16px;
  color: #0366d6;
}

.brand-title {
  font-size: 14px;
  font-weight: 600;
  margin: 0;
  color: #24292e;
}

/* Section */
.section {
  margin-bottom: 12px;
  background: white;
  border: 1px solid #e1e4e8;
  border-radius: 6px;
  overflow: hidden;
}

.section-header {
  display: flex;
  align-items: center;
  gap: 6px;
  padding: 8px 12px;
  background: #f6f8fa;
  border-bottom: 1px solid #e1e4e8;
}

.section-icon {
  width: 12px;
  height: 12px;
  color: #586069;
}

.section-title {
  font-size: 12px;
  font-weight: 600;
  margin: 0;
  color: #24292e;
}

.settings-content {
  padding: 8px 12px; /* compact like Home */
}

.setting-item {
  display: flex;
  flex-direction: column;
  gap: 4px;
  margin-bottom: 8px;
}

.setting-label {
  font-size: 11px;
  font-weight: 500;
  color: #24292e;
}

.text-input {
  width: 100%;
  height: 28px; /* compact input height */
  padding: 0 8px;
  border: 1px solid #d0d7de;
  border-radius: 4px;
  background: white;
  font-size: 11px;
  color: #24292e;
  transition: all 0.15s ease;
}

.text-input:focus {
  outline: none;
  border-color: #0366d6;
  box-shadow: 0 0 0 2px rgba(3, 102, 214, 0.1);
}

.text-input:disabled {
  background: #f6f8fa;
  cursor: not-allowed;
  opacity: 0.7;
}

.primary-btn {
  display: flex;
  align-items: center;
  justify-content: center;
  gap: 6px;
  height: 32px;
  border: none;
  border-radius: 6px;
  font-size: 12px;
  font-weight: 500;
  cursor: pointer;
  transition: all 0.15s ease;
  background: #0366d6;
  color: white;
}

.primary-btn:hover:not(:disabled) {
  background: #0256cc;
}

.primary-btn:disabled {
  opacity: 0.5;
  cursor: not-allowed;
}

.error {
  margin-top: 8px;
  color: #d1242f;
  font-size: 11px;
}

.hint {
  margin-top: 6px;
  color: #586069;
  font-size: 11px;
}

/* Dark mode support for Word */
@media (prefers-color-scheme: dark) {
  .login-page {
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
  .text-input {
    background: #0d1117;
    border-color: #30363d;
    color: #e1e4e8;
  }
  .text-input:disabled {
    background: #161b22;
  }
}
</style>