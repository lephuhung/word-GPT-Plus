import axios from 'axios'
import { appConfig } from '@/config'

export interface LoginPayload {
  username: string
  password: string
}

export interface LoginResponse {
  token?: string
  success: boolean
  message?: string
}

export async function login(payload: LoginPayload): Promise<LoginResponse> {
  // Dev-mode mock: allow demo credentials without backend
  if (import.meta.env?.DEV && payload.username === 'demo' && payload.password === 'demo123') {
    return { success: true, token: 'mock-token-demo', message: 'Logged in (mock)' }
  }
  try {
    const { data } = await axios.post(appConfig.loginUrl, payload, {
      headers: { 'Content-Type': 'application/json' }
    })
    // Expect server to return { success, token?, message? }
    return {
      success: Boolean(data?.success ?? false),
      token: data?.token,
      message: data?.message
    }
  } catch (error: any) {
    // In dev, if backend is unavailable, allow any non-empty credentials
    if (import.meta.env?.DEV && payload.username && payload.password) {
      return { success: true, token: 'mock-token-dev', message: 'Logged in (fallback mock)' }
    }
    const message = error?.response?.data?.message || error?.message || 'Login failed'
    return { success: false, message }
  }
}