import axios from 'axios'
import { BaseChatCompletionOptions } from './types'
import { updateResult, handleError, finishLoading } from './utils'

function deepExtractString(payload: any): string {
  if (!payload) return ''
  if (typeof payload === 'string') return payload
  if (Array.isArray(payload)) {
    for (const item of payload) {
      const found = deepExtractString(item)
      if (found) return found
    }
    return ''
  }
  if (typeof payload === 'object') {
    // Preferred keys first
    const preferredKeys = ['content', 'text', 'response', 'output', 'output_text']
    for (const key of preferredKeys) {
      const val = (payload as any)[key]
      const found = deepExtractString(val)
      if (found) return found
    }
    // Generic deep search
    for (const val of Object.values(payload)) {
      const found = deepExtractString(val)
      if (found) return found
    }
  }
  return ''
}

interface ChatCompletionStreamOptions extends BaseChatCompletionOptions {
  ollamaEndpoint: string
  ollamaModel: string
  jwtToken?: string
}

async function createChatCompletionStream(
  options: ChatCompletionStreamOptions
): Promise<void> {
  try {
    const formatedEndpoint = options.ollamaEndpoint.replace(/\/$/, '')
    const isOpenAICompatible = /\/v1(\/|$)/.test(formatedEndpoint)

    const openaiBase = formatedEndpoint
      .replace(/\/v1\/models$/, '/v1')
      .replace(/\/v1\/chat\/completions$/, '/v1')
    const nativeBase = formatedEndpoint.replace(/\/api\/chat$/, '/api')

    const url = isOpenAICompatible
      ? `${openaiBase}/chat/completions`
      : nativeBase.endsWith('/api')
        ? `${nativeBase}/chat`
        : `${formatedEndpoint}/v1/chat/completions`

    const payload = isOpenAICompatible
      ? {
          model: options.ollamaModel,
          temperature: options.temperature,
          stream: false,
          messages: options.messages
        }
      : {
          model: options.ollamaModel,
          options: { temperature: options.temperature },
          stream: false,
          messages: options.messages
        }

    const response = await axios.post(url, payload, {
      headers: {
        'Content-Type': 'application/json',
        Authorization: `Bearer ${options.jwtToken ?? 'demo'}`
      }
    })

    if (response.status !== 200) {
      throw new Error(`Status code: ${response.status}`)
    }

    // Normalize and extract content from different response shapes
    const data = response.data as any
    let content: string = ''

    if (isOpenAICompatible) {
      // OpenAI-compatible providers may return either message.content or text
      content =
        data?.choices?.[0]?.message?.content ??
        data?.choices?.[0]?.text ??
        ''
    } else {
      // Native Ollama-style response
      content =
        data?.message?.content ??
        data?.message ??
        data?.response ??
        ''
    }

    // Fallback: deep search if still empty
    if (!content || typeof content !== 'string' || content.length === 0) {
      // Try the first choice or whole payload
      content = deepExtractString(data?.choices?.[0]) || deepExtractString(data)
    }

    // Replace escaped newlines with real newlines for consistent display
    content = String(content).replace(/\\n/g, '\n')
    console.info('[Ollama] Raw data:', JSON.stringify(data).slice(0, 500) + '...')
    console.info('[Ollama] Response:', content)
    updateResult({ content }, options.result, options.historyDialog)
  } catch (error) {
    handleError(error as Error, options.result, options.errorIssue)
  } finally {
    finishLoading(options.loading)
  }
}

/**
 * Fetch available models from Ollama endpoint.
 * Uses the configured endpoint in this module and calls `/models`.
 */
async function listModels(endpointOrUrl: string): Promise<{ label: string; value: string }[]> {
  const base = endpointOrUrl.replace(/\/$/, '')
  let url = base
  // Decide URL based on provided endpoint style
  if (/\/v1\/models$/.test(base) || /\/api\/tags$/.test(base)) {
    url = base // already full path
  } else if (/\/v1$/.test(base)) {
    url = `${base}/models` // OpenAI-compatible base
  } else if (/\/api$/.test(base)) {
    url = `${base}/tags` // Native Ollama base
  } else if (/\/models$/.test(base)) {
    url = base // already /models
  } else {
    // Default to OpenAI-compatible path
    url = `${base}/v1/models`
  }
  try {
    console.info('[Ollama] Fetching models from:', url)
    const res = await axios.get(url, {
      headers: {
        'Content-Type': 'application/json',
        Authorization: 'Bearer demo'
      }
    })
    const data = res.data
    // Parse various response formats
    let models: string[] = []
    if (Array.isArray(data)) {
      models = data.map((m: any) => (typeof m === 'string' ? m : m?.id || m?.name || m?.model)).filter(Boolean)
    } else if (Array.isArray(data?.models)) {
      // Ollama native /api/tags
      models = data.models.map((m: any) => m?.name || m?.model || m?.id).filter(Boolean)
    } else if (Array.isArray(data?.data)) {
      // OpenAI-compatible /v1/models
      models = data.data.map((m: any) => (typeof m === 'string' ? m : m?.id || m?.name || m?.model)).filter(Boolean)
    }
    if (models.length === 0) {
      // Try a fallback path (switch between native and openai-compatible)
      let fallbackUrl: string | null = null
      if (/\/api\/tags$/.test(url)) {
        // switch to openai-compatible
        fallbackUrl = `${base}/v1/models`
      } else if (/\/v1\/models$/.test(url)) {
        // switch to native
        fallbackUrl = `${base}/api/tags`
      } else if (/\/models$/.test(url)) {
        // ambiguous /models -> try native
        fallbackUrl = `${base}/api/tags`
      }
      if (fallbackUrl) {
        try {
          console.info('[Ollama] Fallback fetching models from:', fallbackUrl)
          const res2 = await axios.get(fallbackUrl, { headers: { 'Content-Type': 'application/json' } })
          const data2 = res2.data
          if (Array.isArray(data2)) {
            models = data2.map((m: any) => (typeof m === 'string' ? m : m?.id || m?.name || m?.model)).filter(Boolean)
          } else if (Array.isArray(data2?.models)) {
            models = data2.models.map((m: any) => m?.name || m?.model || m?.id).filter(Boolean)
          } else if (Array.isArray(data2?.data)) {
            models = data2.data.map((m: any) => (typeof m === 'string' ? m : m?.id || m?.name || m?.model)).filter(Boolean)
          }
        } catch (fallbackErr) {
          console.error('Fallback fetch failed:', fallbackErr)
        }
      }
    }
    const unique = Array.from(new Set(models))
    return unique.map(name => ({ label: name, value: name }))
  } catch (err) {
    console.error('Failed to fetch Ollama models:', err)
    return []
  }
}

export default { createChatCompletionStream, listModels }
