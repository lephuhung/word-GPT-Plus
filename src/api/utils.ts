import { Ref } from 'vue'
import { CompletionResponse } from './types'
import { addChatMessage, createAssistantMessage } from '@/utils/chatHelpers'
import { getOrCreateDocumentId } from '@/utils/documentStorage'

export function updateResult(
  response: CompletionResponse,
  result: Ref<string>,
  historyDialog: Ref<any[]>
): void {
  result.value = response.content

  // Check if response is JSON with actions - if so, don't add message here
  // The executeWordAction will handle adding messages for JSON responses
  try {
    const text = response.content
    const fence = text.match(/```json\s*([\s\S]*?)```/i) || text.match(/```\s*([\s\S]*?)```/i)
    if (fence && fence[1]) {
      const parsed = JSON.parse(fence[1])
      if (parsed && (parsed.action || Array.isArray(parsed.actions))) {
        // This is a JSON response with actions, don't add message here
        return
      }
    }
    const start = text.indexOf('{')
    const end = text.lastIndexOf('}')
    if (start >= 0 && end > start) {
      const candidate = text.slice(start, end + 1)
      const parsed = JSON.parse(candidate)
      if (parsed && (parsed.action || Array.isArray(parsed.actions))) {
        // This is a JSON response with actions, don't add message here
        return
      }
    }
  } catch (e) {
    // Not JSON, continue with normal message handling
  }

  // Add to reactive array and save to localStorage for non-JSON responses
  const assistantMessage = createAssistantMessage(response.content)
  
  // Get current document ID to save the message
  getOrCreateDocumentId().then(documentId => {
    addChatMessage(historyDialog.value, documentId, assistantMessage)
  }).catch(error => {
    console.error('Failed to get document ID for saving bot message:', error)
    // Fallback: just add to reactive array without saving
    historyDialog.value.push({
      role: response.role || 'assistant',
      content: response.content
    })
  })
}

export function handleError(
  error: Error,
  result: Ref<string>,
  errorIssue: Ref<boolean>
): void {
  result.value = String(error)
  errorIssue.value = true
  console.error(error)
}

export function finishLoading(loading: Ref<boolean>): void {
  loading.value = false
}
