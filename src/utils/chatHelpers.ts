/**
 * Helper functions for chat message management
 */

import { type ChatMessage, addMessageToHistory } from './documentStorage'

/**
 * Add a message to the chat history array and save to localStorage
 */
export function addChatMessage(
  historyArray: ChatMessage[], 
  documentId: string, 
  message: Omit<ChatMessage, 'id' | 'timestamp'>
): void {
  const chatMessage: ChatMessage = {
    ...message,
    id: `msg_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`,
    timestamp: Date.now()
  }
  
  console.log('Adding chat message:', chatMessage, 'to document:', documentId)
  
  // Add to reactive array
  historyArray.push(chatMessage)
  
  // Save to localStorage
  addMessageToHistory(documentId, chatMessage)
  
  console.log('Chat message saved. Current localStorage:', localStorage.getItem(`chat_history_${documentId}`))
}

/**
 * Create a user message object
 */
export function createUserMessage(content: string): Omit<ChatMessage, 'id' | 'timestamp'> {
  return {
    role: 'user',
    content
  }
}

/**
 * Create an assistant message object
 */
export function createAssistantMessage(
  content: string, 
  noActions?: boolean
): Omit<ChatMessage, 'id' | 'timestamp'> {
  return {
    role: 'assistant',
    content,
    noActions
  }
}