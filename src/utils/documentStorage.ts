/**
 * Utility functions for managing document ID via Custom Document Properties
 * and chat history via localStorage
 */

export interface ChatMessage {
  id: string
  role: 'user' | 'assistant'
  content: string
  timestamp: number
  noActions?: boolean
}

export interface DocumentChatHistory {
  documentId: string
  messages: ChatMessage[]
  lastUpdated: number
}

// Custom Document Properties key for storing document ID
const DOCUMENT_ID_PROPERTY = 'WordGPT_DocumentId'

/**
 * Generate a unique document ID
 */
function generateDocumentId(): string {
  return `doc_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`
}

/**
 * Get document ID from Custom Document Properties
 * If not exists, create a new one and save it to the document
 */
export async function getOrCreateDocumentId(): Promise<string> {
  return Word.run(async (context) => {
    try {
      console.log('Attempting to access Word document custom properties...')
      
      // Try to get existing document ID from custom properties
      const customProperties = context.document.properties.customProperties
      customProperties.load('items')
      await context.sync()

      console.log('Custom properties loaded. Total properties:', customProperties.items.length)
      console.log('All custom properties:', customProperties.items.map(p => ({ key: p.key, value: p.value })))

      // Look for existing document ID property
      const existingProperty = customProperties.items.find(
        prop => prop.key === DOCUMENT_ID_PROPERTY
      )

      console.log('Looking for property with key:', DOCUMENT_ID_PROPERTY)
      console.log('Found existing property:', existingProperty ? 'YES' : 'NO')

      if (existingProperty) {
        existingProperty.load('value')
        await context.sync()
        console.log('Found existing document ID in custom properties:', existingProperty.value)
        return existingProperty.value as string
      }

      // Create new document ID if not exists
      console.log('No existing document ID found, creating new one...')
      const newDocumentId = generateDocumentId()
      console.log('Generated new document ID:', newDocumentId)
      
      customProperties.add(DOCUMENT_ID_PROPERTY, newDocumentId)
      await context.sync()

      console.log('New document ID created and saved to custom properties:', newDocumentId)
      
      // Verify it was saved
      customProperties.load('items')
      await context.sync()
      const verifyProperty = customProperties.items.find(prop => prop.key === DOCUMENT_ID_PROPERTY)
      console.log('Verification - property saved:', verifyProperty ? 'YES' : 'NO')
      if (verifyProperty) {
        verifyProperty.load('value')
        await context.sync()
        console.log('Verification - saved value:', verifyProperty.value)
      }
      
      return newDocumentId
    } catch (error) {
      console.error('Error managing document ID in Word API:', error)
      // Fallback to session-based ID if custom properties fail
      const fallbackId = `session_${Date.now()}`
      console.log('Using session fallback ID:', fallbackId)
      return fallbackId
    }
  })
}

/**
 * Get document ID from Custom Document Properties (read-only)
 */
export async function getDocumentId(): Promise<string | null> {
  return Word.run(async (context) => {
    try {
      const customProperties = context.document.properties.customProperties
      customProperties.load('items')
      await context.sync()

      const existingProperty = customProperties.items.find(
        prop => prop.key === DOCUMENT_ID_PROPERTY
      )

      if (existingProperty) {
        existingProperty.load('value')
        await context.sync()
        return existingProperty.value as string
      }

      return null
    } catch (error) {
      console.error('Error reading document ID:', error)
      return null
    }
  })
}

/**
 * Set document ID in Custom Document Properties
 */
export async function setDocumentId(documentId: string): Promise<void> {
  return Word.run(async (context) => {
    try {
      const customProperties = context.document.properties.customProperties
      customProperties.load('items')
      await context.sync()

      // Remove existing property if exists
      const existingProperty = customProperties.items.find(
        prop => prop.key === DOCUMENT_ID_PROPERTY
      )
      if (existingProperty) {
        existingProperty.delete()
      }

      // Add new property
      customProperties.add(DOCUMENT_ID_PROPERTY, documentId)
      await context.sync()
    } catch (error) {
      console.error('Error setting document ID:', error)
      throw error
    }
  })
}

/**
 * Get chat history for a specific document from localStorage
 */
export function getChatHistory(documentId: string): ChatMessage[] {
  try {
    const key = `wordgpt_chat_${documentId}`
    console.log('Getting chat history for key:', key)
    const stored = localStorage.getItem(key)
    console.log('Stored data:', stored)
    if (!stored) return []

    const history: DocumentChatHistory = JSON.parse(stored)
    console.log('Parsed history:', history)
    return history.messages || []
  } catch (error) {
    console.error('Error loading chat history:', error)
    return []
  }
}

/**
 * Save chat history for a specific document to localStorage
 */
export function saveChatHistory(documentId: string, messages: ChatMessage[]): void {
  try {
    const key = `wordgpt_chat_${documentId}`
    const history: DocumentChatHistory = {
      documentId,
      messages,
      lastUpdated: Date.now()
    }
    console.log('Saving chat history with key:', key, 'data:', history)
    localStorage.setItem(key, JSON.stringify(history))
    console.log('Chat history saved successfully')
  } catch (error) {
    console.error('Error saving chat history:', error)
  }
}

/**
 * Add a message to chat history for a specific document
 */
export function addMessageToHistory(documentId: string, message: ChatMessage): void {
  const currentHistory = getChatHistory(documentId)
  currentHistory.push(message)
  saveChatHistory(documentId, currentHistory)
}

/**
 * Clear chat history for a specific document
 */
export function clearChatHistory(documentId: string): void {
  try {
    const key = `wordgpt_chat_${documentId}`
    localStorage.removeItem(key)
  } catch (error) {
    console.error('Error clearing chat history:', error)
  }
}

/**
 * Get all document IDs that have chat history
 */
export function getAllDocumentIds(): string[] {
  try {
    const documentIds: string[] = []
    for (let i = 0; i < localStorage.length; i++) {
      const key = localStorage.key(i)
      if (key && key.startsWith('wordgpt_chat_')) {
        const documentId = key.replace('wordgpt_chat_', '')
        documentIds.push(documentId)
      }
    }
    return documentIds
  } catch (error) {
    console.error('Error getting document IDs:', error)
    return []
  }
}

/**
 * Clean up old chat histories (older than specified days)
 */
export function cleanupOldChatHistories(daysToKeep: number = 30): void {
  try {
    const cutoffTime = Date.now() - (daysToKeep * 24 * 60 * 60 * 1000)
    const keysToRemove: string[] = []

    for (let i = 0; i < localStorage.length; i++) {
      const key = localStorage.key(i)
      if (key && key.startsWith('wordgpt_chat_')) {
        try {
          const stored = localStorage.getItem(key)
          if (stored) {
            const history: DocumentChatHistory = JSON.parse(stored)
            if (history.lastUpdated < cutoffTime) {
              keysToRemove.push(key)
            }
          }
        } catch (parseError) {
          // Remove corrupted entries
          keysToRemove.push(key)
        }
      }
    }

    keysToRemove.forEach(key => localStorage.removeItem(key))
    console.log(`Cleaned up ${keysToRemove.length} old chat histories`)
  } catch (error) {
    console.error('Error cleaning up chat histories:', error)
  }
}