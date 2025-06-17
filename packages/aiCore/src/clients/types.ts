import { CoreMessage } from 'ai'

export type ProviderOptions = {
  name: string
  apiKey?: string
  apiHost: string
  apiVersion?: string
  headers?: Record<string, string | unknown>
}

export interface AiCoreRequest {
  modelId: string
  messages: CoreMessage[]
  tools?: Record<string, any>
  maxTokens?: number
  temperature?: number
  topP?: number
  frequencyPenalty?: number
  presencePenalty?: number
  maxRetries?: number
  abortSignal?: AbortSignal
  headers?: Record<string, string | undefined>
}
