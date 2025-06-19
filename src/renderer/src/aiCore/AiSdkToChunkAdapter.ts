/**
 * AI SDK 到 Cherry Studio Chunk 适配器
 * 用于将 AI SDK 的 fullStream 转换为 Cherry Studio 的 chunk 格式
 */

import { TextStreamPart } from '@cherry-studio/ai-core'
import { Chunk, ChunkType } from '@renderer/types/chunk'

export interface CherryStudioChunk {
  type: 'text-delta' | 'text-complete' | 'tool-call' | 'tool-result' | 'finish' | 'error'
  text?: string
  toolCall?: any
  toolResult?: any
  finishReason?: string
  usage?: any
  error?: any
}

/**
 * AI SDK 到 Cherry Studio Chunk 适配器类
 * 处理 fullStream 到 Cherry Studio chunk 的转换
 */
export class AiSdkToChunkAdapter {
  constructor(private onChunk: (chunk: Chunk) => void) {}

  /**
   * 处理 AI SDK 流结果
   * @param aiSdkResult AI SDK 的流结果对象
   * @returns 最终的文本内容
   */
  async processStream(aiSdkResult: any): Promise<string> {
    // 如果是流式且有 fullStream
    if (aiSdkResult.fullStream) {
      await this.readFullStream(aiSdkResult.fullStream)
    }

    // 使用 streamResult.text 获取最终结果
    return await aiSdkResult.text
  }

  /**
   * 读取 fullStream 并转换为 Cherry Studio chunks
   * @param fullStream AI SDK 的 fullStream (ReadableStream)
   */
  private async readFullStream(fullStream: ReadableStream<TextStreamPart<any>>) {
    const reader = fullStream.getReader()
    const final = {
      text: '',
      reasoning_content: ''
    }
    try {
      while (true) {
        const { done, value } = await reader.read()

        if (done) {
          break
        }

        // 转换并发送 chunk
        this.convertAndEmitChunk(value, final)
      }
    } finally {
      reader.releaseLock()
    }
  }

  /**
   * 转换 AI SDK chunk 为 Cherry Studio chunk 并调用回调
   * @param chunk AI SDK 的 chunk 数据
   */
  private convertAndEmitChunk(chunk: TextStreamPart<any>, final: { text: string; reasoning_content: string }) {
    console.log('AI SDK chunk type:', chunk.type, chunk)
    switch (chunk.type) {
      // === 文本相关事件 ===
      case 'text-delta':
        final.text += chunk.textDelta || ''
        this.onChunk({
          type: ChunkType.TEXT_DELTA,
          text: chunk.textDelta || ''
        })
        if (final.reasoning_content) {
          this.onChunk({
            type: ChunkType.THINKING_COMPLETE,
            text: final.reasoning_content || ''
          })
          final.reasoning_content = ''
        }
        break

      // === 推理相关事件 ===
      case 'reasoning':
        final.reasoning_content += chunk.textDelta || ''
        this.onChunk({
          type: ChunkType.THINKING_DELTA,
          text: chunk.textDelta || ''
        })
        break

      case 'reasoning-signature':
        // 推理签名，可以映射到思考完成
        this.onChunk({
          type: ChunkType.THINKING_COMPLETE,
          text: chunk.signature || ''
        })
        break

      case 'redacted-reasoning':
        // 被编辑的推理内容，也映射到思考
        this.onChunk({
          type: ChunkType.THINKING_DELTA,
          text: chunk.data || ''
        })
        break

      // === 工具调用相关事件 ===
      case 'tool-call-streaming-start':
        // 开始流式工具调用
        this.onChunk({
          type: ChunkType.MCP_TOOL_CREATED,
          tool_calls: [
            {
              id: chunk.toolCallId,
              name: chunk.toolName,
              args: {}
            }
          ]
        })
        break

      case 'tool-call-delta':
        // 工具调用参数的增量更新
        this.onChunk({
          type: ChunkType.MCP_TOOL_IN_PROGRESS,
          responses: [
            {
              id: chunk.toolCallId,
              tool: {
                id: chunk.toolName,
                // TODO: serverId,serverName
                serverId: 'ai-sdk',
                serverName: 'AI SDK',
                name: chunk.toolName,
                description: '',
                inputSchema: {
                  type: 'object',
                  title: chunk.toolName,
                  properties: {}
                }
              },
              arguments: {},
              status: 'invoking',
              response: chunk.argsTextDelta,
              toolCallId: chunk.toolCallId
            }
          ]
        })
        break

      case 'tool-call':
        // 完整的工具调用
        this.onChunk({
          type: ChunkType.MCP_TOOL_CREATED,
          tool_calls: [
            {
              id: chunk.toolCallId,
              name: chunk.toolName,
              args: chunk.args
            }
          ]
        })
        break

      case 'tool-result':
        // 工具调用结果
        this.onChunk({
          type: ChunkType.MCP_TOOL_COMPLETE,
          responses: [
            {
              id: chunk.toolCallId,
              tool: {
                id: chunk.toolName,
                // TODO: serverId,serverName
                serverId: 'ai-sdk',
                serverName: 'AI SDK',
                name: chunk.toolName,
                description: '',
                inputSchema: {
                  type: 'object',
                  title: chunk.toolName,
                  properties: {}
                }
              },
              arguments: chunk.args || {},
              status: 'done',
              response: chunk.result,
              toolCallId: chunk.toolCallId
            }
          ]
        })
        break

      // === 步骤相关事件 ===
      //   case 'step-start':
      //     this.onChunk({
      //       type: ChunkType.LLM_RESPONSE_CREATED
      //     })
      //     break
      case 'step-finish':
        this.onChunk({
          type: ChunkType.BLOCK_COMPLETE,
          response: {
            text: final.text || '',
            reasoning_content: final.reasoning_content || '',
            usage: {
              completion_tokens: chunk.usage.completionTokens || 0,
              prompt_tokens: chunk.usage.promptTokens || 0,
              total_tokens: chunk.usage.totalTokens || 0
            },
            metrics: chunk.usage
              ? {
                  completion_tokens: chunk.usage.completionTokens || 0,
                  time_completion_millsec: 0
                }
              : undefined
          }
        })
        break

      case 'finish':
        this.onChunk({
          type: ChunkType.TEXT_COMPLETE,
          text: final.text || '' // TEXT_COMPLETE 需要 text 字段
        })
        this.onChunk({
          type: ChunkType.LLM_RESPONSE_COMPLETE,
          response: {
            text: final.text || '',
            reasoning_content: final.reasoning_content || '',
            usage: {
              completion_tokens: chunk.usage.completionTokens || 0,
              prompt_tokens: chunk.usage.promptTokens || 0,
              total_tokens: chunk.usage.totalTokens || 0
            },
            metrics: chunk.usage
              ? {
                  completion_tokens: chunk.usage.completionTokens || 0,
                  time_completion_millsec: 0
                }
              : undefined
          }
        })
        break

      // === 源和文件相关事件 ===
      case 'source':
        // 源信息，可以映射到知识搜索完成
        this.onChunk({
          type: ChunkType.KNOWLEDGE_SEARCH_COMPLETE,
          knowledge: [
            {
              id: Number(chunk.source.id) || Date.now(),
              content: chunk.source.title || '',
              sourceUrl: chunk.source.url || '',
              type: 'url'
            }
          ]
        })
        break

      case 'file':
        // 文件相关事件，可能是图片生成
        this.onChunk({
          type: ChunkType.IMAGE_COMPLETE,
          image: {
            type: 'base64',
            images: [chunk.base64]
          }
        })
        break
      case 'error':
        this.onChunk({
          type: ChunkType.ERROR,
          error: {
            message: chunk.error || 'Unknown error'
          }
        })
        break

      default:
        // 其他类型的 chunk 可以忽略或记录日志
        console.log('Unhandled AI SDK chunk type:', chunk.type, chunk)
    }
  }
}

export default AiSdkToChunkAdapter
