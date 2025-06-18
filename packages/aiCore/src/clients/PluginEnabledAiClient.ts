/**
 * AI Client - Cherry Studio AI Core 的主要客户端接口
 * 默认集成插件系统，提供完整的 AI 调用能力
 *
 * ## 使用方式
 *
 * ```typescript
 * import { AiClient } from '@cherry-studio/ai-core'
 *
 * // 创建客户端（默认带插件系统）
 * const client = AiClient.create('openai', {
 *   name: 'openai',
 *   apiKey: process.env.OPENAI_API_KEY
 * }, [LoggingPlugin, ContentFilterPlugin])
 *
 * // 使用方式与 UniversalAiSdkClient 完全相同
 * const result = await client.generateText('gpt-4', {
 *   messages: [{ role: 'user', content: 'Hello!' }]
 * })
 * ```
 */

import { generateObject, generateText, streamObject, streamText } from 'ai'

import { AiPlugin, createContext, PluginManager } from '../plugins'
import { ApiClientFactory } from './ApiClientFactory'
import { type ProviderId, type ProviderSettingsMap } from './types'
import { UniversalAiSdkClient } from './UniversalAiSdkClient'

/**
 * Cherry Studio AI Core 的主要客户端
 * 默认集成插件系统，提供完整的 AI 调用能力
 */
export class PluginEnabledAiClient<T extends ProviderId = ProviderId> {
  private pluginManager: PluginManager
  private baseClient: UniversalAiSdkClient<T>

  constructor(
    private readonly providerId: T,
    private readonly options: ProviderSettingsMap[T],
    plugins: AiPlugin[] = []
  ) {
    this.pluginManager = new PluginManager(plugins)
    this.baseClient = UniversalAiSdkClient.create(providerId, options)
  }

  /**
   * 添加单个插件
   */
  use(plugin: AiPlugin): this {
    this.pluginManager.use(plugin)
    return this
  }

  /**
   * 批量添加插件
   */
  usePlugins(plugins: AiPlugin[]): this {
    plugins.forEach((plugin) => this.pluginManager.use(plugin))
    return this
  }

  /**
   * 移除插件
   */
  removePlugin(pluginName: string): this {
    this.pluginManager.remove(pluginName)
    return this
  }

  /**
   * 获取插件统计信息
   */
  getPluginStats() {
    return this.pluginManager.getStats()
  }

  /**
   * 获取插件列表
   */
  getPlugins() {
    return this.pluginManager.getPlugins()
  }

  /**
   * 执行插件处理的通用逻辑
   * 1-5步骤的通用处理
   */
  private async executeWithPlugins<TParams, TResult>(
    methodName: string,
    modelId: string,
    params: TParams,
    executor: (finalModelId: string, transformedParams: TParams) => Promise<TResult>
  ): Promise<TResult> {
    // 创建请求上下文
    const context = createContext(this.providerId, modelId, params)

    try {
      // 1. 触发请求开始事件
      await this.pluginManager.executeParallel('onRequestStart', context)

      // 2. 解析模型别名
      const resolvedModelId = await this.pluginManager.executeFirst<string>('resolveModel', modelId, context)
      const finalModelId = resolvedModelId || modelId

      // 3. 转换请求参数
      const transformedParams = await this.pluginManager.executeSequential('transformParams', params, context)

      // 4. 执行具体的 API 调用
      const result = await executor(finalModelId, transformedParams)

      // 5. 转换结果（对于非流式调用）
      const transformedResult = await this.pluginManager.executeSequential('transformResult', result, context)

      // 6. 触发完成事件
      await this.pluginManager.executeParallel('onRequestEnd', context, transformedResult)

      return transformedResult
    } catch (error) {
      // 7. 触发错误事件
      await this.pluginManager.executeParallel('onError', context, undefined, error as Error)
      throw error
    }
  }

  /**
   * 执行流式调用的通用逻辑
   * 流式调用的特殊处理（支持流转换器）
   */
  private async executeStreamWithPlugins<TParams, TResult>(
    methodName: string,
    modelId: string,
    params: TParams,
    executor: (finalModelId: string, transformedParams: TParams, streamTransforms: any[]) => Promise<TResult>
  ): Promise<TResult> {
    // 创建请求上下文
    const context = createContext(this.providerId, modelId, params)

    try {
      // 1. 触发请求开始事件
      await this.pluginManager.executeParallel('onRequestStart', context)

      // 2. 解析模型别名
      const resolvedModelId = await this.pluginManager.executeFirst<string>('resolveModel', modelId, context)
      const finalModelId = resolvedModelId || modelId

      // 3. 转换请求参数
      const transformedParams = await this.pluginManager.executeSequential('transformParams', params, context)

      // 4. 收集流转换器
      const streamTransforms = this.pluginManager.collectStreamTransforms()

      // 5. 执行流式 API 调用
      const result = await executor(finalModelId, transformedParams, streamTransforms)

      // 6. 触发完成事件（注意：对于流式调用，这里触发的是开始流式响应的事件）
      await this.pluginManager.executeParallel('onRequestEnd', context, { stream: true })

      return result
    } catch (error) {
      // 7. 触发错误事件
      await this.pluginManager.executeParallel('onError', context, undefined, error as Error)
      throw error
    }
  }

  /**
   * 流式文本生成 - 集成插件系统
   */
  async streamText(
    modelId: string,
    params: Omit<Parameters<typeof streamText>[0], 'model'>
  ): Promise<ReturnType<typeof streamText>> {
    return this.executeStreamWithPlugins(
      'streamText',
      modelId,
      params,
      async (finalModelId, transformedParams, streamTransforms) => {
        // 对于流式调用，需要直接调用 AI SDK 以支持流转换器
        const model = await ApiClientFactory.createClient(this.providerId, finalModelId, this.options)

        return streamText({
          model,
          ...transformedParams,
          experimental_transform: streamTransforms.length > 0 ? streamTransforms : undefined
        })
      }
    )
  }

  /**
   * 生成文本 - 集成插件系统
   */
  async generateText(
    modelId: string,
    params: Omit<Parameters<typeof generateText>[0], 'model'>
  ): Promise<ReturnType<typeof generateText>> {
    return this.executeWithPlugins('generateText', modelId, params, async (finalModelId, transformedParams) => {
      return this.baseClient.generateText(finalModelId, transformedParams)
    })
  }

  /**
   * 生成结构化对象 - 集成插件系统
   */
  async generateObject(
    modelId: string,
    params: Omit<Parameters<typeof generateObject>[0], 'model'>
  ): Promise<ReturnType<typeof generateObject>> {
    return this.executeWithPlugins('generateObject', modelId, params, async (finalModelId, transformedParams) => {
      return this.baseClient.generateObject(finalModelId, transformedParams)
    })
  }

  /**
   * 流式生成结构化对象 - 集成插件系统
   * 注意：streamObject 目前不支持流转换器，所以使用普通的插件处理
   */
  async streamObject(
    modelId: string,
    params: Omit<Parameters<typeof streamObject>[0], 'model'>
  ): Promise<ReturnType<typeof streamObject>> {
    return this.executeWithPlugins('streamObject', modelId, params, async (finalModelId, transformedParams) => {
      return this.baseClient.streamObject(finalModelId, transformedParams)
    })
  }

  /**
   * 获取客户端信息
   */
  getClientInfo() {
    return this.baseClient.getClientInfo()
  }

  /**
   * 获取底层客户端实例（用于高级用法）
   */
  getBaseClient(): UniversalAiSdkClient<T> {
    return this.baseClient
  }

  // === 静态工厂方法 ===

  /**
   * 创建 OpenAI Compatible 客户端
   */
  static createOpenAICompatible(
    config: ProviderSettingsMap['openai-compatible'],
    plugins: AiPlugin[] = []
  ): PluginEnabledAiClient<'openai-compatible'> {
    return new PluginEnabledAiClient('openai-compatible', config, plugins)
  }

  /**
   * 创建标准提供商客户端
   */
  static create<T extends ProviderId>(
    providerId: T,
    options: ProviderSettingsMap[T],
    plugins?: AiPlugin[]
  ): PluginEnabledAiClient<T>

  static create(
    providerId: string,
    options: ProviderSettingsMap['openai-compatible'],
    plugins?: AiPlugin[]
  ): PluginEnabledAiClient<'openai-compatible'>

  static create(providerId: string, options: any, plugins: AiPlugin[] = []): PluginEnabledAiClient {
    if (providerId in ({} as ProviderSettingsMap)) {
      return new PluginEnabledAiClient(providerId as ProviderId, options, plugins)
    } else {
      // 对于未知 provider，使用 openai-compatible
      return new PluginEnabledAiClient('openai-compatible', options, plugins)
    }
  }
}

/**
 * 创建 AI 客户端的工厂函数（默认带插件系统）
 */
export function createClient<T extends ProviderId>(
  providerId: T,
  options: ProviderSettingsMap[T],
  plugins?: AiPlugin[]
): PluginEnabledAiClient<T>

export function createClient(
  providerId: string,
  options: ProviderSettingsMap['openai-compatible'],
  plugins?: AiPlugin[]
): PluginEnabledAiClient<'openai-compatible'>

export function createClient(providerId: string, options: any, plugins: AiPlugin[] = []): PluginEnabledAiClient {
  return PluginEnabledAiClient.create(providerId, options, plugins)
}

/**
 * 创建 OpenAI Compatible 客户端的便捷函数
 */
export function createCompatibleClient(
  config: ProviderSettingsMap['openai-compatible'],
  plugins: AiPlugin[] = []
): PluginEnabledAiClient<'openai-compatible'> {
  return PluginEnabledAiClient.createOpenAICompatible(config, plugins)
}
