import { contextBridge, ipcRenderer } from 'electron'

/**
 * 渲染进程可用的本地能力桥。
 *
 * - invoke: 调用主进程 IPC handler
 * - onStreamEvent: 订阅流式生成事件（status/delta/done/fail），返回取消订阅函数
 * - abort: 中断一次生成（requestId 或 streamId）
 */
contextBridge.exposeInMainWorld('aixplore', {
  invoke: (channel: string, payload?: unknown): Promise<unknown> =>
    ipcRenderer.invoke(channel, payload),
  onStreamEvent: (
    streamId: string,
    callback: (event: { eventName: string; payload: Record<string, unknown> }) => void
  ): (() => void) => {
    const listener = (_event: unknown, data: { streamId: string; eventName: string; payload: Record<string, unknown> }) => {
      if (data && data.streamId === streamId) callback(data)
    }
    ipcRenderer.on('aixplore:stream-event', listener)
    return () => ipcRenderer.removeListener('aixplore:stream-event', listener)
  },
  abort: (id: string): Promise<unknown> => ipcRenderer.invoke('excel:abort', { id })
})
