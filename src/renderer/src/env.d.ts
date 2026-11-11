/// <reference types="vite/client" />

// 运行时注入生成代码的库（generated-view-host.ts 使用，无需精确类型）
declare module '@babel/standalone'
declare module 'd3'

declare module '*.vue' {
  import type { DefineComponent } from 'vue'
  const component: DefineComponent<{}, {}, any>
  export default component
}

interface AixploreBridge {
  invoke: (channel: string, payload?: unknown) => Promise<any>
  onStreamEvent: (
    streamId: string,
    callback: (event: { eventName: string; payload: any }) => void
  ) => () => void
  abort: (id: string) => Promise<unknown>
}

interface Window {
  aixplore?: AixploreBridge
}
