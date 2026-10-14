import { reactive } from 'vue'

export type ConfirmOptions = {
  title: string
  message?: string
  confirmLabel?: string
  cancelLabel?: string
  danger?: boolean
}

type ConfirmState = {
  open: boolean
  title: string
  message: string
  confirmLabel: string
  cancelLabel: string
  danger: boolean
}

const state = reactive<ConfirmState>({
  open: false,
  title: '',
  message: '',
  confirmLabel: 'Подтвердить',
  cancelLabel: 'Отмена',
  danger: true,
})

let resolver: ((value: boolean) => void) | null = null

function settle(value: boolean) {
  if (!resolver) return
  const resolve = resolver
  resolver = null
  state.open = false
  resolve(value)
}

export function uiConfirm(options: ConfirmOptions): Promise<boolean> {
  if (resolver) settle(false)

  state.title = options.title
  state.message = options.message ?? ''
  state.confirmLabel = options.confirmLabel ?? 'Подтвердить'
  state.cancelLabel = options.cancelLabel ?? 'Отмена'
  state.danger = options.danger ?? true
  state.open = true

  return new Promise<boolean>((resolve) => {
    resolver = resolve
  })
}

export function confirmAccept() {
  settle(true)
}

export function confirmCancel() {
  settle(false)
}

export function useConfirmState() {
  return state
}
