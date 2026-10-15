import { onBeforeUnmount, ref, shallowRef } from 'vue'
import { setSelectionInjector } from './selection'

export type DocTextRange = {
  setText: (text: string) => boolean
  getText?: () => string
  describe?: () => { startOffset?: number; endOffset?: number; text?: string }
  setTextStyle?: (style: {
    bl?: number
    it?: number
    fs?: number
    ul?: { s?: number }
    cl?: { rgb?: string }
  }) => boolean
}

export type DocParagraphInfo = {
  paragraphIndex: number
  startOffset: number
  endOffset: number
  paragraph?: { paragraphStyle?: { namedStyleType?: number } }
}

export type DocParagraph = {
  getId?: () => string
  getText: () => string
  setText: (text: string) => boolean
  appendText: (text: string) => boolean
  getInfo?: () => DocParagraphInfo
  getRange?: () => { startOffset?: number; endOffset?: number }
  setStyle?: (style: { namedStyleType?: number }) => boolean
  isListItem?: () => boolean
  findAllText: (
    text: string,
    options?: { matchCase?: boolean },
  ) => DocTextRange[]
}

export type UniverDocument = {
  getId: () => string
  save: () => { body?: { dataStream?: string } }
  getSnapshot?: () => { body?: { dataStream?: string } }
  getBody?: (segmentId?: string) => { dataStream?: string }
  insertText: (index: number, text: string, segmentId?: string) => boolean
  appendParagraph: (text?: string, segmentId?: string) => DocParagraph
  insertParagraph?: (index: number, text?: string, segmentId?: string) => DocParagraph
  getParagraphs: (segmentId?: string) => DocParagraph[]
  deleteRange: (range: { startOffset: number; endOffset: number; segmentId?: string }) => boolean
  getTextRange?: (startOffset: number, endOffset: number, segmentId?: string) => DocTextRange
  setSelection?: (startOffset: number, endOffset: number) => void
  findParagraphByText?: (text: string) => DocParagraph | null
}

export type DataValidationBuilder = {
  requireValueInList: (
    values: string[],
    multiple?: boolean,
    showDropdown?: boolean,
  ) => DataValidationBuilder
  requireNumberBetween: (
    start: number,
    end: number,
    isInteger?: boolean,
  ) => DataValidationBuilder
  requireNumberGreaterThan: (
    num: number,
    isInteger?: boolean,
  ) => DataValidationBuilder
  requireNumberLessThan: (
    num: number,
    isInteger?: boolean,
  ) => DataValidationBuilder
  setOptions: (options: {
    allowBlank?: boolean
    showErrorMessage?: boolean
    error?: string
    showDropDown?: boolean
  }) => DataValidationBuilder
  build: () => unknown
}

export type UniverApi = {
  createWorkbook: (data?: object) => unknown
  createDocument: (data?: object) => unknown
  getActiveWorkbook: () => { save: () => unknown; getId: () => string } | null
  getActiveDocument: () => UniverDocument | null
  disposeUnit: (id: string) => void
  getHooks?: () => unknown
  newDataValidation?: () => DataValidationBuilder
}

export function useUniverHost() {
  const container = ref<HTMLElement | null>(null)
  const api = shallowRef<UniverApi | null>(null)
  const univerRef = shallowRef<{
    dispose: () => void
    __getInjector: () => { get: (token: unknown) => unknown }
  } | null>(null)
  const unitId = ref<string | null>(null)

  async function mountSheet(snapshot?: object) {
    await destroy()
    if (!container.value) return
    const { createUniver, LocaleType, mergeLocales } = await import(
      '@univerjs/presets'
    )
    const { UniverSheetsCorePreset } = await import(
      '@univerjs/preset-sheets-core'
    )
    const { UniverSheetsDrawingPreset } = await import(
      '@univerjs/preset-sheets-drawing'
    )
    const { UniverSheetsSortPreset } = await import(
      '@univerjs/preset-sheets-sort'
    )
    const { UniverSheetsFilterPreset } = await import(
      '@univerjs/preset-sheets-filter'
    )
    const { UniverSheetsConditionalFormattingPreset } = await import(
      '@univerjs/preset-sheets-conditional-formatting'
    )
    const { UniverSheetsDataValidationPreset } = await import(
      '@univerjs/preset-sheets-data-validation'
    )
    const locale = (
      await import('@univerjs/preset-sheets-core/locales/ru-RU')
    ).default
    const drawingLocale = (
      await import('@univerjs/preset-sheets-drawing/locales/ru-RU')
    ).default
    const sortLocale = (
      await import('@univerjs/preset-sheets-sort/locales/ru-RU')
    ).default
    const filterLocale = (
      await import('@univerjs/preset-sheets-filter/locales/ru-RU')
    ).default
    const conditionalFormattingLocale = (
      await import(
        '@univerjs/preset-sheets-conditional-formatting/locales/ru-RU'
      )
    ).default
    const dataValidationLocale = (
      await import('@univerjs/preset-sheets-data-validation/locales/ru-RU')
    ).default
    await import('@univerjs/preset-sheets-core/lib/index.css')
    await import('@univerjs/preset-sheets-drawing/lib/index.css')
    await import('@univerjs/preset-sheets-sort/lib/index.css')
    await import('@univerjs/preset-sheets-filter/lib/index.css')
    await import('@univerjs/preset-sheets-conditional-formatting/lib/index.css')
    await import('@univerjs/preset-sheets-data-validation/lib/index.css')

    const created = createUniver({
      locale: LocaleType.RU_RU,
      locales: {
        [LocaleType.RU_RU]: mergeLocales(
          locale,
          drawingLocale,
          sortLocale,
          filterLocale,
          conditionalFormattingLocale,
          dataValidationLocale,
        ),
      },
      presets: [
        UniverSheetsCorePreset({
          container: container.value,
        }),
        UniverSheetsDrawingPreset(),
        UniverSheetsSortPreset(),
        UniverSheetsFilterPreset(),
        UniverSheetsConditionalFormattingPreset(),
        UniverSheetsDataValidationPreset(),
      ],
    })
    const univer = created.univer as {
      dispose: () => void
      __getInjector: () => { get: (token: unknown) => unknown }
    }
    univerRef.value = univer
    api.value = created.univerAPI as unknown as UniverApi
    setSelectionInjector(univer.__getInjector())
    const wb = api.value.createWorkbook(snapshot ?? {}) as { getId: () => string }
    unitId.value = wb.getId?.() ?? null
  }

  async function mountDoc(snapshot?: object) {
    await destroy()
    if (!container.value) return
    const { createUniver, LocaleType, mergeLocales } = await import(
      '@univerjs/presets'
    )
    const { UniverDocsCorePreset } = await import('@univerjs/preset-docs-core')
    const locale = (
      await import('@univerjs/preset-docs-core/locales/ru-RU')
    ).default
    await import('@univerjs/preset-docs-core/lib/index.css')

    const created = createUniver({
      locale: LocaleType.RU_RU,
      locales: {
        [LocaleType.RU_RU]: mergeLocales(locale),
      },
      presets: [
        UniverDocsCorePreset({
          container: container.value,
        }),
      ],
    })
    const univer = created.univer as {
      dispose: () => void
      __getInjector: () => { get: (token: unknown) => unknown }
    }
    univerRef.value = univer
    api.value = created.univerAPI as unknown as UniverApi
    setSelectionInjector(univer.__getInjector())
    const doc = api.value.createDocument(snapshot ?? {}) as {
      getId: () => string
    }
    unitId.value = doc.getId?.() ?? null
  }

  function getSnapshot(type: 'SHEET' | 'DOC') {
    if (!api.value) return null
    if (type === 'SHEET') {
      return api.value.getActiveWorkbook()?.save() ?? null
    }
    return api.value.getActiveDocument()?.save() ?? null
  }

  async function destroy() {
    if (api.value && unitId.value) {
      try {
        api.value.disposeUnit(unitId.value)
      } catch {
        /* ignore */
      }
    }
    try {
      univerRef.value?.dispose?.()
    } catch {
      /* ignore */
    }
    api.value = null
    univerRef.value = null
    unitId.value = null
    setSelectionInjector(null)
    if (container.value) container.value.innerHTML = ''
  }

  onBeforeUnmount(() => {
    void destroy()
  })

  return { container, api, mountSheet, mountDoc, getSnapshot, destroy }
}
