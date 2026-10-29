import { onBeforeUnmount, ref, shallowRef, type Ref } from 'vue';
import { SLIDE_TOOL_NAMES, type OfficeToolDef } from '@office/shared';
import type {
  EditorHandle,
  EditorLocale,
  EditorTool,
  ShapeRef,
} from '@office-kit/pptx-editor';

type MountHooks = {
  onDirty: () => void;
  onSave: (pptx: Uint8Array) => Promise<void>;
};

function toOfficeTool(tool: EditorTool): OfficeToolDef {
  return {
    type: 'function',
    function: {
      name: tool.name,
      description: tool.description,
      parameters: tool.input_schema as Record<string, unknown>,
    },
  };
}

function collectShadowRoots(node: Element | ShadowRoot, roots: ShadowRoot[]) {
  if (node instanceof ShadowRoot) roots.push(node);
  const shadow = node instanceof Element ? node.shadowRoot : null;
  if (shadow) collectShadowRoots(shadow, roots);
  Array.from(node.children).forEach((child) => {
    collectShadowRoots(child, roots);
  });
}

function tuneEditorChrome(host: HTMLElement) {
  const roots: ShadowRoot[] = [];
  collectShadowRoots(host, roots);
  roots.forEach((root) => {
    if (root.querySelector('style[data-office-editor-chrome]')) return;
    const style = document.createElement('style');
    style.dataset.officeEditorChrome = '';
    style.textContent = [
      '.share-anchor { display: none !important; }',
      '.filename { display: none !important; }',
      'button.qat:has(path[d^="M2.5 2.5h9"]) { display: none !important; }',
      '.right { display: none !important; }',
      'button.comments { display: none !important; }',
      'button.labelled:has(path[d^="M1.5 1.5h12v8"]) { display: none !important; }',
      '.cluster[aria-label="Комментарии"] { display: none !important; }',
    ].join('\n');
    root.append(style);
  });
}

const EDITOR_LIGHT = `
.ok-editor {
  color-scheme: light !important;
  --ok-accent: #b4235a !important;
  --ok-accent-2: #d6336c !important;
  --ok-bg: #f4f5f7 !important;
  --ok-panel: #ffffff !important;
  --ok-panel-2: #f9fafb !important;
  --ok-ribbon: #f7f8fa !important;
  --ok-ribbon-active: #ffffff !important;
  --ok-border: #dfe2e7 !important;
  --ok-border-strong: #c9cdd6 !important;
  --ok-text: #15171c !important;
  --ok-text-2: #545b69 !important;
  --ok-text-3: #8a909c !important;
  --ok-hover: #eef0f3 !important;
  --ok-selected: #fdebf1 !important;
  --ok-selected-border: #b4235a !important;
  --ok-canvas-bg: #d3d6dc !important;
  --ok-danger: #c92a2a !important;
  --ok-shadow: 0 1px 2px rgba(16, 18, 23, 0.08), 0 0 0 1px rgba(16, 18, 23, 0.04) !important;
  --ok-shadow-lg: 0 12px 32px -8px rgba(16, 18, 23, 0.22), 0 2px 6px rgba(16, 18, 23, 0.08) !important;
}`;

const EDITOR_DARK = `
.ok-editor {
  color-scheme: dark !important;
  --ok-accent: #e85a90 !important;
  --ok-accent-2: #ff9cbf !important;
  --ok-bg: #1f1f1f !important;
  --ok-panel: #2b2b2b !important;
  --ok-panel-2: #252525 !important;
  --ok-ribbon: #2a2a2a !important;
  --ok-ribbon-active: #333333 !important;
  --ok-border: #3d3d3d !important;
  --ok-border-strong: #555555 !important;
  --ok-text: #f0f0f0 !important;
  --ok-text-2: #c4c4c4 !important;
  --ok-text-3: #8f8f8f !important;
  --ok-hover: #3a3a3a !important;
  --ok-selected: #3a2230 !important;
  --ok-selected-border: #ff7aa8 !important;
  --ok-canvas-bg: #1b1b1b !important;
  --ok-danger: #f1707a !important;
  --ok-shadow: 0 1px 2px rgba(0, 0, 0, 0.45), 0 0 0 1px rgba(255, 255, 255, 0.04) !important;
  --ok-shadow-lg: 0 12px 32px -8px rgba(0, 0, 0, 0.6), 0 2px 6px rgba(0, 0, 0, 0.35) !important;
}`;

function applyEditorTheme(host: HTMLElement) {
  const css =
    document.documentElement.dataset.theme === 'dark'
      ? EDITOR_DARK
      : EDITOR_LIGHT;
  const roots: ShadowRoot[] = [];
  collectShadowRoots(host, roots);
  roots.forEach((root) => {
    const current = root.querySelector('style[data-office-editor-theme]');
    const style = current instanceof HTMLStyleElement ? current : null;
    if (style) {
      style.textContent = css;
      return;
    }
    const next = document.createElement('style');
    next.dataset.officeEditorTheme = '';
    next.textContent = css;
    root.append(next);
  });
}

export function formatSlideSelection(selection: readonly ShapeRef[]): string {
  if (!selection.length) return 'Выделение: ничего не выбрано';
  return `Выделение фигур: ${JSON.stringify(selection)}`;
}

export function useSlideEditor(container: Ref<HTMLElement | null>) {
  const handle = shallowRef<EditorHandle | null>(null);
  const toolDefs = ref<OfficeToolDef[]>([]);
  const ready = ref(false);
  let stops: Array<() => void> = [];
  let themeObserver: MutationObserver | null = null;

  async function destroy() {
    themeObserver?.disconnect();
    themeObserver = null;
    stops.forEach((stop) => stop());
    stops = [];
    handle.value?.destroy();
    handle.value = null;
    toolDefs.value = [];
    ready.value = false;
  }

  async function mount(bytes: Uint8Array, fileName: string, hooks: MountHooks) {
    await destroy();
    const host = container.value;
    if (!host) return;
    const { mountEditor } = await import('@office-kit/pptx-editor');
    const editor = mountEditor(host, {
      source: bytes,
      fileName: fileName.toLowerCase().endsWith('.pptx')
        ? fileName
        : `${fileName}.pptx`,
      locale: 'ru' as EditorLocale,
      compact: true,
      onSave: hooks.onSave,
    });
    handle.value = editor;
    await editor.ready;
    if (handle.value !== editor) return;
    tuneEditorChrome(host);
    applyEditorTheme(host);
    themeObserver?.disconnect();
    themeObserver = new MutationObserver(() => applyEditorTheme(host));
    themeObserver.observe(document.documentElement, {
      attributes: true,
      attributeFilter: ['data-theme'],
    });
    ready.value = true;
    stops = [
      editor.on('change', () => hooks.onDirty()),
      editor.on('dirtychange', ({ dirty }) => {
        if (dirty) hooks.onDirty();
      }),
    ];
    try {
      const listed = await editor.tools();
      if (handle.value !== editor) return;
      toolDefs.value = listed
        .filter((tool) => SLIDE_TOOL_NAMES.has(tool.name))
        .map(toOfficeTool);
    } catch (error) {
      console.error(error);
    }
  }

  async function snapshot(): Promise<Uint8Array | null> {
    if (!handle.value) return null;
    return handle.value.snapshot();
  }

  async function save() {
    if (!handle.value) return;
    await handle.value.save();
  }

  async function open(bytes: Uint8Array) {
    if (!handle.value) return;
    await handle.value.open(bytes, { unsaved: true });
  }

  async function run(name: string, input: unknown) {
    if (!handle.value) throw new Error('Редактор презентации ещё не готов');
    return handle.value.run(name, input);
  }

  function selectionText() {
    const editor = handle.value;
    if (!editor) return '';
    return formatSlideSelection(editor.selection());
  }

  onBeforeUnmount(() => {
    void destroy();
  });

  return {
    ready,
    toolDefs,
    mount,
    destroy,
    snapshot,
    save,
    open,
    run,
    selectionText,
  };
}
