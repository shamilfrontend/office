import { computed, ref } from 'vue';

const WIDTH_KEY = 'office_chat_width';

function readWidth() {
  const n = Number(localStorage.getItem(WIDTH_KEY));
  return Number.isFinite(n) && n >= 320 ? n : 420;
}

export function usePanelResize(isCollapsed: () => boolean) {
  const width = ref(readWidth());

  const panelStyle = computed(() =>
    isCollapsed() ? undefined : { width: `${width.value}px` },
  );

  function startResize(event: PointerEvent) {
    if (isCollapsed()) return;
    event.preventDefault();
    const startX = event.clientX;
    const startW = width.value;
    const move = (ev: PointerEvent) => {
      const max = Math.max(320, Math.floor(window.innerWidth * 0.5));
      width.value = Math.min(
        max,
        Math.max(320, startW + (startX - ev.clientX)),
      );
    };
    const up = () => {
      localStorage.setItem(WIDTH_KEY, String(Math.round(width.value)));
      window.removeEventListener('pointermove', move);
      window.removeEventListener('pointerup', up);
    };
    window.addEventListener('pointermove', move);
    window.addEventListener('pointerup', up);
  }

  return {
    width,
    panelStyle,
    startResize,
  };
}
