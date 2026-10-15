import { computed, onBeforeUnmount, watch, type Ref } from 'vue';

export function useFloatingMenus(options: {
  modeMenu: Ref<boolean>;
  modelMenu: Ref<boolean>;
  historyOpen: Ref<boolean>;
  planOpen: Ref<boolean>;
  modeWrap: () => HTMLElement | null | undefined;
  modelWrap: () => HTMLElement | null | undefined;
  historyPanel: () => HTMLElement | null | undefined;
  historyBtn: () => HTMLElement | null | undefined;
  planModal: () => HTMLElement | null | undefined;
}) {
  const {
    modeMenu,
    modelMenu,
    historyOpen,
    planOpen,
    modeWrap,
    modelWrap,
    historyPanel,
    historyBtn,
    planModal,
  } = options;

  function closeFloatingOnOutside(event: PointerEvent) {
    const { target } = event;
    if (!(target instanceof Node)) return;

    if (modeMenu.value && !modeWrap()?.contains(target)) {
      modeMenu.value = false;
    }
    if (modelMenu.value && !modelWrap()?.contains(target)) {
      modelMenu.value = false;
    }
    if (
      historyOpen.value &&
      !historyPanel()?.contains(target) &&
      !historyBtn()?.contains(target)
    ) {
      historyOpen.value = false;
    }
    if (planOpen.value && !planModal()?.contains(target)) {
      planOpen.value = false;
    }
  }

  function onFloatingKeydown(event: KeyboardEvent) {
    if (event.key !== 'Escape') return;
    if (planOpen.value) {
      planOpen.value = false;
      return;
    }
    if (modeMenu.value || modelMenu.value || historyOpen.value) {
      modeMenu.value = false;
      modelMenu.value = false;
      historyOpen.value = false;
    }
  }

  const floatingUiOpen = computed(
    () =>
      modeMenu.value || modelMenu.value || historyOpen.value || planOpen.value,
  );

  watch(
    floatingUiOpen,
    (open, _prev, onCleanup) => {
      if (!open) return;
      document.addEventListener('pointerdown', closeFloatingOnOutside, true);
      document.addEventListener('keydown', onFloatingKeydown);
      onCleanup(() => {
        document.removeEventListener(
          'pointerdown',
          closeFloatingOnOutside,
          true,
        );
        document.removeEventListener('keydown', onFloatingKeydown);
      });
    },
    { flush: 'post' },
  );

  onBeforeUnmount(() => {
    document.removeEventListener('pointerdown', closeFloatingOnOutside, true);
    document.removeEventListener('keydown', onFloatingKeydown);
  });
}
