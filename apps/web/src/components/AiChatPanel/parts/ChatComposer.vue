<script setup lang="ts">
import { ref } from 'vue';
import type { ChatContextUsage } from '@/ai/estimateChatContextUsage';
import ContextUsageIndicator from '@/components/ContextUsageIndicator.vue';
import type { ChatMode } from '@/univer/officeTools';
import { modeLabel } from '../formatters';
import type { ModelOption } from '../types';

defineProps<{
  draft: string;
  mode: ChatMode;
  modelLabelText: string;
  placeholder: string;
  streaming: boolean;
  attachments: Array<{ name: string; text: string }>;
  attachError: string;
  attachHelpText: string;
  contextUsage: ChatContextUsage | null;
  modeMenu: boolean;
  modelMenu: boolean;
  modelOptions: ModelOption[];
  modelMenuLoading: boolean;
  modelMenuError: string;
  providersCount: number;
  activeProviderId: string;
  activeModel: string;
}>();

const emit = defineEmits<{
  'update:draft': [value: string];
  'update:modeMenu': [value: boolean];
  'update:modelMenu': [value: boolean];
  'set-mode': [mode: ChatMode];
  'toggle-model-menu': [];
  'set-model': [option: ModelOption];
  'pick-attach': [];
  'attach-change': [event: Event];
  'remove-attach': [index: number];
  send: [];
  stop: [];
  'composer-key': [event: KeyboardEvent];
  'composer-input': [];
}>();

const composerEl = ref<HTMLTextAreaElement | null>(null);
const modeWrap = ref<HTMLElement | null>(null);
const modelWrap = ref<HTMLElement | null>(null);
const fileInput = ref<HTMLInputElement | null>(null);

defineExpose({ composerEl, modeWrap, modelWrap, fileInput });
</script>

<template>
  <div class="composer drop-up">
    <p v-if="attachError" class="attach-error">
      {{ attachError }}
    </p>
    <div v-if="attachments.length" class="context-chips">
      <button
        v-for="(file, index) in attachments"
        :key="file.name + index"
        type="button"
        class="chip on attach-chip"
        :title="file.name"
        @click="emit('remove-attach', index)"
      >
        {{ file.name }} ×
      </button>
    </div>
    <textarea
      ref="composerEl"
      :value="draft"
      rows="1"
      :placeholder="placeholder"
      :disabled="streaming"
      @input="
        emit('update:draft', ($event.target as HTMLTextAreaElement).value);
        emit('composer-input');
      "
      @keydown="emit('composer-key', $event)"
    />
    <div class="composer-bar">
      <div class="menus">
        <div ref="modeWrap" class="menu-wrap">
          <button
            type="button"
            class="pill"
            :class="mode"
            :aria-expanded="modeMenu"
            @click="
              emit('update:modeMenu', !modeMenu);
              emit('update:modelMenu', false);
            "
          >
            <span v-if="mode === 'agent'" class="glyph" aria-hidden="true">
              <svg viewBox="0 0 24 24" width="12" height="12">
                <circle cx="12" cy="12" r="3.5" fill="currentColor" />
                <path
                  d="M12 3v3M12 18v3M3 12h3M18 12h3"
                  stroke="currentColor"
                  stroke-width="1.8"
                  stroke-linecap="round"
                  fill="none"
                />
              </svg>
            </span>
            <span v-else-if="mode === 'plan'" class="glyph" aria-hidden="true">
              <svg viewBox="0 0 24 24" width="12" height="12">
                <path
                  d="M5 7h14M5 12h14M5 17h10"
                  fill="none"
                  stroke="currentColor"
                  stroke-width="1.8"
                  stroke-linecap="round"
                />
              </svg>
            </span>
            <span v-else class="glyph" aria-hidden="true">
              <svg viewBox="0 0 24 24" width="12" height="12">
                <circle
                  cx="12"
                  cy="12"
                  r="7"
                  fill="none"
                  stroke="currentColor"
                  stroke-width="1.8"
                />
              </svg>
            </span>
            <span class="pill-label">{{ modeLabel(mode) }}</span>
            <span class="caret">▾</span>
          </button>
          <div v-if="modeMenu" class="menu">
            <button type="button" @click="emit('set-mode', 'agent')">
              Агент
            </button>
            <button type="button" @click="emit('set-mode', 'plan')">
              План
            </button>
            <button type="button" @click="emit('set-mode', 'chat')">Чат</button>
          </div>
        </div>
        <div ref="modelWrap" class="menu-wrap model-wrap">
          <button
            type="button"
            class="model"
            :title="modelLabelText"
            :aria-expanded="modelMenu"
            @click="emit('toggle-model-menu')"
          >
            <span class="model-label">{{ modelLabelText }}</span>
            <span class="caret">▾</span>
          </button>
          <div v-if="modelMenu" class="menu model-menu">
            <p v-if="modelMenuLoading" class="menu-hint">Загрузка…</p>
            <p v-else-if="modelMenuError" class="menu-hint">
              {{ modelMenuError }}
            </p>
            <button
              v-for="option in modelOptions"
              :key="`${option.providerId}:${option.model}`"
              type="button"
              :class="{
                active:
                  activeProviderId === option.providerId &&
                  activeModel === option.model,
              }"
              @click="emit('set-model', option)"
            >
              <span v-if="providersCount > 1" class="menu-provider"
                >{{ option.provider }} ·
              </span>
              {{ option.model }}
            </button>
          </div>
        </div>
      </div>
      <div class="composer-actions">
        <ContextUsageIndicator :usage="contextUsage" />
        <div class="attach-wrap">
          <button
            type="button"
            class="icon-btn"
            aria-label="Вложение"
            aria-describedby="attach-help-popover"
            @click="emit('pick-attach')"
          >
            <svg viewBox="0 0 24 24" width="16" height="16" aria-hidden="true">
              <path
                fill="none"
                stroke="currentColor"
                stroke-width="1.8"
                d="M8 12.5 14.2 6.3a3 3 0 0 1 4.2 4.2l-7.4 7.4a4.2 4.2 0 0 1-6-6L12 5"
              />
            </svg>
          </button>
          <div id="attach-help-popover" class="attach-popover" role="tooltip">
            {{ attachHelpText }}
          </div>
          <input
            ref="fileInput"
            type="file"
            accept=".xlsx,.xls,.docx,.pptx,.pdf,.txt,.md,.csv,.json"
            hidden
            @change="emit('attach-change', $event)"
          />
        </div>
        <button
          type="button"
          class="send"
          :class="mode"
          :disabled="!streaming && !draft.trim()"
          @click="streaming ? emit('stop') : emit('send')"
        >
          <span v-if="streaming" class="stop-square" />
          <span v-else>↑</span>
        </button>
      </div>
    </div>
  </div>
</template>

<style scoped lang="scss">
.composer {
  position: relative;
  margin: 8px;
  border: 1px solid var(--border);
  border-radius: var(--radius-lg);
  background: var(--bg-2);
  padding: 8px;
}
.glyph {
  display: inline-flex;
  align-items: center;
  justify-content: center;
  margin-right: 2px;
}
.context-chips {
  display: flex;
  flex-wrap: wrap;
  gap: 6px;
  padding: 2px 4px 6px;
  min-height: 0;
}
.chip {
  max-width: 100%;
  border: 1px solid var(--border);
  background: transparent;
  color: var(--fg-3);
  border-radius: var(--radius-sm);
  padding: 3px 9px;
  font-size: 0.74rem;
  cursor: pointer;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}
.chip.on {
  color: var(--fg-0);
  border-color: color-mix(in srgb, var(--accent) 55%, var(--border));
  background: color-mix(in srgb, var(--accent) 12%, transparent);
}
.chip:disabled {
  opacity: 0.45;
  cursor: default;
}
.attach-chip {
  max-width: 140px;
}
.composer textarea {
  width: 100%;
  resize: none;
  border: 0;
  outline: none;
  box-shadow: none;
  background: transparent;
  color: var(--fg-0);
  padding: 4px 6px;
  min-height: 56px;
  max-height: min(40vh, 280px);
  overflow-y: auto;
  border-radius: var(--radius-sm);
}
.composer textarea:focus-visible {
  outline: none;
  box-shadow: var(--focus-ring);
}
.composer :is(button, .pill, .model, .send, .icon-btn, .chip):focus {
  outline: none;
}
.composer :is(button, .pill, .model, .send, .icon-btn, .chip):focus-visible {
  outline: none;
  box-shadow: var(--focus-ring);
}
.composer-bar {
  display: flex;
  justify-content: space-between;
  align-items: center;
  gap: 6px;
  min-width: 0;
}
.menus {
  display: flex;
  flex: 1 1 auto;
  flex-wrap: nowrap;
  align-items: center;
  gap: 6px;
  min-width: 0;
}
.menu-wrap {
  position: relative;
  display: flex;
  flex: 0 1 auto;
  max-width: 42%;
  min-width: 0;
}
.model-wrap {
  flex: 1 1 auto;
  max-width: none;
}
.pill,
.model {
  display: inline-flex;
  align-items: center;
  gap: 4px;
  width: 100%;
  max-width: 100%;
  border: 1px solid var(--border);
  background: transparent;
  color: var(--fg-2);
  border-radius: var(--radius-sm);
  padding: 3px 8px;
  cursor: pointer;
  font-size: 0.78rem;
  white-space: nowrap;
  min-width: 0;
}
.pill-label,
.model-label {
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
  min-width: 0;
}
.pill.agent {
  color: #93c5fd;
  border-color: color-mix(in srgb, #3b82f6 55%, var(--border));
}
.pill.plan {
  color: #facc15;
  border-color: color-mix(in srgb, #eab308 55%, var(--border));
}
.caret {
  flex-shrink: 0;
  color: var(--fg-3);
}
.menu {
  position: absolute;
  left: 0;
  bottom: calc(100% + 6px);
  z-index: 20;
  min-width: 160px;
  max-width: min(280px, 70vw);
  background: var(--bg-1);
  border: 1px solid var(--border);
  border-radius: var(--radius-md);
  box-shadow: var(--shadow);
  padding: 4px;
  display: grid;
}
.menu button {
  text-align: left;
  background: transparent;
  border: 0;
  color: var(--fg-0);
  padding: 8px 10px;
  border-radius: 8px;
  cursor: pointer;
}
.menu button:hover {
  background: var(--bg-hover);
}
.menu button.active {
  background: var(--bg-hover);
  color: var(--accent);
}
.menu-hint {
  margin: 0;
  padding: 8px 10px;
  color: var(--fg-3);
  font-size: 0.82rem;
}
.menu-provider {
  color: var(--fg-3);
}
.model-menu {
  left: auto;
  right: 0;
  min-width: min(220px, 70vw);
  max-height: 280px;
  overflow-y: auto;
}
.composer-actions {
  display: flex;
  flex-shrink: 0;
  align-items: center;
  gap: 4px;
}
.attach-wrap {
  position: relative;
  display: flex;
  align-items: center;
}
.attach-popover {
  position: absolute;
  right: 0;
  bottom: calc(100% + 8px);
  z-index: 6;
  width: 240px;
  padding: 10px 12px;
  background: var(--bg-1);
  border: 1px solid var(--border);
  border-radius: var(--radius-md);
  box-shadow: var(--shadow);
  color: var(--fg-2);
  font-size: 0.78rem;
  line-height: 1.4;
  pointer-events: none;
  opacity: 0;
  visibility: hidden;
  transition:
    opacity 0.12s ease,
    visibility 0.12s ease;
}
.attach-wrap:hover .attach-popover,
.attach-wrap:focus-within .attach-popover {
  opacity: 1;
  visibility: visible;
}
.icon-btn {
  box-sizing: border-box;
  display: grid;
  place-items: center;
  width: 30px;
  height: 30px;
  padding: 0;
  border: 0;
  background: transparent;
  color: var(--fg-2);
  cursor: pointer;
  border-radius: 8px;
  line-height: 0;
}
.icon-btn svg {
  display: block;
}
.icon-btn:hover {
  color: var(--fg-0);
  background: var(--bg-hover);
}
.send {
  width: 30px;
  height: 30px;
  border-radius: 50%;
  border: 0;
  cursor: pointer;
  display: grid;
  place-items: center;
  font-size: 0.95rem;
}
.send:disabled {
  background: var(--bg-3);
  color: var(--fg-3);
  cursor: default;
}
.send.agent:not(:disabled),
.send.chat:not(:disabled) {
  background: var(--accent);
  color: white;
}
.send.plan:not(:disabled) {
  background: #eab308;
  color: #111;
}
.stop-square {
  width: 10px;
  height: 10px;
  background: currentColor;
  border-radius: 2px;
}
.attach-error {
  color: var(--danger);
  font-size: 0.78rem;
  margin: 0 6px 4px;
}
</style>
