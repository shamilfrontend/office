import type { Plugin } from 'vite';
import { EDITOR_RU } from './src/slides/editor-ru.ts';

const T_FN = `function t(key) {
	if (get(locale) === "en") return key;
	return ja[key] ?? key;
}
function capLabel(cap) {
	return get(locale) === "ja" ? cap.labelJa : cap.labelEn;
}
function catLabel(label) {
	return get(locale) === "ja" ? label.ja : label.en;
}`;

const T_NEXT = `const editorRu = ${JSON.stringify(EDITOR_RU)};
function editorRuText(key) {
	return editorRu[key] ?? key;
}
function t(key) {
	if (get(locale) === "en") return key;
	if (get(locale) === "ru") return editorRuText(key);
	return ja[key] ?? key;
}
function capLabel(cap) {
	if (get(locale) === "ja") return cap.labelJa;
	if (get(locale) === "ru") return editorRuText(cap.labelEn);
	return cap.labelEn;
}
function catLabel(label) {
	if (get(locale) === "ja") return label.ja;
	if (get(locale) === "ru") return editorRuText(label.en);
	return label.en;
}`;

const LOCALES = `const LOCALES = [{
	id: "en",
	label: "English"
}, {
	id: "ja",
	label: "日本語"
}];`;

const LOCALES_NEXT = `const LOCALES = [{
	id: "ru",
	label: "Русский"
}, {
	id: "en",
	label: "English"
}, {
	id: "ja",
	label: "日本語"
}];`;

const MENUS = `const MENUS_WITHOUT_FEEDBACK = {
	en: withoutFeedback(NATIVE_MENUBAR.en),
	ja: withoutFeedback(NATIVE_MENUBAR.ja)
};
function nativeMenus(editor) {
	return (editor.feedbackUrl === void 0 ? MENUS_WITHOUT_FEEDBACK : NATIVE_MENUBAR)[getLocale()];
}`;

const MENUS_NEXT = `function translateEditorMenu(entries) {
	const mapped = [];
	for (const entry of entries) {
		if (entry !== "-" && entry.id === "file/share") continue;
		if (entry === "-") {
			mapped.push(entry);
			continue;
		}
		const label = entry.label === "Edit" ? "Правка" : editorRuText(entry.label);
		const next = { ...entry, label };
		if (entry.items) next.items = translateEditorMenu(entry.items);
		if (entry.children) next.children = translateEditorMenu(entry.children);
		mapped.push(next);
	}
	return mapped.filter((entry, index) => {
		if (entry !== "-") return true;
		const prev = mapped[index - 1];
		const next = mapped[index + 1];
		return prev !== "-" && next != null && next !== "-";
	});
}
const MENUS_WITHOUT_FEEDBACK = {
	en: withoutFeedback(NATIVE_MENUBAR.en),
	ja: withoutFeedback(NATIVE_MENUBAR.ja)
};
NATIVE_MENUBAR.ru = translateEditorMenu(NATIVE_MENUBAR.en);
MENUS_WITHOUT_FEEDBACK.ru = withoutFeedback(NATIVE_MENUBAR.ru);
function nativeMenus(editor) {
	const menus = editor.feedbackUrl === void 0 ? MENUS_WITHOUT_FEEDBACK : NATIVE_MENUBAR;
	return menus[getLocale()] ?? menus.en;
}`;

const APPLY = 'label: APPLY_TO_DEFAULTS[getLocale()]';
const APPLY_NEXT =
  'label: getLocale() === "ru" ? editorRuText(APPLY_TO_DEFAULTS.en) : APPLY_TO_DEFAULTS[getLocale()]';

/** The published editor only ships en/ja. Inject Russian at build time. */
export function pptxEditorRu(): Plugin {
  return {
    name: 'pptx-editor-ru',
    enforce: 'pre',
    transform(code, id) {
      if (
        !id.includes('@office-kit/pptx-editor') ||
        !code.includes('function t(key)')
      ) {
        return null;
      }
      const missing = [T_FN, LOCALES, MENUS, APPLY].filter(
        (needle) => !code.includes(needle),
      );
      if (missing.length) {
        this.error(
          '@office-kit/pptx-editor i18n changed; the Russian patch no longer matches.',
        );
      }
      const next = code
        .replace(T_FN, T_NEXT)
        .replace(LOCALES, LOCALES_NEXT)
        .replace(MENUS, MENUS_NEXT)
        .replace(APPLY, APPLY_NEXT);
      return { code: next, map: null };
    },
  };
}
