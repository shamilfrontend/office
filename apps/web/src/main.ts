import { createApp } from 'vue'
import { createPinia } from 'pinia'
import App from './App.vue'
import router from './router'
import { initTheme } from './composables/useTheme'
import './styles/tokens.scss'

initTheme()
createApp(App).use(createPinia()).use(router).mount('#app')
