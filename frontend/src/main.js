import { createApp } from 'vue'
import { createPinia } from 'pinia'
import {
  Document,
  Odometer,
  Plus,
  Promotion,
  Setting,
  Upload,
  User,
} from '@element-plus/icons-vue'
import App from './App.vue'
import router from './router'

const app = createApp(App)
const pinia = createPinia()

app.use(pinia)
app.use(router)

for (const [key, component] of Object.entries({
  Document,
  Odometer,
  Plus,
  Promotion,
  Setting,
  Upload,
  User,
})) {
  app.component(key, component)
}

app.mount('#app')
