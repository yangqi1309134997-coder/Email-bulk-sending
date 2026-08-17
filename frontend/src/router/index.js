import { createRouter, createWebHistory } from 'vue-router'
import { useAuthStore } from '../stores/auth'

const routes = [
  {
    path: '/login',
    name: 'Login',
    component: () => import('../views/Login.vue'),
  },
  {
    path: '/register',
    name: 'Register',
    component: () => import('../views/Register.vue'),
  },
  {
    path: '/',
    component: () => import('../views/Dashboard.vue'),
    meta: { requiresAuth: true },
    children: [
      { path: '', redirect: '/dashboard' },
      { path: 'dashboard', name: 'DashboardHome', component: () => import('../views/DashboardHome.vue') },
      { path: 'send', name: 'SendTask', component: () => import('../views/SendTask.vue') },
      { path: 'tasks/:id', name: 'TaskDetail', component: () => import('../views/TaskDetail.vue') },
      { path: 'templates', name: 'Templates', component: () => import('../views/Templates.vue') },
      { path: 'settings', name: 'Settings', component: () => import('../views/Settings.vue') },
      { path: 'users', name: 'Users', component: () => import('../views/Users.vue'), meta: { requiresAdmin: true } },
    ],
  },
]

const router = createRouter({
  history: createWebHistory(),
  routes,
})

router.beforeEach(async (to) => {
  const authStore = useAuthStore()
  const requiresAuth = to.matched.some((route) => route.meta.requiresAuth)
  const requiresAdmin = to.matched.some((route) => route.meta.requiresAdmin)

  if (requiresAuth && !authStore.isLoggedIn) return '/login'
  if (authStore.isLoggedIn && !authStore.user) {
    await authStore.fetchUser()
  }
  if (requiresAuth && !authStore.isLoggedIn) return '/login'
  if (requiresAdmin && authStore.user?.role !== 'admin') return '/dashboard'
  if (to.path === '/login' && authStore.isLoggedIn) return '/dashboard'
  return true
})

export default router
