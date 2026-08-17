import axios from 'axios'

const api = axios.create({
  baseURL: '',
  timeout: 30000,
})

let refreshPromise = null

api.interceptors.request.use((config) => {
  const token = localStorage.getItem('access_token')
  if (token) {
    config.headers.Authorization = 'Bearer ' + token
  }
  return config
})

api.interceptors.response.use(
  (response) => response,
  async (error) => {
    const original = error.config || {}
    if (
      error.response?.status === 401
      && !original._retry
      && !String(original.url || '').includes('/api/auth/refresh')
    ) {
      const refreshToken = localStorage.getItem('refresh_token')
      if (refreshToken) {
        try {
          original._retry = true
          if (!refreshPromise) {
            refreshPromise = axios
              .post('/api/auth/refresh', { refresh_token: refreshToken })
              .finally(() => { refreshPromise = null })
          }
          const pendingRefresh = refreshPromise
          const res = await pendingRefresh
          localStorage.setItem('access_token', res.data.access_token)
          localStorage.setItem('refresh_token', res.data.refresh_token)
          original.headers = original.headers || {}
          original.headers.Authorization = 'Bearer ' + res.data.access_token
          return api(original)
        } catch {
          localStorage.removeItem('access_token')
          localStorage.removeItem('refresh_token')
          window.location.href = '/login'
        }
      } else {
        localStorage.removeItem('access_token')
        localStorage.removeItem('refresh_token')
        window.location.href = '/login'
      }
    }
    return Promise.reject(error)
  }
)

export default api
