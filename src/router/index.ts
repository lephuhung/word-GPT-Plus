import { createMemoryHistory, createRouter } from 'vue-router'

const router = createRouter({
  history: createMemoryHistory(),
  routes: [
    {
      path: '/',
      name: 'Home',
      component: () => import('../pages/HomePage.vue')
    },
    {
      path: '/login',
      name: 'Login',
      component: () => import('../pages/LoginPage.vue')
    },
    {
      path: '/settings',
      name: 'Settings',
      component: () => import('../pages/SettingsPage.vue')
    },
    {
      path: '/user',
      name: 'User',
      component: () => import('../pages/UserPage.vue')
    },
    {
      path: '/:pathMatch(.*)*',
      redirect: '/'
    }
  ]
})

router.beforeEach((to, _from, next) => {
  // Bypass login ở chế độ dev hoặc khi bật cờ ép bypass
  const isDev = import.meta.env.DEV || import.meta.env.MODE === 'development'
  const forceBypass =
    typeof localStorage !== 'undefined' && localStorage.getItem('DEV_BYPASS_LOGIN') === '1'
  if (isDev || forceBypass) {
    next()
    return
  }
  const token = typeof localStorage !== 'undefined' ? localStorage.getItem('authToken') : null
  if (!token && to.path !== '/login') {
    next('/login')
    return
  }
  next()
})

export default router
