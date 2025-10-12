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
      path: '/:pathMatch(.*)*',
      redirect: '/'
    }
  ]
})

router.beforeEach((to, _from, next) => {
  const token = typeof localStorage !== 'undefined' ? localStorage.getItem('authToken') : null
  if (!token && to.path !== '/login') {
    next('/login')
    return
  }
  next()
})

export default router
