import { createRouter, createWebHistory } from 'vue-router';
import { useAuthStore } from '@/stores/auth';

const router = createRouter({
  history: createWebHistory(),
  routes: [
    {
      path: '/',
      name: 'landing',
      component: () => import('@/views/LandingView.vue'),
      meta: { guest: true },
    },
    {
      path: '/auth/callback',
      component: () => import('@/views/AuthCallbackView.vue'),
    },
    {
      path: '/home',
      name: 'home',
      component: () => import('@/views/HomeView.vue'),
      meta: { auth: true },
    },
    {
      path: '/edit/:id',
      name: 'edit',
      component: () => import('@/views/EditorView.vue'),
      meta: { auth: true },
    },
  ],
});

router.beforeEach(async (to) => {
  const auth = useAuthStore();
  if (!auth.loaded) await auth.init();

  if (to.meta.auth && !auth.isAuthenticated) {
    return { name: 'landing' };
  }
  if (to.meta.guest && auth.isAuthenticated && to.name === 'landing') {
    return { name: 'home' };
  }
  return true;
});

export default router;
