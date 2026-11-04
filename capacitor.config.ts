import type { CapacitorConfig } from '@capacitor/cli';

const config: CapacitorConfig = {
  appId: 'com.coletaflow.app',
  appName: 'ColetaFlow',
  webDir: 'dist',
  ...(process.env.CF_ANDROID_LOCAL === 'true' ? {} : {
    server: {
      url: 'https://coletaflow-two.vercel.app',
      cleartext: false,
      allowNavigation: ['coletaflow-two.vercel.app']
    }
  }),
  android: {
    allowMixedContent: false
  },
  plugins: {
    SplashScreen: {
      launchShowDuration: 1800,
      launchAutoHide: true,
      backgroundColor: '#ffffff',
      showSpinner: false
    }
  }
};

export default config;
