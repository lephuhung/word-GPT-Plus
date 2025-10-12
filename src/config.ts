export interface AppConfig {
  loginUrl: string
}

// Default config; can be extended to read from env or localStorage
export const appConfig: AppConfig = {
  loginUrl:
    (typeof localStorage !== 'undefined' &&
      localStorage.getItem('loginUrl')) ||
    '/api/login'
}

export function setLoginUrl(url: string) {
  if (typeof localStorage !== 'undefined') {
    localStorage.setItem('loginUrl', url)
  }
  appConfig.loginUrl = url
}