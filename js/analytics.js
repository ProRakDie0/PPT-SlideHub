/**
 * SlideHub — Аналитика использования
 * Отслеживание действий пользователя
 */

const ANALYTICS_ENABLED = true;
const ANALYTICS_ENDPOINT = ''; // URL для отправки данных (опционально)

function trackEvent(eventName, data = {}) {
  if (!ANALYTICS_ENABLED) return;
  
  const event = {
    event: eventName,
    timestamp: Date.now(),
    sessionId: getSessionId(),
    data: data,
    userAgent: navigator.userAgent,
    host: Office && Office.context ? Office.context.host : 'unknown'
  };
  
  console.log('[Analytics]', event);
  
  // Сохраняем локально
  saveAnalyticsEvent(event);
  
  // Отправляем на сервер (если настроен)
  if (ANALYTICS_ENDPOINT) {
    sendAnalyticsEvent(event);
  }
}

function getSessionId() {
  let sessionId = localStorage.getItem('slidehub_session_id');
  if (!sessionId) {
    sessionId = 'session_' + Date.now() + '_' + Math.random().toString(36).substr(2, 9);
    localStorage.setItem('slidehub_session_id', sessionId);
  }
  return sessionId;
}

function saveAnalyticsEvent(event) {
  try {
    const events = JSON.parse(localStorage.getItem('slidehub_analytics') || '[]');
    events.push(event);
    // Оставляем только последние 1000 событий
    if (events.length > 1000) {
      events.splice(0, events.length - 1000);
    }
    localStorage.setItem('slidehub_analytics', JSON.stringify(events));
  } catch (error) {
    console.warn('Ошибка сохранения аналитики:', error);
  }
}

async function sendAnalyticsEvent(event) {
  if (!ANALYTICS_ENDPOINT) return;
  
  try {
    await fetch(ANALYTICS_ENDPOINT, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(event)
    });
  } catch (error) {
    console.warn('Ошибка отправки аналитики:', error);
  }
}

// Экспорт в глобальную область
window.trackEvent = trackEvent;