/**
 * SlideHub — Система уведомлений
 */

let notificationTimeout = null;

function showNotification(message, type = 'info', duration = 3000) {
  const el = document.getElementById('notification');
  
  if (!el) return;
  
  // Очищаем предыдущий таймер
  if (notificationTimeout) {
    clearTimeout(notificationTimeout);
    notificationTimeout = null;
  }

  // Устанавливаем сообщение и класс
  el.textContent = message;
  el.className = 'notification ' + type;
  
  // Показываем
  el.classList.add('show');
  
  // Скрываем через duration
  notificationTimeout = setTimeout(() => {
    el.classList.remove('show');
    notificationTimeout = null;
  }, duration);
}

// Быстрые методы
function showSuccess(message) {
  showNotification(message, 'success');
}

function showError(message) {
  showNotification(message, 'error');
}

function showWarning(message) {
  showNotification(message, 'warning');
}

// Заглушка для совместимости с остальным кодом
window.showNotification = showNotification;