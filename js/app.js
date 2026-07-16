/**
 * SlideHub — Главный файл приложения
 * Инициализация Office, обработчики событий
 */

let isInitialized = false;

Office.onReady((info) => {
  if (info.host === Office.HostType.PowerPoint) {
    console.log('SlideHub: Инициализация...');
    
    // Инициализируем модули
    initApp();
  } else {
    document.getElementById('catalogContainer').innerHTML = `
      <div class="empty-message">
        <span class="empty-icon">⚠️</span>
        Пожалуйста, откройте это приложение внутри PowerPoint
      </div>
    `;
  }
});

function initApp() {
  if (isInitialized) return;
  isInitialized = true;

  // Загружаем каталог
  loadCatalog();

  // Поиск с debounce
  let searchTimeout;
  document.getElementById('searchInput').addEventListener('input', () => {
    clearTimeout(searchTimeout);
    searchTimeout = setTimeout(renderCatalog, 250);
  });

  // Фильтры
  document.getElementById('filters').addEventListener('click', (e) => {
    if (e.target.classList.contains('filter-btn')) {
      document.querySelectorAll('.filter-btn').forEach(b => b.classList.remove('active'));
      e.target.classList.add('active');
      activeFilter = e.target.dataset.category;
      
      // Аналитика
      trackEvent('filter', { category: activeFilter });
      
      renderCatalog();
    }
  });

  // Обновление каталога
  document.getElementById('refreshBtn').addEventListener('click', () => {
    trackEvent('refresh', {});
    loadCatalog(true); // Принудительная перезагрузка
    showNotification('Каталог обновлён', 'success');
  });

  // История
  document.getElementById('historyBtn').addEventListener('click', toggleHistory);

  // Глобальная обработка ошибок
  window.addEventListener('error', (e) => {
    console.error('SlideHub Error:', e.error);
    showNotification('Произошла ошибка. Проверьте консоль.', 'error');
  });

  console.log('SlideHub: Инициализация завершена');
}

function toggleHistory() {
  const panel = document.getElementById('historyPanel');
  const btn = document.getElementById('historyBtn');
  
  if (panel.style.display === 'none') {
    panel.style.display = 'block';
    btn.textContent = '✕';
    renderHistory();
  } else {
    panel.style.display = 'none';
    btn.textContent = '📊';
  }
}