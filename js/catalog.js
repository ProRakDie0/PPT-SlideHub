/**
 * SlideHub — Управление каталогом
 * Загрузка, фильтрация, рендеринг
 */

let slideLibrary = [];
let activeFilter = 'all';
const CACHE_KEY = 'slidehub_catalog';
const CACHE_TIME_KEY = 'slidehub_catalog_time';
const CACHE_TTL = 3600000; // 1 час

// Загрузка каталога
async function loadCatalog(forceRefresh = false) {
  const container = document.getElementById('catalogContainer');
  container.innerHTML = '<div class="empty-message">⏳ Загрузка каталога...</div>';

  try {
    // Проверяем кэш
    if (!forceRefresh) {
      const cached = localStorage.getItem(CACHE_KEY);
      const cacheTime = localStorage.getItem(CACHE_TIME_KEY);
      
      if (cached && cacheTime && (Date.now() - parseInt(cacheTime) < CACHE_TTL)) {
        slideLibrary = JSON.parse(cached);
        renderCatalog();
        updateStats();
        console.log('SlideHub: Загружено из кэша');
        return;
      }
    }

    // Загружаем из сети
    const response = await fetch('data/slides.json');
    if (!response.ok) throw new Error(`HTTP ${response.status}`);
    
    slideLibrary = await response.json();
    
    // Сохраняем в кэш
    localStorage.setItem(CACHE_KEY, JSON.stringify(slideLibrary));
    localStorage.setItem(CACHE_TIME_KEY, String(Date.now()));
    
    renderCatalog();
    updateStats();
    console.log('SlideHub: Загружено из сети');
    
  } catch (error) {
    console.error('SlideHub: Ошибка загрузки каталога:', error);
    
    // Пробуем восстановить из кэша даже если он устарел
    const cached = localStorage.getItem(CACHE_KEY);
    if (cached) {
      slideLibrary = JSON.parse(cached);
      renderCatalog();
      updateStats();
      showNotification('Загружена кэшированная версия каталога', 'warning');
    } else {
      container.innerHTML = `
        <div class="empty-message">
          <span class="empty-icon">❌</span>
          Не удалось загрузить каталог<br>
          <small style="color:#999;">Проверьте подключение к интернету</small>
        </div>
      `;
      showNotification('Ошибка загрузки каталога', 'error');
    }
  }
}

// Рендеринг каталога
function renderCatalog() {
  const searchTerm = document.getElementById('searchInput').value.toLowerCase().trim();
  const container = document.getElementById('catalogContainer');

  // Фильтрация
  const filtered = slideLibrary.filter(item => {
    const matchSearch = !searchTerm || 
      item.title.toLowerCase().includes(searchTerm) ||
      item.description.toLowerCase().includes(searchTerm) ||
      (item.tags && item.tags.some(t => t.toLowerCase().includes(searchTerm)));
    
    const matchCategory = activeFilter === 'all' || item.category === activeFilter;
    
    return matchSearch && matchCategory;
  });

  // Обновляем статистику
  document.getElementById('shownCount').textContent = filtered.length;

  if (filtered.length === 0) {
    container.innerHTML = `
      <div class="empty-message">
        <span class="empty-icon">🔍</span>
        ${searchTerm ? 'Ничего не найдено по вашему запросу' : 'Нет слайдов в этой категории'}
        ${searchTerm ? `<br><small style="color:#999;">Попробуйте изменить поисковый запрос</small>` : ''}
      </div>
    `;
    return;
  }

  container.innerHTML = filtered.map(item => `
    <div class="card" data-id="${item.id}">
      <div class="preview ${item.previewClass || ''}">
        ${item.thumbnail ? `<img src="${item.thumbnail}" alt="${item.title}" loading="lazy">` : '📄'}
        <span class="preview-label">${getCategoryLabel(item.category)}</span>
      </div>
      <div class="info">
        <h3>${escapeHtml(item.title)}</h3>
        <span class="category-badge">${getCategoryLabel(item.category)}</span>
        <p>${escapeHtml(item.description)}</p>
      </div>
      <button class="insert-btn" data-id="${item.id}">Вставить</button>
    </div>
  `).join('');

  // Навешиваем обработчики на кнопки вставки
  document.querySelectorAll('.insert-btn').forEach(btn => {
    btn.addEventListener('click', async (e) => {
      const slideId = parseInt(btn.dataset.id);
      const slide = slideLibrary.find(s => s.id === slideId);
      
      if (!slide) return;
      
      // Показываем состояние загрузки
      btn.textContent = '⏳';
      btn.classList.add('loading');
      
      try {
        await insertSlide(slideId);
        trackEvent('insert', { slideId, slideTitle: slide.title });
        showNotification(`Слайд "${slide.title}" вставлен!`, 'success');
      } catch (error) {
        console.error('Ошибка вставки:', error);
        showNotification('Не удалось вставить слайд', 'error');
      } finally {
        btn.textContent = 'Вставить';
        btn.classList.remove('loading');
      }
    });
  });
}

// Обновление статистики
function updateStats() {
  document.getElementById('totalCount').textContent = slideLibrary.length;
  document.getElementById('shownCount').textContent = slideLibrary.length;
}

// Вспомогательные функции
function getCategoryLabel(category) {
  const labels = {
    'title': 'Титульный',
    'content': 'Контент',
    'chart': 'Аналитика',
    'two-col': 'Структура'
  };
  return labels[category] || category;
}

function escapeHtml(text) {
  const div = document.createElement('div');
  div.textContent = text;
  return div.innerHTML;
}