/**
 * SlideHub — Вставка слайдов с генерацией PPTX
 */

async function insertSlide(slideId) {
  const template = slideLibrary.find(s => s.id === slideId);
  if (!template) {
    throw new Error('Шаблон не найден');
  }

  try {
    // Показываем уведомление о генерации
    showNotification('🔄 Генерируем слайд...', 'info', 2000);

    // Генерируем PPTX
    const arrayBuffer = await generateSlidePPTX(template);
    const base64 = arrayBufferToBase64(arrayBuffer);

    // Вставляем в PowerPoint
    await PowerPoint.run(async (context) => {
      const presentation = context.presentation;
      
      // Добавляем новый слайд
      const slide = presentation.slides.add();
      
      // Вставляем сгенерированный файл
      slide.insertFileFromBase64(base64);
      
      await context.sync();
      
      // Сохраняем в историю
      addToHistory(slideId);
      
      // Отправляем аналитику
      trackEvent('insert', { 
        slideId, 
        slideTitle: template.title,
        category: template.category 
      });
    });

    showNotification(`✅ Слайд "${template.title}" вставлен!`, 'success');

  } catch (error) {
    console.error('Ошибка вставки:', error);
    showNotification('❌ Не удалось вставить слайд', 'error');
    throw error;
  }
}

// Вспомогательные функции
function arrayBufferToBase64(buffer) {
  const bytes = new Uint8Array(buffer);
  let binary = '';
  for (let i = 0; i < bytes.length; i++) {
    binary += String.fromCharCode(bytes[i]);
  }
  return window.btoa(binary);
}

function addToHistory(slideId) {
  try {
    const history = JSON.parse(localStorage.getItem('slidehub_history') || '[]');
    const filtered = history.filter(h => h.slideId !== slideId);
    filtered.unshift({ 
      slideId, 
      timestamp: Date.now(),
      date: new Date().toLocaleString()
    });
    localStorage.setItem('slidehub_history', JSON.stringify(filtered.slice(0, 50)));
  } catch (error) {
    console.warn('Ошибка сохранения истории:', error);
  }
}