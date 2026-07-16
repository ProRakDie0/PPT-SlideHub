/**
 * SlideHub — Генератор PowerPoint слайдов
 * Создает .pptx файлы прямо в браузере
 */

function generateSlidePPTX(template) {
  return new Promise((resolve, reject) => {
    try {
      // Создаем презентацию
      const pptx = new PptxGenJS();
      pptx.defineLayout({ name: 'WIDE', width: 13.33, height: 7.5 });
      pptx.layout = 'WIDE';
      
      // Настройки
      pptx.author = 'SlideHub';
      pptx.company = 'SlideHub Team';
      pptx.title = template.title;

      // Создаем слайд
      const slide = pptx.addSlide();

      // ============================================
      // ФОН
      // ============================================
      slide.background = { fill: '#F8F9FA' };

      // ============================================
      // ДЕКОРАТИВНАЯ ПОЛОСА СВЕРХУ
      // ============================================
      slide.addShape(pptx.ShapeType.rect, {
        x: 0,
        y: 0,
        w: 13.33,
        h: 0.15,
        fill: { color: getCategoryColor(template.category) },
        rectRadius: 0,
      });

      // ============================================
      // ЗАГОЛОВОК
      // ============================================
      slide.addText(template.content.title || template.title, {
        x: 0.8,
        y: 0.5,
        w: 11.5,
        h: 1.0,
        fontSize: 32,
        fontFace: 'Segoe UI',
        color: '1A202C',
        bold: true,
        align: 'left',
      });

      // Линия под заголовком
      slide.addShape(pptx.ShapeType.rect, {
        x: 0.8,
        y: 1.6,
        w: 2.5,
        h: 0.06,
        fill: { color: getCategoryColor(template.category) },
      });

      // ============================================
      // ГЕНЕРАЦИЯ КОНТЕНТА ПО ТИПУ
      // ============================================
      let yPos = 2.1;

      // --- АНАЛИЗ АНАЛОГОВ ---
      if (template.category === 'analytics' && template.content.sections) {
        yPos = generateAnalogsContent(slide, template.content, yPos);
      }
      
      // --- БЮДЖЕТ ---
      else if (template.category === 'finance' && template.content.budget) {
        yPos = generateBudgetContent(slide, template.content, yPos);
      }
      
      // --- UNIT-ЭКОНОМИКА ---
      else if (template.category === 'analytics' && template.content.metrics) {
        yPos = generateMetricsContent(slide, template.content, yPos);
      }
      
      // --- ROADMAP ---
      else if (template.category === 'roadmap' && template.content.phases) {
        yPos = generateRoadmapContent(slide, template.content, yPos);
      }
      
      // --- КОМАНДА ---
      else if (template.category === 'team' && template.content.members) {
        yPos = generateTeamContent(slide, template.content, yPos);
      }
      
      // --- РАСПРЕДЕЛЕНИЕ БЮДЖЕТА ---
      else if (template.category === 'summary' && template.content.distribution) {
        yPos = generateDistributionContent(slide, template.content, yPos);
      }
      
      // --- БУЛЛЕТЫ (базовый режим) ---
      else if (template.content.bullets) {
        yPos = generateBulletsContent(slide, template.content.bullets, yPos);
      }

      // ============================================
      // НОМЕР СЛАЙДА (внизу справа)
      // ============================================
      slide.addText('SlideHub', {
        x: 11.5,
        y: 7.0,
        w: 1.5,
        h: 0.4,
        fontSize: 10,
        color: 'A0AEC0',
        align: 'right',
        fontFace: 'Segoe UI',
      });

      // Генерируем файл
      pptx.writeFile({ outputType: 'arraybuffer' })
        .then((arrayBuffer) => {
          resolve(arrayBuffer);
        })
        .catch((err) => {
          reject(err);
        });

    } catch (error) {
      reject(error);
    }
  });
}

// ============================================
// ГЕНЕРАТОРЫ КОНТЕНТА
// ============================================

// 1. АНАЛИЗ АНАЛОГОВ
function generateAnalogsContent(slide, content, startY) {
  let y = startY;
  const sections = content.sections || [];
  const width = 11.0;
  const colWidth = width / sections.length;

  // Заголовок "Сравнение аналогов"
  slide.addText('Сравнение аналогов', {
    x: 0.8,
    y: y,
    w: 11.0,
    h: 0.5,
    fontSize: 16,
    fontFace: 'Segoe UI',
    color: '4A5568',
    align: 'center',
  });
  y += 0.6;

  // Колонки с аналогами
  sections.forEach((section, index) => {
    const x = 0.8 + index * colWidth;
    
    // Карточка с фоном
    slide.addShape(pptx.ShapeType.roundRect, {
      x: x,
      y: y,
      w: colWidth - 0.2,
      h: 3.0,
      fill: { color: 'FFFFFF' },
      line: { color: getAnalyticsColor(index), width: 2 },
      rectRadius: 8,
      shadow: { type: 'outer', color: '000000', opacity: 0.05, blur: 4, offset: 2 },
    });

    // Название
    slide.addText(section.name, {
      x: x + 0.1,
      y: y + 0.2,
      w: colWidth - 0.4,
      h: 0.6,
      fontSize: 18,
      fontFace: 'Segoe UI',
      color: getAnalyticsColor(index),
      bold: true,
      align: 'center',
    });

    // Особенности
    const features = section.features || [];
    let fy = y + 0.9;
    features.forEach((feature) => {
      slide.addText(`• ${feature}`, {
        x: x + 0.3,
        y: fy,
        w: colWidth - 0.6,
        h: 0.5,
        fontSize: 13,
        fontFace: 'Segoe UI',
        color: '2D3748',
        align: 'left',
        valign: 'middle',
      });
      fy += 0.6;
    });
  });

  y += 3.7;

  // Вывод
  if (content.conclusion) {
    slide.addShape(pptx.ShapeType.roundRect, {
      x: 0.8,
      y: y,
      w: 11.0,
      h: 0.8,
      fill: { color: 'EBF8FF' },
      line: { color: '0078D4', width: 1 },
      rectRadius: 6,
    });
    
    slide.addText(`💡 ${content.conclusion}`, {
      x: 1.2,
      y: y + 0.1,
      w: 10.2,
      h: 0.6,
      fontSize: 14,
      fontFace: 'Segoe UI',
      color: '2A69AC',
      align: 'left',
      valign: 'middle',
    });
    y += 1.0;
  }

  return y;
}

// 2. БЮДЖЕТ
function generateBudgetContent(slide, content, startY) {
  let y = startY;
  const budget = content.budget || [];

  // Таблица бюджетов
  const headers = ['Статья расходов', 'Сумма', '%'];
  const rows = budget.map(item => [
    item.category,
    item.amount,
    `${item.percentage}%`
  ]);

  slide.addTable(rows, {
    x: 0.8,
    y: y,
    w: 11.0,
    h: 3.0,
    fontSize: 14,
    fontFace: 'Segoe UI',
    border: { type: 'solid', color: 'E2E8F0', pt: 1 },
    fill: { color: 'FFFFFF' },
    color: '2D3748',
    rowH: 0.7,
    colW: [6.0, 3.0, 2.0],
    align: ['left', 'right', 'center'],
    valign: 'middle',
  });

  // Стилизация заголовка таблицы
  // (PptxGenJS не поддерживает прямой стиль заголовка, делаем через отдельную строку)
  // Можно добавить цветные строки позже

  y += 3.2;

  // Примечание
  if (content.note) {
    slide.addText(`📌 ${content.note}`, {
      x: 0.8,
      y: y,
      w: 11.0,
      h: 0.5,
      fontSize: 11,
      fontFace: 'Segoe UI',
      color: '718096',
      align: 'left',
    });
    y += 0.6;
  }

  return y;
}

// 3. UNIT-ЭКОНОМИКА
function generateMetricsContent(slide, content, startY) {
  let y = startY;
  const metrics = content.metrics || [];

  metrics.forEach((metric, index) => {
    const x = 0.8 + index * 3.8;
    
    // Карточка
    slide.addShape(pptx.ShapeType.roundRect, {
      x: x,
      y: y,
      w: 3.4,
      h: 2.0,
      fill: { color: 'FFFFFF' },
      line: { color: getMetricColor(index), width: 2 },
      rectRadius: 10,
      shadow: { type: 'outer', color: '000000', opacity: 0.05, blur: 4, offset: 2 },
    });

    // Значение (крупно)
    slide.addText(metric.value, {
      x: x + 0.1,
      y: y + 0.2,
      w: 3.2,
      h: 0.8,
      fontSize: 28,
      fontFace: 'Segoe UI',
      color: getMetricColor(index),
      bold: true,
      align: 'center',
    });

    // Название
    slide.addText(metric.name, {
      x: x + 0.1,
      y: y + 0.9,
      w: 3.2,
      h: 0.4,
      fontSize: 12,
      fontFace: 'Segoe UI',
      color: '4A5568',
      align: 'center',
    });

    // Примечание
    if (metric.note) {
      slide.addText(metric.note, {
        x: x + 0.1,
        y: y + 1.3,
        w: 3.2,
        h: 0.5,
        fontSize: 9,
        fontFace: 'Segoe UI',
        color: 'A0AEC0',
        align: 'center',
        valign: 'middle',
      });
    }
  });

  y += 2.3;
  return y;
}

// 4. ROADMAP
function generateRoadmapContent(slide, content, startY) {
  let y = startY;
  const phases = content.phases || [];

  phases.forEach((phase, index) => {
    const x = 0.8 + index * 3.8;
    
    // Линия времени (горизонтальная)
    if (index === 0) {
      slide.addShape(pptx.ShapeType.rect, {
        x: 0.8,
        y: y + 0.3,
        w: 11.0,
        h: 0.06,
        fill: { color: 'CBD5E0' },
      });
    }

    // Круг-маркер
    slide.addShape(pptx.ShapeType.ellipse, {
      x: x + 1.4,
      y: y + 0.1,
      w: 0.4,
      h: 0.4,
      fill: { color: getRoadmapColor(index) },
      line: { color: 'FFFFFF', width: 2 },
    });

    // Карточка
    slide.addShape(pptx.ShapeType.roundRect, {
      x: x,
      y: y + 0.5,
      w: 3.4,
      h: 2.2,
      fill: { color: 'FFFFFF' },
      line: { color: getRoadmapColor(index), width: 2 },
      rectRadius: 8,
      shadow: { type: 'outer', color: '000000', opacity: 0.05, blur: 4, offset: 2 },
    });

    // Квартал
    slide.addText(phase.quarter, {
      x: x + 0.1,
      y: y + 0.6,
      w: 3.2,
      h: 0.4,
      fontSize: 14,
      fontFace: 'Segoe UI',
      color: getRoadmapColor(index),
      bold: true,
      align: 'center',
    });

    // Название
    slide.addText(phase.name, {
      x: x + 0.1,
      y: y + 1.0,
      w: 3.2,
      h: 0.4,
      fontSize: 12,
      fontFace: 'Segoe UI',
      color: '2D3748',
      bold: true,
      align: 'center',
    });

    // Описание
    slide.addText(phase.description, {
      x: x + 0.2,
      y: y + 1.5,
      w: 3.0,
      h: 1.0,
      fontSize: 10,
      fontFace: 'Segoe UI',
      color: '718096',
      align: 'center',
      valign: 'middle',
    });
  });

  y += 3.0;
  return y;
}

// 5. КОМАНДА
function generateTeamContent(slide, content, startY) {
  let y = startY;
  const members = content.members || [];

  members.forEach((member, index) => {
    const x = 0.8 + (index % 2) * 5.6;
    const row = Math.floor(index / 2);
    const yPos = y + row * 1.2;

    // Карточка участника
    slide.addShape(pptx.ShapeType.roundRect, {
      x: x,
      y: yPos,
      w: 5.2,
      h: 1.0,
      fill: { color: 'FFFFFF' },
      line: { color: 'E2E8F0', width: 1 },
      rectRadius: 6,
    });

    // Имя
    slide.addText(member.name, {
      x: x + 0.2,
      y: yPos + 0.1,
      w: 2.0,
      h: 0.4,
      fontSize: 14,
      fontFace: 'Segoe UI',
      color: '2D3748',
      bold: true,
      align: 'left',
    });

    // Роль
    slide.addText(member.role, {
      x: x + 0.2,
      y: yPos + 0.4,
      w: 2.0,
      h: 0.3,
      fontSize: 10,
      fontFace: 'Segoe UI',
      color: '0078D4',
      align: 'left',
    });

    // Обязанности
    slide.addText(member.responsibilities, {
      x: x + 2.3,
      y: yPos + 0.1,
      w: 2.7,
      h: 0.8,
      fontSize: 10,
      fontFace: 'Segoe UI',
      color: '718096',
      align: 'left',
      valign: 'middle',
    });
  });

  const rows = Math.ceil(members.length / 2);
  y += rows * 1.3;
  return y;
}

// 6. РАСПРЕДЕЛЕНИЕ БЮДЖЕТА
function generateDistributionContent(slide, content, startY) {
  let y = startY;
  const distribution = content.distribution || [];

  // Строим круговую диаграмму в виде полос
  let currentX = 0.8;
  distribution.forEach((item) => {
    const width = (item.percentage / 100) * 11.0;
    
    slide.addShape(pptx.ShapeType.roundRect, {
      x: currentX,
      y: y,
      w: width,
      h: 0.8,
      fill: { color: item.color || '#0078D4' },
      rectRadius: 4,
    });

    // Подпись
    slide.addText(`${item.label} (${item.percentage}%)`, {
      x: currentX + 0.1,
      y: y + 0.1,
      w: width - 0.2,
      h: 0.6,
      fontSize: 12,
      fontFace: 'Segoe UI',
      color: 'FFFFFF',
      bold: true,
      align: 'center',
      valign: 'middle',
    });

    currentX += width;
  });

  y += 1.2;

  // Примечание
  if (content.note) {
    slide.addText(`📌 ${content.note}`, {
      x: 0.8,
      y: y,
      w: 11.0,
      h: 0.5,
      fontSize: 11,
      fontFace: 'Segoe UI',
      color: '718096',
      align: 'left',
    });
    y += 0.6;
  }

  return y;
}

// 7. БУЛЛЕТЫ (базовый режим)
function generateBulletsContent(slide, bullets, startY) {
  let y = startY;
  
  bullets.forEach((bullet) => {
    slide.addText(`• ${bullet}`, {
      x: 1.0,
      y: y,
      w: 10.5,
      h: 0.5,
      fontSize: 16,
      fontFace: 'Segoe UI',
      color: '2D3748',
      align: 'left',
    });
    y += 0.6;
  });

  return y;
}

// ============================================
// ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
// ============================================

function getCategoryColor(category) {
  const colors = {
    'analytics': '#4A7C59',
    'finance': '#0078D4',
    'roadmap': '#D4A017',
    'team': '#7C4A6E',
    'summary': '#1A3A6B',
  };
  return colors[category] || '#0078D4';
}

function getAnalyticsColor(index) {
  const colors = ['#0078D4', '#4A7C59', '#D4A017'];
  return colors[index % colors.length];
}

function getMetricColor(index) {
  const colors = ['#0078D4', '#4A7C59', '#D4A017'];
  return colors[index % colors.length];
}

function getRoadmapColor(index) {
  const colors = ['#0078D4', '#4A7C59', '#D4A017'];
  return colors[index % colors.length];
}

function arrayBufferToBase64(buffer) {
  const bytes = new Uint8Array(buffer);
  let binary = '';
  for (let i = 0; i < bytes.length; i++) {
    binary += String.fromCharCode(bytes[i]);
  }
  return window.btoa(binary);
}