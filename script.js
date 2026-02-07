document.addEventListener('DOMContentLoaded', () => {
  const form = document.getElementById('metricForm');
  const sectionsContainer = document.getElementById('sectionsContainer');
  const exportBtn = document.getElementById('exportBtn');
  const importFile = document.getElementById('importFile');
  const clearBtn = document.getElementById('clearBtn');
  const generateSummaryBtn = document.getElementById('generateSummaryBtn');

  const MONTHS_RU = [
    'Январь', 'Февраль', 'Март', 'Апрель', 'Май', 'Июнь',
    'Июль', 'Август', 'Сентябрь', 'Октябрь', 'Ноябрь', 'Декабрь'
  ];

  let metrics = JSON.parse(localStorage.getItem('metrics')) || [];

  const now = new Date();
  const currentMonth = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}`;
  document.getElementById('metricPeriod').value = currentMonth;

  // Подписи разделов
  const SECTION_LABELS = {
    po: 'РП Ц5 "Отечественные решения',
    kpi: 'KPI цифровизации',
    municipal: 'Муниципальные услуги',
    other: 'Прочее'
  };

  function saveMetrics() {
    localStorage.setItem('metrics', JSON.stringify(metrics));
    renderMetrics();
  }

  function formatPeriod(periodStr) {
    const [year, month] = periodStr.split('-');
    const monthIndex = parseInt(month, 10) - 1;
    return `${MONTHS_RU[monthIndex]} ${year}`;
  }

  // Загрузка шаблона
  function loadTemplate(url) {
    return fetch(url)
      .then(response => {
        if (!response.ok) {
          throw new Error(`Не удалось загрузить шаблон: ${response.status} ${response.statusText}`);
        }
        return response.arrayBuffer();
      });
  }

  // Генерация отдельного отчёта
  async function generateReport(metric) {
    try {
      const templateFile = metric.template || 'report_template.docx';
      const templateArrayBuffer = await loadTemplate(templateFile);

      const data = {
        metric_value: metric.value,
        current_date: new Date().toLocaleDateString('ru-RU', {
          day: '2-digit',
          month: '2-digit',
          year: 'numeric'
        })
      };

      const zip = new PizZip(templateArrayBuffer);
      const doc = new docxtemplater(zip, {
        paragraphLoop: true,
        lineBreaks: true,
        nullGetter: () => ''
      });

      doc.setData(data);
      doc.render();

      const blob = doc.getZip().generate({
        type: 'blob',
        mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
      });

      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `Отчёт_${sanitizeFilename(metric.name)}_${metric.period}.docx`;
      document.body.appendChild(a);
      a.click();
      setTimeout(() => {
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
      }, 100);

    } catch (error) {
      let msg = error.message || 'Неизвестная ошибка';
      if (error.properties?.errors) {
        msg = error.properties.errors.map(e => e.reason).join('\n');
      }
      alert('❌ Ошибка генерации отчёта:\n' + msg);
      console.error('Ошибка:', error);
    }
  }

  // Генерация сводного отчёта
  async function generateSummaryReport(selectedMetrics) {
    try {
      const templateArrayBuffer = await loadTemplate('summary_report_template.docx');

      const data = {
        current_date: new Date().toLocaleDateString('ru-RU', {
          day: '2-digit',
          month: '2-digit',
          year: 'numeric'
        })
      };

      selectedMetrics.forEach((metric, i) => {
        const key = metric.name.toLowerCase()
          .replace(/\s+/g, '_')
          .replace(/[^a-z0-9_]/g, '');
        data[key + '_value'] = metric.value;
        data[key + '_period'] = formatPeriod(metric.period);
      });

      const zip = new PizZip(templateArrayBuffer);
      const doc = new docxtemplater(zip, {
        paragraphLoop: true,
        lineBreaks: true,
        nullGetter: () => ''
      });

      doc.setData(data);
      doc.render();

      const blob = doc.getZip().generate({
        type: 'blob',
        mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
      });

      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `Сводный_отчёт_${new Date().toISOString().slice(0,10)}.docx`;
      document.body.appendChild(a);
      a.click();
      setTimeout(() => {
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
      }, 100);

    } catch (error) {
      let msg = error.message || 'Неизвестная ошибка';
      if (error.properties?.errors) {
        msg = error.properties.errors.map(e => e.reason).join('\n');
      }
      alert('❌ Ошибка генерации сводного отчёта:\n' + msg);
      console.error('Ошибка:', error);
    }
  }

  // Санитизация имени файла
  function sanitizeFilename(name) {
    return name.replace(/[<>:"/\\|?*]/g, '_').substring(0, 50);
  }

  // Рендеринг по разделам
  function renderMetrics() {
    sectionsContainer.innerHTML = '';

    if (metrics.length === 0) {
      sectionsContainer.innerHTML = '<p class="empty">Нет данных. Добавьте первый показатель!</p>';
      return;
    }

    // Группировка по разделам
    const sections = {};
    metrics.forEach((metric, index) => {
      const sec = metric.section || 'other';
      if (!sections[sec]) sections[sec] = [];
      sections[sec].push({ ...metric, originalIndex: index });
    });

    // Рендерим каждый раздел
    Object.keys(sections).forEach(sectionKey => {
      const items = sections[sectionKey];
      const sectionId = `section-${sectionKey}`;

      const sectionEl = document.createElement('div');
      sectionEl.className = 'section';

      sectionEl.innerHTML = `
        <div class="section-header" data-section="${sectionKey}">
          <div class="section-title">${SECTION_LABELS[sectionKey] || sectionKey}</div>
          <div class="section-count">${items.length}</div>
        </div>
        <div class="section-content" id="${sectionId}"></div>
      `;

      sectionsContainer.appendChild(sectionEl);

      // Добавляем карточки
      const contentEl = document.getElementById(sectionId);
      items.forEach((item, idx) => {
        const card = document.createElement('div');
        card.className = 'metric-card';
        const displayPeriod = formatPeriod(item.period);
        card.innerHTML = `
          <div>
            <div class="metric-name">${escapeHtml(item.name)}</div>
            <div class="metric-period">${escapeHtml(displayPeriod)}</div>
          </div>
          <div style="display:flex; align-items: center; gap: 0.75rem;">
            <span class="metric-value">${escapeHtml(item.value)}</span>
            <button class="delete-btn" data-index="${item.originalIndex}">×</button>
          </div>
          <button class="report-btn" data-index="${item.originalIndex}">📄</button>
        `;
        contentEl.appendChild(card);

        // Анимация появления
        setTimeout(() => {
          card.classList.add('visible');
        }, 100 * idx);
      });

      // Обработчики кнопок
      contentEl.querySelectorAll('.delete-btn').forEach(btn => {
        btn.addEventListener('click', () => {
          const idx = parseInt(btn.dataset.index, 10);
          metrics.splice(idx, 1);
          saveMetrics();
        });
      });

      contentEl.querySelectorAll('.report-btn').forEach(btn => {
        btn.addEventListener('click', () => {
          const idx = parseInt(btn.dataset.index, 10);
          generateReport(metrics[idx]);
        });
      });

      // Сворачивание/разворачивание
      const header = sectionEl.querySelector('.section-header');
      header.addEventListener('click', () => {
        const content = sectionEl.querySelector('.section-content');
        content.classList.toggle('collapsed');
      });
    });
  }

  // Сводный отчёт
  generateSummaryBtn.addEventListener('click', () => {
    const checkedBoxes = document.querySelectorAll('.metric-checkbox:checked');
    if (checkedBoxes.length === 0) {
      alert('Выберите хотя бы один показатель для сводного отчёта');
      return;
    }
    const selected = Array.from(checkedBoxes).map(box => {
      const idx = parseInt(box.dataset.index, 10);
      return metrics[idx];
    });
    generateSummaryReport(selected);
  });

  // Экспорт
  exportBtn.addEventListener('click', () => {
    const dataStr = JSON.stringify(metrics, null, 2);
    const blob = new Blob([dataStr], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'metrics.json';
    document.body.appendChild(a);
    a.click();
    setTimeout(() => {
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
    }, 0);
  });

  // Импорт
  importFile.addEventListener('change', (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (event) => {
      try {
        const loaded = JSON.parse(event.target.result);
        if (Array.isArray(loaded)) {
          const valid = loaded.every(m =>
            typeof m.name === 'string' &&
            typeof m.value === 'string' &&
            typeof m.section === 'string' &&
            m.value.trim() !== '' &&
            m.section.trim() !== '' &&
            /^\d{4}-\d{2}$/.test(m.period)
          );
          if (!valid) throw new Error('Неверный формат данных');
          metrics = loaded;
          saveMetrics();
          alert('Данные успешно загружены!');
        } else {
          throw new Error('Ожидается массив');
        }
      } catch (err) {
        alert('Ошибка при загрузке файла:\n' + err.message);
      }
      importFile.value = '';
    };
    reader.readAsText(file);
  });

  // Очистка
  clearBtn.addEventListener('click', () => {
    if (confirm('Удалить все показатели?')) {
      metrics = [];
      saveMetrics();
    }
  });

  // Добавление
  form.addEventListener('submit', (e) => {
    e.preventDefault();
    const name = document.getElementById('metricName').value.trim();
    const value = document.getElementById('metricValue').value.trim();
    const period = document.getElementById('metricPeriod').value;
    const section = document.getElementById('metricSection').value;

    if (name && value !== '' && period && section) {
      metrics.push({ name, value, period, section });
      saveMetrics();
      form.reset();
      document.getElementById('metricPeriod').value = currentMonth;
    }
  });

  function escapeHtml(text) {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
  }

  renderMetrics();
});