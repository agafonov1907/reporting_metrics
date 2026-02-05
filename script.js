document.addEventListener('DOMContentLoaded', () => {
  const form = document.getElementById('metricForm');
  const metricsList = document.getElementById('metricsList');
  const exportBtn = document.getElementById('exportBtn');
  const importFile = document.getElementById('importFile');
  const clearBtn = document.getElementById('clearBtn');

  const MONTHS_RU = [
    'Январь', 'Февраль', 'Март', 'Апрель', 'Май', 'Июнь',
    'Июль', 'Август', 'Сентябрь', 'Октябрь', 'Ноябрь', 'Декабрь'
  ];

  let metrics = JSON.parse(localStorage.getItem('metrics')) || [];

  const now = new Date();
  const currentMonth = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}`;
  document.getElementById('metricPeriod').value = currentMonth;

  // Проверка загрузки библиотек
  function checkLibrariesLoaded() {
    if (typeof PizZip === 'undefined') {
      console.error('PizZip не загружен!');
      return false;
    }
    if (typeof docxtemplater === 'undefined') {
      console.error('docxtemplater не загружен!');
      return false;
    }
    return true;
  }

  function saveMetrics() {
    localStorage.setItem('metrics', JSON.stringify(metrics));
    renderMetrics();
  }

  function formatPeriod(periodStr) {
    const [year, month] = periodStr.split('-');
    const monthIndex = parseInt(month, 10) - 1;
    return `${MONTHS_RU[monthIndex]} ${year}`;
  }

  // Загрузка шаблона через fetch
  function loadTemplate(url) {
    return fetch(url)
      .then(response => {
        if (!response.ok) {
          throw new Error(`Не удалось загрузить шаблон: ${response.status} ${response.statusText}`);
        }
        return response.arrayBuffer();
      })
      .catch(error => {
        console.error('Ошибка загрузки шаблона:', error);
        throw new Error(`Не удалось загрузить шаблон DOCX: ${error.message}`);
      });
  }

  // Генерация отчёта
  async function generateReport(metric) {
    try {
      // Проверяем загружены ли библиотеки
      if (!checkLibrariesLoaded()) {
        throw new Error('Библиотеки для генерации DOCX не загружены. Проверьте подключение скриптов.');
      }

      const templateArrayBuffer = await loadTemplate('report_template.docx');

      const data = {
        metric_value: metric.value,
        current_date: new Date().toLocaleDateString('ru-RU', {
          day: '2-digit',
          month: '2-digit',
          year: 'numeric'
        })
      };

      // Используем правильные имена классов
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

      if (blob.size === 0) {
        throw new Error('Сгенерированный файл пуст');
      }

      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `Отчёт_ПО_${metric.period}.docx`;
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
      
      // Более информативное сообщение об ошибке
      console.error('Полная ошибка генерации отчёта:', error);
      
      if (msg.includes('PizZip') || msg.includes('docxtemplater')) {
        alert('❌ Ошибка: Библиотеки для работы с DOCX не загружены.\n\n' +
              'Пожалуйста:\n' +
              '1. Проверьте подключение к интернету\n' +
              '2. Обновите страницу\n' +
              '3. Если проблема сохраняется, свяжитесь с поддержкой');
      } else {
        alert('❌ Ошибка генерации отчёта:\n' + msg);
      }
    }
  }

  // Рендеринг карточек
  function renderMetrics() {
    metricsList.innerHTML = '';
    if (metrics.length === 0) {
      metricsList.innerHTML = '<p class="empty">Нет данных. Добавьте первый показатель!</p>';
      return;
    }

    const sorted = [...metrics].sort((a, b) => b.period.localeCompare(a.period) || a.name.localeCompare(b.name));

    sorted.forEach((metric, index) => {
      const card = document.createElement('div');
      card.className = 'metric-card';
      const displayPeriod = formatPeriod(metric.period);
      
      // Проверяем доступность генерации отчетов
      const canGenerateReport = checkLibrariesLoaded();
      
      card.innerHTML = `
        <div>
          <div class="metric-name">${escapeHtml(metric.name)}</div>
          <div class="metric-period">${escapeHtml(displayPeriod)}</div>
        </div>
        <div style="display:flex; align-items: center; gap: 0.75rem;">
          <span class="metric-value">${escapeHtml(metric.value)}</span>
          <button class="delete-btn" data-index="${index}">×</button>
        </div>
        <button class="report-btn" data-index="${index}" ${!canGenerateReport ? 'title="Генерация отчетов недоступна" disabled' : ''}>
          📄
        </button>
      `;
      metricsList.appendChild(card);

      setTimeout(() => {
        card.classList.add('visible');
      }, 100 * index);
    });

    // Обработчики для кнопок удаления
    document.querySelectorAll('.delete-btn').forEach(btn => {
      btn.addEventListener('click', () => {
        const idx = parseInt(btn.dataset.index, 10);
        metrics.splice(idx, 1);
        saveMetrics();
      });
    });

    // Обработчики для кнопок генерации отчетов
    document.querySelectorAll('.report-btn:not(:disabled)').forEach(btn => {
      btn.addEventListener('click', () => {
        const idx = parseInt(btn.dataset.index, 10);
        const metric = metrics[idx];
        generateReport(metric);
      });
    });
  }

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
            m.value.trim() !== '' &&
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

    if (name && value !== '' && period) {
      metrics.push({ name, value, period });
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

  // Проверяем загрузку библиотек при старте
  console.log('PizZip loaded:', typeof PizZip !== 'undefined');
  console.log('docxtemplater loaded:', typeof docxtemplater !== 'undefined');
  
  // Небольшая задержка для гарантированной загрузки библиотек
  setTimeout(() => {
    renderMetrics();
    
    // Если библиотеки не загружены, показываем предупреждение
    if (!checkLibrariesLoaded()) {
      console.warn('Библиотеки для работы с DOCX не загружены. Функция генерации отчетов недоступна.');
    }
  }, 100);
});