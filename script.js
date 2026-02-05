document.addEventListener('DOMContentLoaded', () => {
  const form = document.getElementById('metricForm');
  const metricsList = document.getElementById('metricsList');
  const exportBtn = document.getElementById('exportBtn');
  const importFile = document.getElementById('importFile');
  const clearBtn = document.getElementById('clearBtn');

  // Месяцы на русском
  const MONTHS_RU = [
    'Январь', 'Февраль', 'Март', 'Апрель', 'Май', 'Июнь',
    'Июль', 'Август', 'Сентябрь', 'Октябрь', 'Ноябрь', 'Декабрь'
  ];

  // Загрузка данных из localStorage
  let metrics = JSON.parse(localStorage.getItem('metrics')) || [];

  // Установка текущего месяца по умолчанию
  const now = new Date();
  const currentMonth = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}`;
  document.getElementById('metricPeriod').value = currentMonth;

  // Сохранение и рендеринг
  function saveMetrics() {
    localStorage.setItem('metrics', JSON.stringify(metrics));
    renderMetrics();
  }

  // Форматирование периода: "2026-02" → "Февраль 2026"
  function formatPeriod(periodStr) {
    const [year, month] = periodStr.split('-');
    const monthIndex = parseInt(month, 10) - 1;
    return `${MONTHS_RU[monthIndex]} ${year}`;
  }

  // Рендеринг списка показателей
  function renderMetrics() {
    metricsList.innerHTML = '';
    if (metrics.length === 0) {
      metricsList.innerHTML = '<p class="empty">Нет данных. Добавьте первый показатель!</p>';
      return;
    }

    // Сортировка: новые периоды выше
    const sorted = [...metrics].sort((a, b) => {
      return b.period.localeCompare(a.period) || a.name.localeCompare(b.name);
    });

    sorted.forEach((metric, index) => {
      const card = document.createElement('div');
      card.className = 'metric-card';
      const displayPeriod = formatPeriod(metric.period);
      card.innerHTML = `
        <div>
          <div class="metric-name">${escapeHtml(metric.name)}</div>
          <div class="metric-period">${escapeHtml(displayPeriod)}</div>
        </div>
        <div style="display:flex; align-items: center; gap: 0.75rem;">
          <span class="metric-value">${escapeHtml(metric.value)}</span>
          <button class="delete-btn" data-index="${index}">×</button>
        </div>
        <button class="report-btn" data-index="${index}">📄</button>
      `;
      metricsList.appendChild(card);

      // Анимация появления с задержкой
      setTimeout(() => {
        card.classList.add('visible');
      }, 100 * index);
    });

    // Обработчики удаления
    document.querySelectorAll('.delete-btn').forEach(btn => {
      btn.addEventListener('click', () => {
        const idx = parseInt(btn.dataset.index, 10);
        metrics.splice(idx, 1);
        saveMetrics();
      });
    });

    // Обработчики отчёта
    document.querySelectorAll('.report-btn').forEach(btn => {
      btn.addEventListener('click', () => {
        const idx = parseInt(btn.dataset.index, 10);
        const metric = metrics[idx];
        alert(`Формирование отчёта:\n\nПоказатель: ${metric.name}\nПериод: ${formatPeriod(metric.period)}\nЗначение: ${metric.value}`);
        // 🔜 Здесь вы позже добавите свою логику (PDF, API и т.д.)
      });
    });
  }

  // Экспорт в JSON
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

  // Импорт из JSON
  importFile.addEventListener('change', (e) => {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (event) => {
      try {
        const loaded = JSON.parse(event.target.result);
        if (Array.isArray(loaded)) {
          // Валидация: name и value — строки, period — формат YYYY-MM
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
      importFile.value = ''; // сбросить выбор
    };
    reader.readAsText(file);
  });

  // Очистка всех данных
  clearBtn.addEventListener('click', () => {
    if (confirm('Вы уверены, что хотите удалить все показатели?')) {
      metrics = [];
      saveMetrics();
    }
  });

  // Добавление нового показателя
  form.addEventListener('submit', (e) => {
    e.preventDefault();
    const name = document.getElementById('metricName').value.trim();
    const value = document.getElementById('metricValue').value.trim(); // ← теперь строка!
    const period = document.getElementById('metricPeriod').value;

    if (name && value !== '' && period) {
      metrics.push({ name, value, period });
      saveMetrics();
      form.reset();
      // Вернуть текущий месяц после сброса формы
      document.getElementById('metricPeriod').value = currentMonth;
    }
  });

  // Защита от XSS
  function escapeHtml(text) {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
  }

  // Первый рендер
  renderMetrics();
});