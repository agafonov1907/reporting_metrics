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

  function saveMetrics() {
    localStorage.setItem('metrics', JSON.stringify(metrics));
    renderMetrics();
  }

  function formatPeriod(periodStr) {
    const [year, month] = periodStr.split('-');
    const monthIndex = parseInt(month, 10) - 1;
    return `${MONTHS_RU[monthIndex]} ${year}`;
  }

  // === ФОРМАТИРОВАНИЕ ДАТЫ В ВИДЕ: «от «06» февраля 2026 г.» ===
  function formatOfficialDate(date) {
    const day = String(date.getDate()).padStart(2, '0');
    const monthIndex = date.getMonth();
    const year = date.getFullYear();
    const monthName = MONTHS_RU[monthIndex].toLowerCase();
    return `от «${day}» ${monthName} ${year} г.`;
  }

  // === ГЕНЕРАЦИЯ .DOCX ОТЧЁТА ===
  function loadTemplate(url) {
    return new Promise((resolve, reject) => {
      PizZipUtils.getBinaryContent(url, (error, content) => {
        if (error) reject(error);
        else resolve(content);
      });
    });
  }

  async function generateReport(metric) {
  try {
    console.log('Загрузка шаблона report_template.docx...');
    const templateContent = await loadTemplate('report_template.docx');
    
    const data = {
      metric_value: metric.value,
      current_date: formatOfficialDate(new Date())
    };

    const zip = new PizZip(templateContent);
    const doc = new Docxtemplater(zip, {
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

    console.log('Скачивание файла...');
    saveAs(blob, `Отчёт_ПО_${metric.period}.docx`);

  } catch (error) {
    let msg = 'Неизвестная ошибка';
    if (error.properties && error.properties.errors instanceof Array) {
      msg = error.properties.errors.map(err => err.reason).join('\n');
    } else {
      msg = error.message || error.toString();
    }
    alert('❌ Ошибка генерации отчёта:\n\n' + msg);
    console.error('Ошибка генерации:', error);
  }
}

      const zip = new PizZip(templateContent);
      const doc = new Docxtemplater(zip, {
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

      saveAs(blob, `Отчёт_ПО_${metric.period}.docx`);

    } catch (error) {
      alert('Ошибка генерации отчёта:\n' + (error.message || error));
      console.error(error);
    }
  }

  function sanitizeFilename(name) {
    return name.replace(/[<>:"/\\|?*]/g, '_').substring(0, 50);
  }

  // === РЕНДЕРИНГ ===
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

      setTimeout(() => {
        card.classList.add('visible');
      }, 100 * index);
    });

    document.querySelectorAll('.delete-btn').forEach(btn => {
      btn.addEventListener('click', () => {
        const idx = parseInt(btn.dataset.index, 10);
        metrics.splice(idx, 1);
        saveMetrics();
      });
    });

    document.querySelectorAll('.report-btn').forEach(btn => {
      btn.addEventListener('click', () => {
        const idx = parseInt(btn.dataset.index, 10);
        const metric = metrics[idx];
        generateReport(metric);
      });
    });
  }

  // === ЭКСПОРТ / ИМПОРТ / ОЧИСТКА ===
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

  clearBtn.addEventListener('click', () => {
    if (confirm('Удалить все показатели?')) {
      metrics = [];
      saveMetrics();
    }
  });

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

  renderMetrics();
});