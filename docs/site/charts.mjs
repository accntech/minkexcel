function scale(peak) {
  const magnitude = 10 ** Math.floor(Math.log10((peak || 1) / 4));
  const fraction = peak / 4 / magnitude;
  const step = (fraction <= 1 ? 1 : fraction <= 2 ? 2 : fraction <= 5 ? 5 : 10) * magnitude;
  return { step, maximum: Math.ceil((peak || 1) / step) * step };
}
function bars(rows, own, other, unit, description) {
  const { maximum, step } = scale(Math.max(...rows.flatMap(row => [row[own], row[other]])));
  const size = unit === 'KiB';
  const formatted = value => (size ? value / 1024 : value).toFixed(1);
  const ticks = Array.from({ length: Math.round(maximum / step) + 1 }, (_, index) =>
    `<span>${Number((size ? step * index / 1024 : step * index).toPrecision(3))}</span>`).join('');
  return `<div class="bar-chart" role="img" aria-label="${description}">
    ${rows.map(row => `<div class="bar-group"><p class="workload-label">${row.workload[0].toUpperCase() + row.workload.slice(1)}</p>
      ${[[own, 'mink', 'MinkExcel'], [other, 'excel', 'ExcelJS']].map(([key, kind, label]) => {
        const percent = row[key] / maximum * 100;
        return `<div class="bar-row"><span class="bar-label">${label}</span><div class="bar-track"><div class="bar-${kind}" style="width:${percent}%" data-value="${row[key]}" data-percent="${percent}"></div></div><span class="chart-value">${formatted(row[key])}<small>${unit}</small></span></div>`;
      }).join('')}</div>`).join('')}
    <div class="chart-scale"><span></span><div>${ticks}</div><span></span></div>
    <p class="chart-axis">${size ? 'Compressed file size (KiB)' : 'Time (ms)'} · ${size ? 'smaller' : 'lower'} is better</p>
  </div>`;
}
export function timingChart(rows, operation = 'Export') {
  return bars(rows, `minkexcel${operation}Ms`, `exceljs${operation}Ms`, 'ms',
    `${operation} time in milliseconds for numeric, text and mixed workloads. Lower is better. Exact values follow in the results table.`);
}
export function sizeChart(rows) {
  return bars(rows, 'minkexcelBytes', 'exceljsBytes', 'KiB',
    'Compressed XLSX file sizes for numeric, text and mixed workloads. Exact byte sizes are available in the downloadable JSON.');
}
export function resultsTable(rows) {
  return `<thead><tr><th scope="col">Workload</th><th scope="col">Rows</th><th scope="col">Mink export</th><th scope="col">ExcelJS export</th><th scope="col">Mink import</th><th scope="col">ExcelJS import</th></tr></thead><tbody>${rows.map(row => `<tr><th scope="row">${row.workload}</th><td>${numericText(row.rows.toLocaleString('en-US'))}</td>${['minkexcelExportMs', 'exceljsExportMs', 'minkexcelImportMs', 'exceljsImportMs'].map(key => `<td>${numericText(row[key].toFixed(2))} ms</td>`).join('')}</tr>`).join('')}</tbody>`;
}
import { numericText } from './numbers.mjs';
