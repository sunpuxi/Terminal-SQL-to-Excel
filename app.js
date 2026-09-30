(function () {
  "use strict";

  const PREVIEW_LIMIT = 100;
  const LARGE_DATASET_SIZE = 10000;

  function parseMysqlConsole(text) {
    if (!text || !text.trim()) {
      return { ok: false, error: "请先粘贴 MySQL 查询结果。" };
    }

    const sourceLines = text.replace(/\r\n?/g, "\n").split("\n");
    const rows = [];
    const lineNumbers = [];

    sourceLines.forEach(function (line, index) {
      const trimmed = line.trim();
      if (trimmed.startsWith("|") && trimmed.endsWith("|")) {
        rows.push(trimmed.slice(1, -1).split("|").map(function (cell) { return cell.trim(); }));
        lineNumbers.push(index + 1);
      }
    });

    if (!rows.length) {
      return {
        ok: false,
        error: "没有识别到有效的 MySQL 查询结果。\n请确认粘贴内容包含类似：| id | name | status | 的表格数据。"
      };
    }

    const columnCount = rows[0].length;
    if (!columnCount || rows[0].every(function (cell) { return cell === ""; })) {
      return { ok: false, error: "没有识别到有效的表头。" };
    }

    for (let index = 1; index < rows.length; index += 1) {
      if (rows[index].length !== columnCount) {
        return {
          ok: false,
          error: "第 " + lineNumbers[index] + " 行数据列数与表头不一致（应为 " + columnCount + " 列，实际为 " + rows[index].length + " 列）。"
        };
      }
    }

    return {
      ok: true,
      headers: rows[0],
      rows: rows.slice(1),
      data: rows,
      rowCount: Math.max(rows.length - 1, 0),
      columnCount: columnCount
    };
  }

  function displayLength(value) {
    let length = 0;
    Array.from(String(value)).forEach(function (char) {
      length += /[^\x00-\xff]/.test(char) ? 2 : 1;
    });
    return length;
  }

  function createFilename(now) {
    const pad = function (value) { return String(value).padStart(2, "0"); };
    return "query_result_" + now.getFullYear() + pad(now.getMonth() + 1) + pad(now.getDate()) + "_" + pad(now.getHours()) + pad(now.getMinutes()) + ".xlsx";
  }

  function selectColumns(data, selectedIndexes) {
    if (!Array.isArray(data) || !data.length) return [];
    if (!Array.isArray(selectedIndexes) || !selectedIndexes.length) return [];
    return data.map(function (row) {
      return selectedIndexes.map(function (columnIndex) { return row[columnIndex]; });
    });
  }

  function applyWorksheetFormatting(worksheet, data) {
    worksheet["!cols"] = data[0].map(function (_, columnIndex) {
      let maxLength = 0;
      data.forEach(function (row) {
        maxLength = Math.max(maxLength, displayLength(row[columnIndex] || ""));
      });
      return { wch: Math.min(Math.max(maxLength + 2, 8), 50) };
    });

    data[0].forEach(function (_, columnIndex) {
      const cellAddress = XLSX.utils.encode_cell({ r: 0, c: columnIndex });
      if (worksheet[cellAddress]) {
        worksheet[cellAddress].s = {
          font: { bold: true, color: { rgb: "FFFFFF" } },
          fill: { patternType: "solid", fgColor: { rgb: "176B45" } },
          alignment: { vertical: "center" }
        };
      }
    });
    worksheet["!rows"] = [{ hpt: 22 }];
    worksheet["!autofilter"] = { ref: worksheet["!ref"] };
    worksheet["!freeze"] = { xSplit: 0, ySplit: 1, topLeftCell: "A2", activePane: "bottomLeft", state: "frozen" };
  }

  if (typeof module !== "undefined" && module.exports) {
    module.exports = { parseMysqlConsole, displayLength, createFilename, selectColumns };
  }

  if (typeof document === "undefined") return;

  const input = document.getElementById("sql-input");
  const parseButton = document.getElementById("parse-button");
  const clearButton = document.getElementById("clear-button");
  const exportButton = document.getElementById("export-button");
  const inputCount = document.getElementById("input-count");
  const message = document.getElementById("message");
  const previewSection = document.getElementById("preview-section");
  const previewTable = document.getElementById("preview-table");
  const previewNote = document.getElementById("preview-note");
  const rowCount = document.getElementById("row-count");
  const columnCount = document.getElementById("column-count");
  const columnOptions = document.getElementById("column-options");
  const selectionSummary = document.getElementById("selection-summary");
  const selectAllColumns = document.getElementById("select-all-columns");
  const clearAllColumns = document.getElementById("clear-all-columns");
  const exportCopy = document.getElementById("export-copy");
  let parsedResult = null;
  let pasteTimer = null;

  function setMessage(text, type) {
    message.textContent = text;
    message.className = "message " + (type || "error");
    message.hidden = !text;
  }

  function resetPreview() {
    parsedResult = null;
    previewSection.hidden = true;
    previewTable.replaceChildren();
    columnOptions.replaceChildren();
  }

  function getSelectedColumnIndexes() {
    return Array.from(columnOptions.querySelectorAll("input:checked")).map(function (checkbox) {
      return Number(checkbox.value);
    });
  }

  function updateColumnSelectionState() {
    if (!parsedResult) return;
    const selectedCount = getSelectedColumnIndexes().length;
    const totalCount = parsedResult.columnCount;
    selectionSummary.classList.toggle("selection-error", selectedCount === 0);
    selectionSummary.textContent = selectedCount === 0
      ? "请至少选择 1 个要导出的字段"
      : "已选择 " + selectedCount + " / " + totalCount + " 个字段";
    exportCopy.textContent = selectedCount === totalCount
      ? "Excel 将包含全部已解析数据。"
      : selectedCount > 0
        ? "Excel 将导出已选的 " + selectedCount + " 个字段和全部 " + parsedResult.rowCount + " 行数据。"
        : "当前未选择任何导出字段。";
  }

  function renderColumnSelector(headers) {
    const fragment = document.createDocumentFragment();
    headers.forEach(function (header, index) {
      const label = document.createElement("label");
      label.className = "column-option";
      const checkbox = document.createElement("input");
      checkbox.type = "checkbox";
      checkbox.value = String(index);
      checkbox.checked = true;
      const text = document.createElement("span");
      text.textContent = header || "未命名字段 " + (index + 1);
      text.title = header;
      label.append(checkbox, text);
      fragment.appendChild(label);
    });
    columnOptions.replaceChildren(fragment);
    updateColumnSelectionState();
  }

  function updateInputState() {
    const value = input.value;
    clearButton.disabled = !value;
    if (!value) {
      inputCount.textContent = "等待粘贴数据";
      return;
    }
    const candidateRows = value.split(/\r?\n/).filter(function (line) {
      const trimmed = line.trim();
      return trimmed.startsWith("|") && trimmed.endsWith("|");
    }).length;
    inputCount.textContent = candidateRows > 0 ? "检测到 " + Math.max(candidateRows - 1, 0) + " 行候选数据" : "尚未检测到表格数据";
  }

  function renderTable(result) {
    const head = document.createElement("thead");
    const headRow = document.createElement("tr");
    result.headers.forEach(function (header) {
      const th = document.createElement("th");
      th.textContent = header;
      th.title = header;
      headRow.appendChild(th);
    });
    head.appendChild(headRow);

    const body = document.createElement("tbody");
    result.rows.slice(0, PREVIEW_LIMIT).forEach(function (row) {
      const tr = document.createElement("tr");
      row.forEach(function (value) {
        const td = document.createElement("td");
        td.textContent = value === "" ? "空" : value;
        td.title = value;
        if (value === "") td.className = "empty-cell";
        tr.appendChild(td);
      });
      body.appendChild(tr);
    });

    previewTable.replaceChildren(head, body);
    rowCount.textContent = String(result.rowCount);
    columnCount.textContent = String(result.columnCount);
    previewNote.textContent = result.rowCount > PREVIEW_LIMIT
      ? "当前共有 " + result.rowCount + " 行数据，仅预览前 " + PREVIEW_LIMIT + " 行；导出时仍会包含全部数据。"
      : "已解析全部数据，请确认表头和内容后导出。";
    renderColumnSelector(result.headers);
    previewSection.hidden = false;
  }

  function parseAndRender() {
    setMessage("");
    const result = parseMysqlConsole(input.value);
    if (!result.ok) {
      resetPreview();
      setMessage(result.error, "error");
      return false;
    }
    parsedResult = result;
    renderTable(result);
    if (result.rowCount > LARGE_DATASET_SIZE) {
      setMessage("当前数据量较大，解析和导出可能需要一些时间。", "warning");
    }
    return true;
  }

  input.addEventListener("input", function () {
    updateInputState();
    if (parsedResult) resetPreview();
    setMessage("");
  });

  input.addEventListener("paste", function () {
    window.clearTimeout(pasteTimer);
    pasteTimer = window.setTimeout(function () {
      updateInputState();
      parseAndRender();
    }, 0);
  });

  parseButton.addEventListener("click", parseAndRender);

  clearButton.addEventListener("click", function () {
    input.value = "";
    setMessage("");
    resetPreview();
    updateInputState();
    input.focus();
  });

  columnOptions.addEventListener("change", updateColumnSelectionState);

  selectAllColumns.addEventListener("click", function () {
    columnOptions.querySelectorAll("input").forEach(function (checkbox) { checkbox.checked = true; });
    updateColumnSelectionState();
  });

  clearAllColumns.addEventListener("click", function () {
    columnOptions.querySelectorAll("input").forEach(function (checkbox) { checkbox.checked = false; });
    updateColumnSelectionState();
  });

  exportButton.addEventListener("click", function () {
    if (!parsedResult && !parseAndRender()) return;
    if (typeof XLSX === "undefined") {
      setMessage("Excel 导出组件未加载，请确认 libs/xlsx.full.min.js 文件存在。", "error");
      return;
    }

    try {
      const selectedIndexes = getSelectedColumnIndexes();
      if (!selectedIndexes.length) {
        setMessage("请至少选择 1 个要导出的字段。", "error");
        return;
      }
      setMessage("");
      const exportData = selectColumns(parsedResult.data, selectedIndexes);
      const worksheet = XLSX.utils.aoa_to_sheet(exportData);
      applyWorksheetFormatting(worksheet, exportData);
      const workbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(workbook, worksheet, "Query Result");
      XLSX.writeFile(workbook, createFilename(new Date()), { cellStyles: true, compression: true });
    } catch (error) {
      console.error(error);
      setMessage("导出失败，请重试或检查浏览器是否允许下载文件。", "error");
    }
  });

  updateInputState();
})();
