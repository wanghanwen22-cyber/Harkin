(function () {
  "use strict";

  var STORAGE_KEY = "recon_web_app_v3";
  var DATE_EPOCH_1900 = Date.UTC(1899, 11, 30);
  var BUILTIN_DATE_FORMAT_IDS = {
    14: true, 15: true, 16: true, 17: true, 18: true, 19: true, 20: true, 21: true, 22: true,
    27: true, 30: true, 36: true, 45: true, 46: true, 47: true, 50: true, 57: true
  };
  var DETAIL_STATE = { taskId: null, filter: "all" };
  var state = loadState();

  function loadState() {
    try {
      var raw = localStorage.getItem(STORAGE_KEY);
      if (raw) return normalizeState(JSON.parse(raw));
    } catch (error) {
      console.error("load state failed", error);
    }
    return seedState();
  }

  function saveState() {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
  }

  function normalizeState(input) {
    var seeded = seedState();
    return {
      settings: Object.assign({}, seeded.settings, input.settings || {}),
      tasks: Array.isArray(input.tasks) ? input.tasks.map(finalizeTask) : seeded.tasks
    };
  }

  function seedState() {
    return {
      settings: {
        amountTolerance: 0,
        dateToleranceDays: 0,
        autoMatchScore: 98,
        alertDelayMinutes: 15,
        ingestDirectory: "本地文件上传",
        archiveTarget: "localStorage://reconciliation",
        primaryGateway: "公司结算账户",
        backupGateway: "直接 ACH 网关",
        analystName: ""
      },
      tasks: []
    };
  }

  function finalizeTask(task) {
    task.matchedRecords = Array.isArray(task.matchedRecords) ? task.matchedRecords : [];
    task.unmatchedBankRecords = Array.isArray(task.unmatchedBankRecords) ? task.unmatchedBankRecords : [];
    task.unmatchedLedgerRecords = Array.isArray(task.unmatchedLedgerRecords) ? task.unmatchedLedgerRecords : [];
    task.entries = task.entries && task.entries.length ? task.entries : combineEntries(task);
    var total = Number(task.matchedCount || 0) + Number(task.unmatchedBankCount || 0);
    task.accuracy = total === 0 ? 0 : round((Number(task.matchedCount || 0) / total) * 100, 2);
    task.discrepancyCount = Number(task.unmatchedBankCount || 0) + Number(task.unmatchedLedgerCount || 0);
    task.statusLabel = task.status === "success" ? "成功" : task.status === "processing" ? "处理中" : task.status === "error" ? "错误" : "存在差异";
    task.summary = {
      bank_total: Number(task.matchedCount || 0) + Number(task.unmatchedBankCount || 0),
      ledger_total: Number(task.matchedCount || 0) + Number(task.unmatchedLedgerCount || 0),
      matched: Number(task.matchedCount || 0),
      bank_unmatched: Number(task.unmatchedBankCount || 0),
      ledger_unmatched: Number(task.unmatchedLedgerCount || 0),
      amount_tolerance: String(task.settingsSnapshot && task.settingsSnapshot.amountTolerance != null ? task.settingsSnapshot.amountTolerance : 0),
      date_tolerance_days: Number(task.settingsSnapshot && task.settingsSnapshot.dateToleranceDays != null ? task.settingsSnapshot.dateToleranceDays : 0)
    };
    return task;
  }

  function combineEntries(task) {
    return task.matchedRecords
      .concat(task.unmatchedBankRecords)
      .concat(task.unmatchedLedgerRecords);
  }

  function generateTaskId(seed) {
    if (seed) {
      return "RE-" + String(seed).padStart(8, "0") + "-001";
    }
    var now = new Date();
    var datePart = [
      now.getFullYear(),
      String(now.getMonth() + 1).padStart(2, "0"),
      String(now.getDate()).padStart(2, "0")
    ].join("");
    var prefix = "RE-" + datePart + "-";
    var sequence = state.tasks.filter(function (task) {
      return typeof task.id === "string" && task.id.indexOf(prefix) === 0;
    }).length + 1;
    return prefix + String(sequence).padStart(3, "0");
  }

  function round(value, digits) {
    var factor = Math.pow(10, digits || 0);
    return Math.round(Number(value || 0) * factor) / factor;
  }

  function formatMoney(value) {
    return new Intl.NumberFormat("zh-CN", {
      style: "currency",
      currency: "CNY",
      maximumFractionDigits: 2
    }).format(Number(value || 0));
  }

  function formatDateTime(value) {
    if (!value) return "-";
    return new Intl.DateTimeFormat("zh-CN", {
      year: "numeric",
      month: "2-digit",
      day: "2-digit",
      hour: "2-digit",
      minute: "2-digit"
    }).format(new Date(value));
  }

  function formatDate(value) {
    if (!value) return "-";
    return new Date(value).toISOString().slice(0, 10);
  }

  function parseDate(value) {
    if (value == null || value === "") return null;
    if (value instanceof Date && !Number.isNaN(value.getTime())) return value;
    if (typeof value === "number" && Number.isFinite(value) && value > 20000 && value < 80000) {
      return excelSerialToDate(value);
    }
    var normalized = String(value).trim();
    var patterns = [
      /^(\d{4})-(\d{2})-(\d{2})$/,
      /^(\d{4})\/(\d{2})\/(\d{2})$/,
      /^(\d{4})(\d{2})(\d{2})$/,
      /^(\d{2})\/(\d{2})\/(\d{4})$/,
      /^(\d{2})-(\d{2})-(\d{4})$/
    ];
    for (var i = 0; i < patterns.length; i += 1) {
      var match = normalized.match(patterns[i]);
      if (!match) continue;
      if (i <= 2) return new Date(match[1] + "-" + match[2] + "-" + match[3] + "T00:00:00");
      var day = Number(match[1]);
      var month = Number(match[2]);
      if (month > 12 && day <= 12) {
        var swapped = day;
        day = month;
        month = swapped;
      }
      return new Date(match[3] + "-" + String(month).padStart(2, "0") + "-" + String(day).padStart(2, "0") + "T00:00:00");
    }
    var fallback = new Date(normalized);
    return Number.isNaN(fallback.getTime()) ? null : fallback;
  }

  function excelSerialToDate(serial) {
    var millis = DATE_EPOCH_1900 + Math.round(Number(serial) * 86400000);
    return new Date(millis);
  }

  function parseAmount(value) {
    if (value == null || value === "") return null;
    var numeric = Number(String(value).replace(/,/g, "").trim());
    return Number.isFinite(numeric) ? numeric : null;
  }

  function parseCsv(text) {
    var rows = [];
    var row = [];
    var current = "";
    var inQuotes = false;
    for (var i = 0; i < text.length; i += 1) {
      var char = text[i];
      var next = text[i + 1];
      if (char === "\"") {
        if (inQuotes && next === "\"") {
          current += "\"";
          i += 1;
        } else {
          inQuotes = !inQuotes;
        }
        continue;
      }
      if (char === "," && !inQuotes) {
        row.push(current);
        current = "";
        continue;
      }
      if ((char === "\n" || char === "\r") && !inQuotes) {
        if (char === "\r" && next === "\n") i += 1;
        row.push(current);
        if (row.some(function (cell) { return cell !== ""; })) rows.push(row);
        row = [];
        current = "";
        continue;
      }
      current += char;
    }
    row.push(current);
    if (row.some(function (cell) { return cell !== ""; })) rows.push(row);
    return rowsToObjects(rows);
  }

  function rowsToObjects(rows) {
    if (!rows.length) return [];
    var headers = rows[0].map(function (header, index) {
      var value = String(header == null ? "" : header).trim();
      return value || "column_" + (index + 1);
    });
    return rows.slice(1).map(function (values) {
      var item = {};
      headers.forEach(function (header, index) {
        item[header] = values[index] == null ? "" : values[index];
      });
      return item;
    });
  }

  function normalizeRows(rows) {
    var dateKeys = ["date", "日期", "交易日期", "入账日期", "时间", "timestamp"];
    var amountKeys = ["amount", "金额", "应收金额", "支付金额", "流水金额"];
    var idKeys = ["tx_id", "id", "流水号", "单号", "交易ID", "交易id"];

    return rows.map(function (raw, index) {
      var keys = Object.keys(raw);
      var dateKey = findFirst(keys, dateKeys);
      var amountKey = findFirst(keys, amountKeys);
      var idKey = findFirst(keys, idKeys);
      var parsedDate = parseDate(dateKey ? raw[dateKey] : "");
      var parsedAmount = parseAmount(amountKey ? raw[amountKey] : "");
      return {
        index: index + 1,
        raw: raw,
        amount: parsedAmount,
        txDate: parsedDate,
        txId: idKey ? String(raw[idKey] || "").trim() : "",
        description: String(raw.description || raw.摘要 || raw.备注 || raw.memo || "")
      };
    }).filter(function (row) {
      return row.amount !== null && row.txDate;
    });
  }

  function findFirst(keys, candidates) {
    var lowerKeys = keys.map(function (key) { return String(key).toLowerCase(); });
    for (var i = 0; i < candidates.length; i += 1) {
      var target = String(candidates[i]).toLowerCase();
      var index = lowerKeys.indexOf(target);
      if (index !== -1) return keys[index];
    }
    return keys.find(function (key) {
      return candidates.some(function (candidate) {
        return String(key).toLowerCase().indexOf(String(candidate).toLowerCase()) !== -1;
      });
    }) || null;
  }

  function reconcile(bankRows, ledgerRows) {
    var amountTolerance = Number(state.settings.amountTolerance || 0);
    var dateToleranceDays = Number(state.settings.dateToleranceDays || 0);
    var matched = [];
    var bankUnmatched = {};
    var ledgerUnmatched = {};
    var i;

    for (i = 0; i < bankRows.length; i += 1) bankUnmatched[i] = true;
    for (i = 0; i < ledgerRows.length; i += 1) ledgerUnmatched[i] = true;

    var ledgerIdMap = {};
    for (i = 0; i < ledgerRows.length; i += 1) {
      if (ledgerRows[i].txId) {
        if (!ledgerIdMap[ledgerRows[i].txId]) ledgerIdMap[ledgerRows[i].txId] = [];
        ledgerIdMap[ledgerRows[i].txId].push(i);
      }
    }

    for (i = 0; i < bankRows.length; i += 1) {
      var bank = bankRows[i];
      if (!bankUnmatched[i] || !bank.txId) continue;
      var candidateIndexes = ledgerIdMap[bank.txId] || [];
      for (var c = 0; c < candidateIndexes.length; c += 1) {
        var ledgerIndex = candidateIndexes[c];
        if (!ledgerUnmatched[ledgerIndex]) continue;
        var ledger = ledgerRows[ledgerIndex];
        if (isMatch(bank, ledger, amountTolerance, dateToleranceDays)) {
          matched.push(createMatchedRecord(bank, ledger, "id_match"));
          delete bankUnmatched[i];
          delete ledgerUnmatched[ledgerIndex];
          break;
        }
      }
    }

    var unmatchedBankIndexes = Object.keys(bankUnmatched).map(Number);
    unmatchedBankIndexes.forEach(function (bankIndex) {
      var bank = bankRows[bankIndex];
      var bestLedgerIndex = null;
      var bestScore = null;
      Object.keys(ledgerUnmatched).map(Number).forEach(function (ledgerIndex) {
        var ledger = ledgerRows[ledgerIndex];
        var amountDiff = Math.abs(bank.amount - ledger.amount);
        var dayDiff = dateDeltaDays(bank.txDate, ledger.txDate);
        if (amountDiff > amountTolerance || dayDiff > dateToleranceDays) return;
        var score = [amountDiff, dayDiff];
        if (!bestScore || compareScore(score, bestScore) < 0) {
          bestScore = score;
          bestLedgerIndex = ledgerIndex;
        }
      });
      if (bestLedgerIndex !== null) {
        matched.push(createMatchedRecord(bank, ledgerRows[bestLedgerIndex], "amount_date_match"));
        delete bankUnmatched[bankIndex];
        delete ledgerUnmatched[bestLedgerIndex];
      }
    });

    return {
      matched: matched,
      unmatchedBank: Object.keys(bankUnmatched).map(function (key) { return createUnmatchedRecord(bankRows[Number(key)], "bank_only"); }),
      unmatchedLedger: Object.keys(ledgerUnmatched).map(function (key) { return createUnmatchedRecord(ledgerRows[Number(key)], "ledger_only"); })
    };
  }

  function compareScore(a, b) {
    if (a[0] !== b[0]) return a[0] - b[0];
    return a[1] - b[1];
  }

  function isMatch(bank, ledger, amountTolerance, dateToleranceDays) {
    return Math.abs(bank.amount - ledger.amount) <= amountTolerance && dateDeltaDays(bank.txDate, ledger.txDate) <= dateToleranceDays;
  }

  function dateDeltaDays(a, b) {
    return Math.abs(Math.round((a.getTime() - b.getTime()) / 86400000));
  }

  function createMatchedRecord(bank, ledger, matchType) {
    return {
      matchType: matchType,
      bankRow: bank.index,
      ledgerRow: ledger.index,
      bankAmount: round(bank.amount, 2),
      ledgerAmount: round(ledger.amount, 2),
      difference: round(bank.amount - ledger.amount, 2),
      bankDate: bank.txDate.toISOString(),
      ledgerDate: ledger.txDate.toISOString(),
      bankTxId: bank.txId || "",
      ledgerTxId: ledger.txId || ""
    };
  }

  function createUnmatchedRecord(row, matchType) {
    return {
      matchType: matchType,
      row: row.index,
      amount: round(row.amount, 2),
      date: row.txDate.toISOString(),
      txId: row.txId || "",
      raw: row.raw
    };
  }

  function buildTaskFromResult(fileA, fileB, result) {
    var matchedTotal = result.matched.reduce(function (sum, item) { return sum + Number(item.bankAmount || 0); }, 0);
    var unmatchedTotal = result.unmatchedBank.length + result.unmatchedLedger.length;
    var status = unmatchedTotal ? "warning" : "success";
    return finalizeTask({
      id: generateTaskId(),
      createdAt: new Date().toISOString(),
      sourceAName: fileA.name,
      sourceBName: fileB.name,
      matchedCount: result.matched.length,
      unmatchedBankCount: result.unmatchedBank.length,
      unmatchedLedgerCount: result.unmatchedLedger.length,
      totalAmount: round(matchedTotal, 2),
      status: status,
      settingsSnapshot: {
        amountTolerance: Number(state.settings.amountTolerance || 0),
        dateToleranceDays: Number(state.settings.dateToleranceDays || 0)
      },
      matchedRecords: result.matched,
      unmatchedBankRecords: result.unmatchedBank,
      unmatchedLedgerRecords: result.unmatchedLedger
    });
  }

  function readFileAsText(file) {
    return new Promise(function (resolve, reject) {
      var reader = new FileReader();
      reader.onload = function () { resolve(String(reader.result || "")); };
      reader.onerror = reject;
      reader.readAsText(file, "utf-8");
    });
  }

  function readFileAsArrayBuffer(file) {
    return new Promise(function (resolve, reject) {
      var reader = new FileReader();
      reader.onload = function () { resolve(reader.result); };
      reader.onerror = reject;
      reader.readAsArrayBuffer(file);
    });
  }

  function downloadJson(filename, data) {
    triggerDownload(filename, new Blob([JSON.stringify(data, null, 2)], { type: "application/json" }));
  }

  function downloadCsv(filename, rows) {
    if (!rows.length) {
      triggerDownload(filename, new Blob([""], { type: "text/csv;charset=utf-8" }));
      return;
    }
    var headers = Object.keys(rows[0]);
    var content = [headers.join(",")].concat(rows.map(function (row) {
      return headers.map(function (header) {
        var value = row[header] == null ? "" : String(row[header]);
        return /[",\n]/.test(value) ? "\"" + value.replace(/"/g, "\"\"") + "\"" : value;
      }).join(",");
    })).join("\n");
    triggerDownload(filename, new Blob([content], { type: "text/csv;charset=utf-8" }));
  }

  function toMatchedExportRows(task) {
    return task.matchedRecords.map(function (row) {
      return {
        match_type: row.matchType,
        bank_row: row.bankRow,
        ledger_row: row.ledgerRow,
        bank_amount: row.bankAmount,
        ledger_amount: row.ledgerAmount,
        bank_date: row.bankDate ? formatDate(row.bankDate) : "",
        ledger_date: row.ledgerDate ? formatDate(row.ledgerDate) : "",
        bank_tx_id: row.bankTxId || "",
        ledger_tx_id: row.ledgerTxId || ""
      };
    });
  }

  function toUnmatchedExportRows(rows) {
    return rows.map(function (row) {
      var raw = Object.assign({}, row.raw || {});
      raw._row = row.row;
      raw._amount = row.amount;
      raw._date = row.date ? formatDate(row.date) : "";
      raw._tx_id = row.txId || "";
      return raw;
    });
  }

  function triggerDownload(filename, blob) {
    var url = URL.createObjectURL(blob);
    var link = document.createElement("a");
    link.href = url;
    link.download = filename;
    document.body.appendChild(link);
    link.click();
    link.remove();
    URL.revokeObjectURL(url);
  }

  function latestTask() {
    return state.tasks.slice().sort(sortByCreatedAt)[0] || null;
  }

  function aggregateMonth(monthValue) {
    var month = monthValue || new Date().toISOString().slice(0, 7);
    var tasks = state.tasks.filter(function (task) { return task.createdAt.slice(0, 7) === month; });
    var totalAmount = tasks.reduce(function (sum, task) { return sum + Number(task.totalAmount || 0); }, 0);
    var totalMatched = tasks.reduce(function (sum, task) { return sum + Number(task.matchedCount || 0); }, 0);
    var totalRecords = tasks.reduce(function (sum, task) { return sum + Number(task.matchedCount || 0) + Number(task.unmatchedBankCount || 0); }, 0);
    var discrepancies = tasks.reduce(function (sum, task) { return sum + Number(task.discrepancyCount || 0); }, 0);
    return {
      tasks: tasks,
      totalAmount: totalAmount,
      successRate: totalRecords ? round((totalMatched / totalRecords) * 100, 2) : 0,
      discrepancies: discrepancies
    };
  }

  function renderDashboard() {
    var latest = latestTask();
    var totalAmount = state.tasks.reduce(function (sum, task) { return sum + Number(task.totalAmount || 0); }, 0);
    var totalMatched = state.tasks.reduce(function (sum, task) { return sum + Number(task.matchedCount || 0); }, 0);
    var totalRecords = state.tasks.reduce(function (sum, task) { return sum + Number(task.matchedCount || 0) + Number(task.unmatchedBankCount || 0); }, 0);
    var discrepancies = state.tasks.reduce(function (sum, task) { return sum + Number(task.discrepancyCount || 0); }, 0);
    setText("dashboard-total-amount", formatMoney(totalAmount));
    setText("dashboard-match-rate", totalRecords ? round((totalMatched / totalRecords) * 100, 2) + "%" : "0%");
    setText("dashboard-pending", String(discrepancies));
    setText("latest-task-summary", latest ? latest.id + " / " + latest.sourceAName : "暂无任务");

    var tbody = document.getElementById("recent-task-body");
    if (!tbody) return;
    tbody.innerHTML = "";
    var tasks = state.tasks.slice().sort(sortByCreatedAt).slice(0, 8);
    if (!tasks.length) {
      tbody.innerHTML = "<tr><td colspan=\"8\"><div class=\"empty-state\">还没有本地任务记录。</div></td></tr>";
      return;
    }
    tasks.forEach(function (task) {
      var tr = document.createElement("tr");
      tr.innerHTML = [
        "<td>" + task.id + "</td>",
        "<td>" + formatDateTime(task.createdAt) + "</td>",
        "<td>" + task.sourceAName + "</td>",
        "<td>" + task.sourceBName + "</td>",
        "<td>" + formatMoney(task.totalAmount) + "</td>",
        "<td>" + renderBadge(task) + "</td>",
        "<td>" + task.accuracy + "%</td>",
        "<td><div class=\"table-actions\"><button class=\"inline-button\" data-export-summary=\"" + task.id + "\">导出摘要</button><button class=\"inline-button\" data-detail=\"" + task.id + "\">查看详情</button></div></td>"
      ].join("");
      tbody.appendChild(tr);
    });
  }

  function sortByCreatedAt(a, b) {
    return new Date(b.createdAt) - new Date(a.createdAt);
  }

  function renderAuditLog() {
    var searchValue = valueOf("log-search").toLowerCase();
    var statusValue = valueOf("log-status");
    var fromValue = valueOf("log-from");
    var toValue = valueOf("log-to");
    var rows = state.tasks.slice().sort(sortByCreatedAt).filter(function (task) {
      var created = task.createdAt.slice(0, 10);
      var matchesText = !searchValue || [task.id, task.sourceAName, task.sourceBName].join(" ").toLowerCase().indexOf(searchValue) !== -1;
      var matchesStatus = !statusValue || task.status === statusValue;
      var matchesFrom = !fromValue || created >= fromValue;
      var matchesTo = !toValue || created <= toValue;
      return matchesText && matchesStatus && matchesFrom && matchesTo;
    });
    setText("log-count", String(rows.length));
    var tbody = document.getElementById("log-table-body");
    if (!tbody) return;
    tbody.innerHTML = "";
    if (!rows.length) {
      tbody.innerHTML = "<tr><td colspan=\"8\"><div class=\"empty-state\">没有符合条件的任务。</div></td></tr>";
      return;
    }
    rows.forEach(function (task) {
      var tr = document.createElement("tr");
      tr.innerHTML = [
        "<td>" + task.id + "</td>",
        "<td>" + formatDateTime(task.createdAt) + "</td>",
        "<td>" + task.sourceAName + "</td>",
        "<td>" + task.sourceBName + "</td>",
        "<td>" + task.matchedCount + "</td>",
        "<td>" + task.discrepancyCount + "</td>",
        "<td>" + renderBadge(task) + "</td>",
        "<td><div class=\"table-actions\"><button class=\"inline-button\" data-export-matched=\"" + task.id + "\">导出匹配结果</button><button class=\"inline-button\" data-detail=\"" + task.id + "\">查看详情</button></div></td>"
      ].join("");
      tbody.appendChild(tr);
    });
  }

  function renderReport() {
    var select = document.getElementById("report-month");
    var activeMonth = select && select.value ? select.value : new Date().toISOString().slice(0, 7);
    if (select && !select.value) select.value = activeMonth;
    var summary = aggregateMonth(activeMonth);
    setText("report-total-amount", formatMoney(summary.totalAmount));
    setText("report-success-rate", summary.successRate + "%");
    setText("report-discrepancies", String(summary.discrepancies));
    var healthScore = Math.max(0, Math.min(10, round((summary.successRate / 100) * 10 - summary.discrepancies * 0.05, 1)));
    setText("report-health-score", healthScore.toFixed(1));
    renderReportChart(summary.tasks);
    renderReportTable(summary.tasks);
  }

  function renderReportChart(tasks) {
    var chart = document.getElementById("report-chart");
    if (!chart) return;
    chart.innerHTML = "";
    var weeklyBuckets = [0, 0, 0, 0];
    tasks.forEach(function (task) {
      var day = Number(task.createdAt.slice(8, 10));
      var weekIndex = Math.min(3, Math.floor((day - 1) / 7));
      weeklyBuckets[weekIndex] += Number(task.totalAmount || 0);
    });
    var maxValue = Math.max.apply(Math, weeklyBuckets.concat([1]));
    weeklyBuckets.forEach(function (value, index) {
      var bar = document.createElement("div");
      bar.className = "bar";
      if (value <= 0) {
        bar.classList.add("empty");
        bar.style.height = "0%";
      } else {
        bar.style.height = Math.max(18, (value / maxValue) * 100) + "%";
      }
      bar.innerHTML = "<span>第 " + (index + 1) + " 周</span>";
      chart.appendChild(bar);
    });
  }

  function renderReportTable(tasks) {
    var tbody = document.getElementById("report-table-body");
    if (!tbody) return;
    tbody.innerHTML = "";
    if (!tasks.length) {
      tbody.innerHTML = "<tr><td colspan=\"7\"><div class=\"empty-state\">该月份还没有任务数据。</div></td></tr>";
      return;
    }
    tasks.forEach(function (task) {
      var tr = document.createElement("tr");
      tr.innerHTML = [
        "<td>" + task.sourceAName + "</td>",
        "<td>" + task.sourceBName + "</td>",
        "<td>" + formatDateTime(task.createdAt) + "</td>",
        "<td>" + formatMoney(task.totalAmount) + "</td>",
        "<td>" + task.accuracy + "%</td>",
        "<td>" + task.discrepancyCount + "</td>",
        "<td>" + renderBadge(task) + "</td>"
      ].join("");
      tbody.appendChild(tr);
    });
  }

  function renderSettings() {
    setValue("setting-amount-tolerance", state.settings.amountTolerance);
    setValue("setting-date-tolerance", state.settings.dateToleranceDays);
    setValue("setting-auto-score", state.settings.autoMatchScore);
    setValue("setting-alert-delay", state.settings.alertDelayMinutes);
    setValue("setting-ingest-directory", state.settings.ingestDirectory);
    setValue("setting-archive-target", state.settings.archiveTarget);
    setValue("setting-primary-gateway", state.settings.primaryGateway);
    setValue("setting-backup-gateway", state.settings.backupGateway);
    setText("settings-summary", "金额容差 " + state.settings.amountTolerance + "，日期容差 " + state.settings.dateToleranceDays + " 天");
  }

  function renderTaskDetail(taskId) {
    var modal = document.getElementById("task-detail-modal");
    if (!modal) return;
    var task = getTaskById(taskId);
    if (!task) return;
    DETAIL_STATE.taskId = taskId;
    setText("detail-task-id", task.id);
    setText("detail-files", task.sourceAName + " / " + task.sourceBName);
    setText("detail-created-at", formatDateTime(task.createdAt));
    setText("detail-accuracy", task.accuracy + "%");
    setText("detail-summary-inline", "已匹配 " + task.matchedCount + " 条，银行侧未匹配 " + task.unmatchedBankCount + " 条，台账侧未匹配 " + task.unmatchedLedgerCount + " 条");
    renderDetailTable(task);
    modal.hidden = false;
    modal.style.display = "grid";
    document.body.classList.add("modal-open");
  }

  function renderDetailTable(task) {
    var tbody = document.getElementById("detail-table-body");
    if (!tbody) return;
    tbody.innerHTML = "";
    var filter = DETAIL_STATE.filter;
    var rows = [];
    if (filter === "all" || filter === "matched") rows = rows.concat(task.matchedRecords);
    if (filter === "all" || filter === "bank_only") rows = rows.concat(task.unmatchedBankRecords);
    if (filter === "all" || filter === "ledger_only") rows = rows.concat(task.unmatchedLedgerRecords);

    if (!rows.length) {
      tbody.innerHTML = "<tr><td colspan=\"8\"><div class=\"empty-state\">当前筛选下没有记录。</div></td></tr>";
      return;
    }

    rows.forEach(function (row) {
      var tr = document.createElement("tr");
      tr.innerHTML = [
        "<td>" + formatDetailType(row.matchType) + "</td>",
        "<td>" + (row.bankRow || row.row || "-") + "</td>",
        "<td>" + (row.ledgerRow || "-") + "</td>",
        "<td>" + formatMoney(row.bankAmount != null ? row.bankAmount : row.amount) + "</td>",
        "<td>" + (row.ledgerAmount != null ? formatMoney(row.ledgerAmount) : "-") + "</td>",
        "<td>" + formatMoney(row.difference != null ? row.difference : row.matchType === "ledger_only" ? -row.amount : row.amount) + "</td>",
        "<td>" + formatDate(row.bankDate || row.date) + "</td>",
        "<td>" + (row.bankTxId || row.ledgerTxId || row.txId || "-") + "</td>"
      ].join("");
      tbody.appendChild(tr);
    });
  }

  function formatDetailType(type) {
    if (type === "id_match") return "ID 匹配";
    if (type === "amount_date_match") return "金额+日期";
    if (type === "bank_only") return "仅银行侧";
    if (type === "ledger_only") return "仅台账侧";
    return type || "-";
  }

  function closeTaskDetail() {
    var modal = document.getElementById("task-detail-modal");
    if (!modal) return;
    modal.style.display = "none";
    modal.hidden = true;
    document.body.classList.remove("modal-open");
  }

  function renderBadge(task) {
    var type = task.status === "success" ? "success" : task.status === "error" ? "error" : "warn";
    return "<span class=\"badge " + type + "\">" + task.statusLabel + "</span>";
  }

  function setText(id, value) {
    var el = document.getElementById(id);
    if (el) el.textContent = value;
  }

  function setValue(id, value) {
    var el = document.getElementById(id);
    if (el) el.value = value;
  }

  function formatFileSize(bytes) {
    var size = Number(bytes || 0);
    if (size < 1024) return size + " B";
    if (size < 1024 * 1024) return round(size / 1024, 1) + " KB";
    return round(size / (1024 * 1024), 1) + " MB";
  }

  function fileTypeLabel(file) {
    var name = String(file.name || "").toLowerCase();
    if (name.endsWith(".csv")) return "CSV";
    if (name.endsWith(".json")) return "JSON";
    if (name.endsWith(".xlsx")) return "XLSX";
    return "未知格式";
  }

  function setFileSummary(inputId, summaryId, nameId, metaId, clearId) {
    var input = document.getElementById(inputId);
    var summary = document.getElementById(summaryId);
    var nameEl = document.getElementById(nameId);
    var metaEl = document.getElementById(metaId);
    var clearButton = document.getElementById(clearId);
    if (!input || !summary || !nameEl || !metaEl || !clearButton) return;
    var file = input.files && input.files[0];
    if (!file) {
      nameEl.textContent = "未选择文件";
      metaEl.textContent = "支持 CSV、JSON、XLSX";
      summary.classList.remove("selected");
      clearButton.disabled = true;
      return;
    }
    nameEl.textContent = file.name;
    metaEl.textContent = fileTypeLabel(file) + " | " + formatFileSize(file.size);
    summary.classList.add("selected");
    clearButton.disabled = false;
  }

  function updateRunButtonState() {
    var button = document.getElementById("run-reconcile");
    var fileAInput = document.getElementById("file-a");
    var fileBInput = document.getElementById("file-b");
    if (!button || !fileAInput || !fileBInput) return;
    var ready = fileAInput.files && fileAInput.files[0] && fileBInput.files && fileBInput.files[0];
    button.disabled = !ready;
  }

  function valueOf(id) {
    var el = document.getElementById(id);
    return el ? String(el.value || "") : "";
  }

  function getTaskById(taskId) {
    return state.tasks.find(function (task) { return task.id === taskId; }) || null;
  }

  function bindDashboard() {
    var runButton = document.getElementById("run-reconcile");
    var fileAInput = document.getElementById("file-a");
    var fileBInput = document.getElementById("file-b");
    if (fileAInput) {
      fileAInput.addEventListener("change", function () {
        setFileSummary("file-a", "file-a-summary", "file-a-name", "file-a-meta", "file-a-clear");
        updateRunButtonState();
      });
      setFileSummary("file-a", "file-a-summary", "file-a-name", "file-a-meta", "file-a-clear");
      var clearA = document.getElementById("file-a-clear");
      if (clearA) {
        clearA.addEventListener("click", function () {
          fileAInput.value = "";
          setFileSummary("file-a", "file-a-summary", "file-a-name", "file-a-meta", "file-a-clear");
          updateRunButtonState();
        });
      }
    }
    if (fileBInput) {
      fileBInput.addEventListener("change", function () {
        setFileSummary("file-b", "file-b-summary", "file-b-name", "file-b-meta", "file-b-clear");
        updateRunButtonState();
      });
      setFileSummary("file-b", "file-b-summary", "file-b-name", "file-b-meta", "file-b-clear");
      var clearB = document.getElementById("file-b-clear");
      if (clearB) {
        clearB.addEventListener("click", function () {
          fileBInput.value = "";
          setFileSummary("file-b", "file-b-summary", "file-b-name", "file-b-meta", "file-b-clear");
          updateRunButtonState();
        });
      }
    }
    updateRunButtonState();
    if (!runButton) return;
    runButton.addEventListener("click", function () {
      var fileA = fileAInput.files[0];
      var fileB = fileBInput.files[0];
      if (!fileA || !fileB) {
        alert("请选择两份本地文件后再运行对账。");
        return;
      }
      Promise.all([parseInput(fileA), parseInput(fileB)]).then(function (datasets) {
        var normalizedA = normalizeRows(datasets[0]);
        var normalizedB = normalizeRows(datasets[1]);
        if (!normalizedA.length || !normalizedB.length) {
          throw new Error("未识别到有效记录。请确认文件包含日期和金额字段。");
        }
        var result = reconcile(normalizedA, normalizedB);
        var task = buildTaskFromResult(fileA, fileB, result);
        state.tasks.unshift(task);
        saveState();
        renderAll();
        renderTaskDetail(task.id);
      }).catch(function (error) {
        console.error(error);
        alert(error.message || "本地对账失败。");
      });
    });
  }

  function parseInput(file) {
    var lower = file.name.toLowerCase();
    if (lower.endsWith(".json")) {
      return readFileAsText(file).then(function (text) {
        var json = JSON.parse(text);
        if (!Array.isArray(json)) throw new Error("JSON 文件必须是数组。");
        return json;
      });
    }
    if (lower.endsWith(".csv")) {
      return readFileAsText(file).then(parseCsv);
    }
    if (lower.endsWith(".xlsx")) {
      return readFileAsArrayBuffer(file).then(parseXlsx);
    }
    throw new Error("当前版本支持 CSV、JSON 和 XLSX。");
  }

  function bindAuditLog() {
    ["log-search", "log-status", "log-from", "log-to"].forEach(function (id) {
      var el = document.getElementById(id);
      if (el) {
        el.addEventListener("input", renderAuditLog);
        el.addEventListener("change", renderAuditLog);
      }
    });
    var reset = document.getElementById("log-reset");
    if (reset) {
      reset.addEventListener("click", function () {
        setValue("log-search", "");
        setValue("log-status", "");
        setValue("log-from", "");
        setValue("log-to", "");
        renderAuditLog();
      });
    }
  }

  function bindReport() {
    var month = document.getElementById("report-month");
    if (month) month.addEventListener("change", renderReport);
    var exportButton = document.getElementById("report-export");
    if (exportButton) {
      exportButton.addEventListener("click", function () {
        var currentMonth = valueOf("report-month") || new Date().toISOString().slice(0, 7);
        downloadJson("monthly-report-" + currentMonth + ".json", aggregateMonth(currentMonth));
      });
    }
  }

  function bindSettings() {
    var form = document.getElementById("settings-form");
    if (!form) return;
    form.addEventListener("submit", function (event) {
      event.preventDefault();
      state.settings.amountTolerance = Number(valueOf("setting-amount-tolerance") || 0);
      state.settings.dateToleranceDays = Number(valueOf("setting-date-tolerance") || 0);
      state.settings.autoMatchScore = Number(valueOf("setting-auto-score") || 0);
      state.settings.alertDelayMinutes = Number(valueOf("setting-alert-delay") || 0);
      state.settings.ingestDirectory = valueOf("setting-ingest-directory");
      state.settings.archiveTarget = valueOf("setting-archive-target");
      state.settings.primaryGateway = valueOf("setting-primary-gateway");
      state.settings.backupGateway = valueOf("setting-backup-gateway");
      saveState();
      renderSettings();
      alert("设置已保存。");
    });

    var clearButton = document.getElementById("clear-local-data");
    if (clearButton) {
      clearButton.addEventListener("click", function () {
        if (!window.confirm("确认清除本地任务和设置？此操作无法恢复。")) return;
        state = seedState();
        saveState();
        renderAll();
        closeTaskDetail();
      });
    }
  }

  function bindDetailModal() {
    var modal = document.getElementById("task-detail-modal");
    var close = document.getElementById("detail-close");
    if (close) close.addEventListener("click", closeTaskDetail);
    if (modal) {
      modal.addEventListener("click", function (event) {
        if (event.target === modal) closeTaskDetail();
      });
    }
    ["all", "matched", "bank_only", "ledger_only"].forEach(function (filter) {
      var button = document.querySelector("[data-detail-filter=\"" + filter + "\"]");
      if (button) {
        button.addEventListener("click", function () {
          DETAIL_STATE.filter = filter;
          updateDetailFilterButtons();
          if (DETAIL_STATE.taskId) renderDetailTable(getTaskById(DETAIL_STATE.taskId));
        });
      }
    });

    var exportSummary = document.getElementById("detail-export-summary");
    if (exportSummary) exportSummary.addEventListener("click", function () {
      var task = getTaskById(DETAIL_STATE.taskId);
      if (task) downloadJson(task.id + "-summary.json", task.summary);
    });

    var exportMatched = document.getElementById("detail-export-matched");
    if (exportMatched) exportMatched.addEventListener("click", function () {
      var task = getTaskById(DETAIL_STATE.taskId);
      if (task) downloadCsv(task.id + "-matched.csv", toMatchedExportRows(task));
    });

    var exportBank = document.getElementById("detail-export-bank");
    if (exportBank) exportBank.addEventListener("click", function () {
      var task = getTaskById(DETAIL_STATE.taskId);
      if (task) downloadCsv(task.id + "-unmatched-bank.csv", toUnmatchedExportRows(task.unmatchedBankRecords));
    });

    var exportLedger = document.getElementById("detail-export-ledger");
    if (exportLedger) exportLedger.addEventListener("click", function () {
      var task = getTaskById(DETAIL_STATE.taskId);
      if (task) downloadCsv(task.id + "-unmatched-ledger.csv", toUnmatchedExportRows(task.unmatchedLedgerRecords));
    });
  }

  function updateDetailFilterButtons() {
    var buttons = document.querySelectorAll("[data-detail-filter]");
    buttons.forEach(function (button) {
      button.classList.toggle("active-chip", button.getAttribute("data-detail-filter") === DETAIL_STATE.filter);
    });
  }

  function bindGlobalActions() {
    document.body.addEventListener("click", function (event) {
      var trigger = event.target.closest("[data-detail],[data-export-summary],[data-export-matched]");
      if (!trigger) return;
      var detailId = trigger.getAttribute("data-detail");
      var summaryId = trigger.getAttribute("data-export-summary");
      var matchedId = trigger.getAttribute("data-export-matched");
      if (detailId) {
        DETAIL_STATE.filter = "all";
        updateDetailFilterButtons();
        renderTaskDetail(detailId);
      }
      if (summaryId) {
        var summaryTask = getTaskById(summaryId);
        if (summaryTask) downloadJson(summaryTask.id + "-summary.json", summaryTask.summary);
      }
      if (matchedId) {
        var matchedTask = getTaskById(matchedId);
        if (matchedTask) downloadCsv(matchedTask.id + "-matched.csv", toMatchedExportRows(matchedTask));
      }
    });
  }

  async function parseXlsx(arrayBuffer) {
    var files = await unzipEntries(arrayBuffer);
    var workbookXml = files["xl/workbook.xml"];
    var workbookRelsXml = files["xl/_rels/workbook.xml.rels"];
    if (!workbookXml || !workbookRelsXml) throw new Error("XLSX 结构不完整，缺少 workbook 信息。");
    var sharedStrings = files["xl/sharedStrings.xml"] ? parseSharedStrings(files["xl/sharedStrings.xml"]) : [];
    var styleInfo = files["xl/styles.xml"] ? parseStyles(files["xl/styles.xml"]) : { xfDateFlags: [] };
    var targetSheetPath = resolveFirstWorksheetPath(workbookXml, workbookRelsXml);
    var sheetXml = files[targetSheetPath];
    if (!sheetXml) throw new Error("XLSX 未找到工作表内容。");
    return parseWorksheet(sheetXml, sharedStrings, styleInfo);
  }

  async function unzipEntries(arrayBuffer) {
    var bytes = new Uint8Array(arrayBuffer);
    var view = new DataView(arrayBuffer);
    var eocdOffset = findEocdOffset(bytes);
    if (eocdOffset === -1) throw new Error("无法识别 XLSX 压缩包。");
    var centralDirOffset = view.getUint32(eocdOffset + 16, true);
    var totalEntries = view.getUint16(eocdOffset + 10, true);
    var offset = centralDirOffset;
    var files = {};

    for (var i = 0; i < totalEntries; i += 1) {
      if (view.getUint32(offset, true) !== 0x02014b50) break;
      var compressionMethod = view.getUint16(offset + 10, true);
      var compressedSize = view.getUint32(offset + 20, true);
      var fileNameLength = view.getUint16(offset + 28, true);
      var extraLength = view.getUint16(offset + 30, true);
      var commentLength = view.getUint16(offset + 32, true);
      var localHeaderOffset = view.getUint32(offset + 42, true);
      var fileName = decodeText(bytes.slice(offset + 46, offset + 46 + fileNameLength));
      offset += 46 + fileNameLength + extraLength + commentLength;
      if (/\/$/.test(fileName)) continue;
      files[fileName] = await extractZipEntry(view, bytes, localHeaderOffset, compressedSize, compressionMethod);
    }
    return files;
  }

  function findEocdOffset(bytes) {
    for (var i = bytes.length - 22; i >= 0; i -= 1) {
      if (bytes[i] === 0x50 && bytes[i + 1] === 0x4b && bytes[i + 2] === 0x05 && bytes[i + 3] === 0x06) return i;
    }
    return -1;
  }

  async function extractZipEntry(view, bytes, localHeaderOffset, compressedSize, compressionMethod) {
    if (view.getUint32(localHeaderOffset, true) !== 0x04034b50) throw new Error("XLSX 本地文件头损坏。");
    var fileNameLength = view.getUint16(localHeaderOffset + 26, true);
    var extraLength = view.getUint16(localHeaderOffset + 28, true);
    var dataStart = localHeaderOffset + 30 + fileNameLength + extraLength;
    var compressed = bytes.slice(dataStart, dataStart + compressedSize);
    if (compressionMethod === 0) return decodeText(compressed);
    if (compressionMethod === 8) return decodeText(await inflateRaw(compressed));
    throw new Error("当前 XLSX 压缩方式不支持: " + compressionMethod);
  }

  async function inflateRaw(bytes) {
    if (typeof DecompressionStream === "undefined") {
      throw new Error("当前浏览器暂不支持 XLSX 解压，请使用较新的 Chrome、Edge 或 Safari。");
    }
    var stream = new Blob([bytes]).stream().pipeThrough(new DecompressionStream("deflate-raw"));
    var response = new Response(stream);
    return new Uint8Array(await response.arrayBuffer());
  }

  function decodeText(bytes) {
    return new TextDecoder("utf-8").decode(bytes);
  }

  function parseSharedStrings(xmlText) {
    var doc = new DOMParser().parseFromString(xmlText, "application/xml");
    return Array.prototype.slice.call(doc.getElementsByTagName("si")).map(function (node) {
      var texts = node.getElementsByTagName("t");
      return Array.prototype.slice.call(texts).map(function (t) { return t.textContent || ""; }).join("");
    });
  }

  function parseStyles(xmlText) {
    var doc = new DOMParser().parseFromString(xmlText, "application/xml");
    var customFormats = {};
    Array.prototype.slice.call(doc.getElementsByTagName("numFmt")).forEach(function (node) {
      customFormats[node.getAttribute("numFmtId")] = node.getAttribute("formatCode") || "";
    });
    var xfDateFlags = Array.prototype.slice.call(doc.getElementsByTagName("cellXfs")[0] ? doc.getElementsByTagName("cellXfs")[0].getElementsByTagName("xf") : []).map(function (node) {
      var numFmtId = Number(node.getAttribute("numFmtId") || 0);
      if (BUILTIN_DATE_FORMAT_IDS[numFmtId]) return true;
      var custom = customFormats[String(numFmtId)] || "";
      return /[ymdhHs]/.test(custom);
    });
    return { xfDateFlags: xfDateFlags };
  }

  function resolveFirstWorksheetPath(workbookXml, workbookRelsXml) {
    var workbookDoc = new DOMParser().parseFromString(workbookXml, "application/xml");
    var relDoc = new DOMParser().parseFromString(workbookRelsXml, "application/xml");
    var sheet = workbookDoc.getElementsByTagName("sheet")[0];
    if (!sheet) throw new Error("XLSX 中没有工作表。");
    var relId = sheet.getAttribute("r:id") || sheet.getAttribute("id");
    var rels = relDoc.getElementsByTagName("Relationship");
    for (var i = 0; i < rels.length; i += 1) {
      if (rels[i].getAttribute("Id") === relId) {
        return "xl/" + rels[i].getAttribute("Target").replace(/^\//, "").replace(/^xl\//, "");
      }
    }
    throw new Error("XLSX 工作表关系解析失败。");
  }

  function parseWorksheet(sheetXml, sharedStrings, styleInfo) {
    var doc = new DOMParser().parseFromString(sheetXml, "application/xml");
    var rowNodes = Array.prototype.slice.call(doc.getElementsByTagName("row"));
    var rows = [];
    rowNodes.forEach(function (rowNode) {
      var cells = [];
      Array.prototype.slice.call(rowNode.getElementsByTagName("c")).forEach(function (cellNode) {
        var ref = cellNode.getAttribute("r") || "";
        var colIndex = columnLettersToIndex(ref.replace(/[0-9]/g, ""));
        cells[colIndex] = extractCellValue(cellNode, sharedStrings, styleInfo);
      });
      rows.push(cells);
    });
    return rowsToObjects(rows);
  }

  function columnLettersToIndex(letters) {
    var total = 0;
    for (var i = 0; i < letters.length; i += 1) {
      total = total * 26 + (letters.charCodeAt(i) - 64);
    }
    return Math.max(0, total - 1);
  }

  function extractCellValue(cellNode, sharedStrings, styleInfo) {
    var type = cellNode.getAttribute("t") || "n";
    var styleIndex = Number(cellNode.getAttribute("s") || 0);
    var valueNode = cellNode.getElementsByTagName("v")[0];
    var inlineNode = cellNode.getElementsByTagName("is")[0];
    var raw = valueNode ? valueNode.textContent : inlineNode ? inlineNode.textContent : "";
    if (type === "s") return sharedStrings[Number(raw || 0)] || "";
    if (type === "inlineStr") return raw || "";
    if (type === "b") return raw === "1";
    if (type === "n") {
      var numeric = Number(raw || 0);
      if (styleInfo.xfDateFlags[styleIndex]) return excelSerialToDate(numeric);
      return numeric;
    }
    return raw || "";
  }

  function bindEscapeToClose() {
    document.addEventListener("keydown", function (event) {
      if (event.key === "Escape") closeTaskDetail();
    });
  }

  function renderAll() {
    renderDashboard();
    renderAuditLog();
    renderReport();
    renderSettings();
  }

  document.addEventListener("DOMContentLoaded", function () {
    bindDashboard();
    bindAuditLog();
    bindReport();
    bindSettings();
    bindDetailModal();
    bindGlobalActions();
    bindEscapeToClose();
    renderAll();
  });
})();
