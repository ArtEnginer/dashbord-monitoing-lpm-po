/* global XLSX, Chart, DataTable */
(() => {
  const STORAGE_KEY = "echa-dashboard-state-v1";

  const state = {
    lpmRows: [],
    poRows: [],
    poDetails: [],
    lpmHandle: null,
    poHandle: null,
    charts: {},
    tables: { po: null, lpm: null },
    loaded: { lpm: false, po: false },
    isLoading: false,
    rendered: { po: [], lpm: [] },
    selectedYear: null,
  };

  const MONTHS_ID = [
    "Jan",
    "Feb",
    "Mar",
    "Apr",
    "Mei",
    "Jun",
    "Jul",
    "Agu",
    "Sep",
    "Okt",
    "Nov",
    "Des",
  ];
  const dom = {};

  function $(id) {
    return document.getElementById(id);
  }

  function escapeHtml(value) {
    return String(value ?? "").replace(
      /[&<>"]/g,
      (character) =>
        ({
          "&": "&amp;",
          "<": "&lt;",
          ">": "&gt;",
          '"': "&quot;",
        })[character],
    );
  }

  function fmtIDR(value) {
    const numeric = Number(value);
    if (!Number.isFinite(numeric)) return "Rp 0";
    return "Rp " + Math.round(numeric).toLocaleString("id-ID");
  }

  function fmtNumber(value) {
    const numeric = Number(value);
    if (!Number.isFinite(numeric)) return "-";
    return numeric.toLocaleString("id-ID");
  }

  function setKpiValue(value) {
    const numeric = Number(value);
    if (!Number.isFinite(numeric)) {
      if (dom.kpiValue)
        dom.kpiValue.innerHTML = `<span class="currency">Rp</span> <span class="amount">0</span>`;
      return;
    }
    const amount = Math.round(numeric).toLocaleString("id-ID");
    if (dom.kpiValue)
      dom.kpiValue.innerHTML = `<span class="currency">Rp</span> <span class="amount">${escapeHtml(amount)}</span>`;
  }

  // SVG icons for expand / collapse
  function expandIcon() {
    return `
      <svg viewBox="0 0 24 24" width="16" height="16" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">
        <path d="M8 3H3v5" />
        <path d="M16 3h5v5" />
        <path d="M16 21h5v-5" />
        <path d="M8 21H3v-5" />
      </svg>
    `;
  }

  function collapseIcon() {
    return `
      <svg viewBox="0 0 24 24" width="16" height="16" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">
        <path d="M8 6h8" />
        <path d="M8 18h8" />
        <path d="M12 6v12" />
      </svg>
    `;
  }

  function setBtnToExpand(btn) {
    try {
      btn.innerHTML = expandIcon();
      btn.setAttribute("aria-pressed", "false");
    } catch (e) {}
  }

  function setBtnToCollapse(btn) {
    try {
      btn.innerHTML = collapseIcon();
      btn.setAttribute("aria-pressed", "true");
    } catch (e) {}
  }

  function toISODate(value) {
    if (value === null || value === undefined || value === "") return null;
    if (value instanceof Date && !Number.isNaN(value.getTime())) {
      return value.toISOString().slice(0, 10);
    }
    if (typeof value === "number") {
      const parsed = XLSX.SSF.parse_date_code(value);
      if (!parsed) return null;
      return `${parsed.y}-${String(parsed.m).padStart(2, "0")}-${String(parsed.d).padStart(2, "0")}`;
    }
    if (typeof value === "string") {
      const ddmmyyyy = value.match(/(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})/);
      if (ddmmyyyy) {
        let [, day, month, year] = ddmmyyyy;
        if (year.length === 2) year = "20" + year;
        return `${year}-${String(month).padStart(2, "0")}-${String(day).padStart(2, "0")}`;
      }
      const parsed = new Date(value);
      if (!Number.isNaN(parsed.getTime())) {
        return parsed.toISOString().slice(0, 10);
      }
    }
    return null;
  }

  function fmtDateDisplay(iso) {
    if (!iso) return "-";
    const [year, month, day] = String(iso).split("-");
    return `${day}/${month}/${year}`;
  }

  function monthKey(iso) {
    if (!iso) return null;
    const [year, month] = String(iso).split("-");
    return `${year}-${month}`;
  }

  function normKey(value) {
    return String(value ?? "")
      .trim()
      .toUpperCase()
      .replace(/\s+/g, " ");
  }

  function statusClass(value) {
    return normKey(value).toLowerCase().replace(/\s+/g, "-");
  }

  function getUniqueLpmNumbers() {
    return [
      ...new Set(
        state.lpmRows.map((row) => String(row.no ?? "").trim()).filter(Boolean),
      ),
    ];
  }

  function getAvailableYears() {
    const years = new Set();
    state.lpmRows.forEach((row) => {
      const key = monthKey(row.tanggal);
      if (key) {
        const year = key.split("-")[0];
        years.add(year);
      }
    });
    state.poRows.forEach((row) => {
      const key = monthKey(row.tanggal);
      if (key) {
        const year = key.split("-")[0];
        years.add(year);
      }
    });
    return Array.from(years).sort();
  }

  function populateYearSelector() {
    if (!dom.yearSelect) return;
    const years = getAvailableYears();
    const selectedYear = state.selectedYear;

    // Clear existing options except the first one
    while (dom.yearSelect.options.length > 1) {
      dom.yearSelect.remove(1);
    }

    if (years.length === 0) {
      dom.yearSelect.disabled = true;
      return;
    }

    dom.yearSelect.disabled = false;

    years.forEach((year) => {
      const option = document.createElement("option");
      option.value = year;
      option.textContent = year;
      if (year === selectedYear) {
        option.selected = true;
      }
      dom.yearSelect.appendChild(option);
    });

    // Set default to first year if not set
    if (!state.selectedYear && years.length > 0) {
      state.selectedYear = years[0];
      dom.yearSelect.value = years[0];
    }
  }

  function sheetToRows(worksheet) {
    return XLSX.utils.sheet_to_json(worksheet, {
      header: 1,
      raw: true,
      defval: null,
    });
  }

  function findCol(headerRow, ...candidates) {
    const normalized = headerRow.map(normKey);
    for (const candidate of candidates) {
      const index = normalized.indexOf(normKey(candidate));
      if (index !== -1) return index;
    }
    return -1;
  }

  function setStatus(text, active) {
    dom.statusText.textContent = text;
    dom.statusDot.classList.toggle("on", Boolean(active));
    if (dom.statusMessage) {
      dom.statusMessage.textContent =
        text === "Data siap"
          ? "File sudah terbaca. Gunakan filter, pilih tahun, atau geser grafik untuk menelusuri data lebih cepat."
          : text === "Sebagian data dimuat"
            ? "Satu file sudah siap. Unggah file pasangan supaya ringkasan dan grafik lengkap muncul."
            : "Upload file LPM dan PO untuk memulai monitoring.";
    }
  }

  function getFileName(value) {
    if (!value) return "";
    if (typeof value === "string") return value;
    return value.name || "";
  }

  function setFileLabel(target, file) {
    target.textContent = getFileName(file) || "-";
  }

  function setFileChip(target, file) {
    const name = getFileName(file);
    target.innerHTML = name
      ? `<div class="uploaded-file">File siap: ${escapeHtml(name)}</div>`
      : "";
  }

  function saveStateToStorage() {
    try {
      const snapshot = {
        loaded: state.loaded,
        fileNames: state.fileNames,
        lpmRows: state.lpmRows,
        poRows: state.poRows,
        poDetails: state.poDetails,
        lastUpdated: new Date().toISOString(),
      };
      window.localStorage.setItem(STORAGE_KEY, JSON.stringify(snapshot));
    } catch (error) {
      console.warn("Tidak bisa menyimpan state dashboard:", error);
    }
  }

  function clearStateStorage() {
    try {
      window.localStorage.removeItem(STORAGE_KEY);
    } catch (error) {
      console.warn("Tidak bisa menghapus state dashboard:", error);
    }
  }

  function restoreStateFromStorage() {
    try {
      const raw = window.localStorage.getItem(STORAGE_KEY);
      if (!raw) return false;
      const snapshot = JSON.parse(raw);
      if (!snapshot) return false;

      state.loaded = {
        lpm: Boolean(snapshot.loaded?.lpm),
        po: Boolean(snapshot.loaded?.po),
      };
      state.fileNames = snapshot.fileNames || { lpm: "-", po: "-" };
      state.lpmRows = Array.isArray(snapshot.lpmRows) ? snapshot.lpmRows : [];
      state.poRows = Array.isArray(snapshot.poRows) ? snapshot.poRows : [];
      state.poDetails = Array.isArray(snapshot.poDetails)
        ? snapshot.poDetails
        : [];

      if (state.loaded.lpm) {
        setFileLabel(dom.uploadFileLpm, state.fileNames.lpm);
        setFileChip(dom.fileChipLpm, state.fileNames.lpm);
        dom.uploaderLpm.classList.add("loaded");
      }
      if (state.loaded.po) {
        setFileLabel(dom.uploadFilePo, state.fileNames.po);
        setFileChip(dom.fileChipPo, state.fileNames.po);
        dom.uploaderPo.classList.add("loaded");
      }
      return true;
    } catch (error) {
      console.warn("Tidak bisa memuat state dashboard:", error);
      return false;
    }
  }

  function updateLoadedState() {
    document.body.classList.toggle("is-loading", Boolean(state.isLoading));
    if (dom.uploadGrid) {
      dom.uploadGrid.classList.toggle(
        "hidden",
        state.loaded.lpm && state.loaded.po,
      );
    }
    if (state.isLoading) {
      dom.uploadNote.textContent = "Memproses file, mohon tunggu...";
    } else {
      dom.uploadNote.textContent =
        state.loaded.lpm && state.loaded.po
          ? "Data sudah dimuat di browser. Gunakan Muat Ulang atau Ganti File bila perlu."
          : "Unggah file LPM dan PO. Data diproses hanya di browser Anda.";
    }
    dom.controlBar.classList.toggle(
      "hidden",
      !(state.loaded.lpm || state.loaded.po),
    );
    dom.dashboard.classList.toggle(
      "hidden",
      !(state.loaded.lpm && state.loaded.po),
    );

    if (state.loaded.lpm || state.loaded.po) {
      dom.controlBar.classList.remove("hidden");
      dom.controlFileLpm.textContent = state.loaded.lpm
        ? state.fileNames.lpm
        : "-";
      dom.controlFilePo.textContent = state.loaded.po
        ? state.fileNames.po
        : "-";
      dom.statusFileLpm.textContent = state.loaded.lpm
        ? state.fileNames.lpm
        : "-";
      dom.statusFilePo.textContent = state.loaded.po ? state.fileNames.po : "-";
      dom.lastUpdated.textContent =
        "Diperbarui: " +
        new Date().toLocaleTimeString("id-ID", { hour12: false });
    }

    if (!state.loaded.lpm && !state.loaded.po) {
      setStatus("Menunggu data", false);
    } else if (!(state.loaded.lpm && state.loaded.po)) {
      setStatus("Sebagian data dimuat", false);
    } else {
      setStatus("Data siap", true);
    }
  }

  state.fileNames = { lpm: "-", po: "-" };

  function parseLpmWorkbook(workbook) {
    let rows = null;
    for (const sheetName of workbook.SheetNames) {
      const worksheetRows = sheetToRows(workbook.Sheets[sheetName]);
      if (!worksheetRows.length) continue;
      const header = worksheetRows[0].map(normKey);
      if (header.some((entry) => entry.includes("NO LPM"))) {
        rows = worksheetRows;
        break;
      }
    }

    if (!rows) {
      throw new Error('Sheet LPM tidak ditemukan. Header "NO LPM" tidak ada.');
    }

    const header = rows[0];
    const colNo = findCol(header, "NO LPM");
    const colDate = findCol(header, "TANGGAL PENGAJUAN", "TANGGAL");
    const colBy = findCol(header, "DIBUAT OLEH");
    const colItem = findCol(header, "ITEM");
    const colUnit = findCol(header, "SATUAN");
    const colVolume = findCol(header, "VOLUME");
    const colStatus = findCol(header, "STATUS");

    return rows
      .slice(1)
      .filter(
        (row) =>
          row &&
          row[colNo] !== null &&
          row[colNo] !== undefined &&
          row[colNo] !== "",
      )
      .map((row) => ({
        no: row[colNo],
        tanggal: toISODate(row[colDate]),
        dibuat: row[colBy] || "",
        item: String(row[colItem] || "").trim(),
        satuan: row[colUnit] || "",
        volume: row[colVolume],
        status: String(row[colStatus] || "")
          .trim()
          .toUpperCase(),
      }));
  }

  function parsePoWorkbook(workbook) {
    const poSheetName =
      workbook.SheetNames.find((name) => normKey(name) === "PO") ||
      workbook.SheetNames.find((name) => {
        const rows = sheetToRows(workbook.Sheets[name]);
        return (
          rows.length &&
          rows[0].map(normKey).some((entry) => entry.includes("PO NUMBER"))
        );
      });

    const detailSheetName =
      workbook.SheetNames.find((name) => normKey(name).includes("DETAIL")) ||
      workbook.SheetNames.find((name) => {
        const rows = sheetToRows(workbook.Sheets[name]);
        return (
          rows.length &&
          rows[0].map(normKey).some((entry) => entry.includes("DESC"))
        );
      });

    if (!poSheetName) {
      throw new Error(
        'Sheet PO tidak ditemukan. Header "PO NUMBER" tidak ada.',
      );
    }

    const poRows = sheetToRows(workbook.Sheets[poSheetName]);
    const header = poRows[0] || [];
    const colDate = findCol(header, "PO DATE");
    const colNo = findCol(header, "PO NUMBER");
    const colVendor = findCol(header, "VENDOR");
    const colSubtotal = findCol(header, "SUBTOTAL");
    const colPpn = findCol(header, "PPN");
    const colFreight = findCol(header, "FREIGHT ESTIMATION", "FREIGHT");
    const colTotal = findCol(header, "TOTAL");

    const poRowsParsed = poRows
      .slice(1)
      .filter((row) => row && row[colNo])
      .map((row) => ({
        tanggal: toISODate(row[colDate]),
        no: row[colNo],
        vendor: String(row[colVendor] || "").trim(),
        subtotal: Number(row[colSubtotal]) || 0,
        ppn: Number(row[colPpn]) || 0,
        freight: Number(row[colFreight]) || 0,
        total: Number(row[colTotal]) || 0,
      }));

    const poDetails = [];
    if (detailSheetName) {
      const detailRows = sheetToRows(workbook.Sheets[detailSheetName]);
      const detailHeader = detailRows[0] || [];
      const colDetailNo = findCol(detailHeader, "PO NUMBER");
      const colDesc = findCol(detailHeader, "DESC");
      const colQty = findCol(detailHeader, "QTY");
      const colAmount = findCol(detailHeader, "AMOUNT");

      detailRows.slice(1).forEach((row) => {
        if (!row || !row[colDetailNo]) return;
        poDetails.push({
          no: row[colDetailNo],
          desc: String(row[colDesc] || "").trim(),
          qty: row[colQty],
          amount: row[colAmount],
        });
      });
    }

    return { poRows: poRowsParsed, poDetails };
  }

  async function readWorkbook(file) {
    const buffer = await file.arrayBuffer();
    return XLSX.read(buffer, { type: "array", cellDates: false });
  }

  function destroyChart(name) {
    if (state.charts[name]) {
      state.charts[name].destroy();
      delete state.charts[name];
    }
  }

  function destroyTables() {
    Object.values(state.tables).forEach((table) => {
      if (table && typeof table.destroy === "function") {
        table.destroy();
      }
    });
    state.tables = { po: null, lpm: null };
  }

  function hasDataTable() {
    return typeof window.DataTable === "function";
  }

  function hasChartLibrary() {
    return typeof window.Chart === "function";
  }

  function renderPlainTable(tableElement, rows, type) {
    const tbody = tableElement.querySelector("tbody");
    if (!tbody) return;

    if (type === "po") {
      tbody.innerHTML = rows
        .map(
          (record) => `
            <tr>
              <td><span class="strong">${escapeHtml(record.no || "-")}</span></td>
              <td>${escapeHtml(fmtDateDisplay(record.tanggal))}</td>
              <td>${escapeHtml(record.vendor || "-")}</td>
              <td>${escapeHtml(record.itemSummary && record.itemSummary !== "-" ? record.itemSummary : "-")}</td>
              <td class="dt-body-right">${escapeHtml(fmtIDR(Number(record.total) || 0))}</td>
            </tr>
          `,
        )
        .join("");
      return;
    }

    tbody.innerHTML = rows
      .map((record) => {
        const badge = record.status
          ? `<span class="badge ${statusClass(record.status)}">${escapeHtml(record.status)}</span>`
          : "-";
        return `
          <tr>
            <td><span class="strong">${escapeHtml(record.no || "-")}</span></td>
            <td>${escapeHtml(fmtDateDisplay(record.tanggal))}</td>
            <td>${escapeHtml(record.dibuat || "-")}</td>
            <td>${escapeHtml(record.item || "-")}</td>
            <td class="dt-body-right">${escapeHtml(`${fmtNumber(record.volume)}${record.satuan ? ` ${record.satuan}` : ""}`)}</td>
            <td>${badge}</td>
          </tr>
        `;
      })
      .join("");
  }

  function applyPlainTableFilters() {
    const poQuery = dom.poSearch.value.trim().toLowerCase();
    const poVendor = dom.poVendorFilter.value.trim().toLowerCase();
    const lpmQuery = dom.lpmSearch.value.trim().toLowerCase();
    const lpmStatus = dom.lpmStatusFilter.value.trim().toLowerCase();

    const filteredPo = state.rendered.po.filter((record) => {
      const haystack = [
        record.no,
        record.vendor,
        record.itemSummary,
        fmtDateDisplay(record.tanggal),
      ]
        .join(" ")
        .toLowerCase();
      const matchesQuery = !poQuery || haystack.includes(poQuery);
      const matchesVendor =
        !poVendor || normKey(record.vendor).toLowerCase() === poVendor;
      return matchesQuery && matchesVendor;
    });

    const filteredLpm = state.rendered.lpm.filter((record) => {
      const haystack = [
        record.no,
        record.item,
        record.dibuat,
        fmtDateDisplay(record.tanggal),
      ]
        .join(" ")
        .toLowerCase();
      const matchesQuery = !lpmQuery || haystack.includes(lpmQuery);
      const matchesStatus =
        !lpmStatus || normKey(record.status).toLowerCase() === lpmStatus;
      return matchesQuery && matchesStatus;
    });

    renderPlainTable(dom.poTable, filteredPo, "po");
    renderPlainTable(dom.lpmTable, filteredLpm, "lpm");
  }

  function renderCharts(poEnriched) {
    if (!hasChartLibrary()) {
      return;
    }
    const totalLpm = getUniqueLpmNumbers().length;
    const totalPo = poEnriched.length;
    const uniqueVendors = [
      ...new Set(poEnriched.map((item) => item.vendor).filter(Boolean)),
    ];
    const totalValue = poEnriched.reduce(
      (sum, item) => sum + (Number(item.total) || 0),
      0,
    );

    dom.kpiLpm.textContent = totalLpm.toLocaleString("id-ID");
    dom.kpiPo.textContent = totalPo.toLocaleString("id-ID");
    dom.kpiVendor.textContent = uniqueVendors.length.toLocaleString("id-ID");
    setKpiValue(totalValue);
    dom.kpiValueNote.textContent =
      totalValue > 0
        ? `Total gabungan sekitar ${(totalValue / 1000000).toFixed(2)} juta rupiah`
        : "Belum ada nilai PO yang terbaca";

    Chart.defaults.font.family = '"Plus Jakarta Sans", sans-serif';
    Chart.defaults.color = "#556174";
    const gridColor = "rgba(148, 163, 184, 0.24)";

    const lpmMonthMap = {};
    getUniqueLpmNumbers().forEach((no) => {
      const record = state.lpmRows.find(
        (row) => String(row.no ?? "").trim() === no,
      );
      const key = monthKey(record && record.tanggal);
      if (!key) return;
      const [year] = key.split("-");
      if (year === state.selectedYear) {
        lpmMonthMap[key] = (lpmMonthMap[key] || 0) + 1;
      }
    });

    // Create full 12 months array for selected year
    const lpmMonthLabels = [];
    const lpmMonthData = [];
    for (let month = 1; month <= 12; month++) {
      const key = `${state.selectedYear}-${String(month).padStart(2, "0")}`;
      lpmMonthLabels.push(MONTHS_ID[month - 1]);
      lpmMonthData.push(lpmMonthMap[key] || 0);
    }

    destroyChart("lpm");
    state.charts.lpm = new Chart(dom.chartLpm, {
      type: "line",
      data: {
        labels: lpmMonthLabels,
        datasets: [
          {
            data: lpmMonthData,
            borderColor: "#3b82f6",
            backgroundColor: "rgba(59, 130, 246, 0.1)",
            borderWidth: 3,
            pointRadius: 5,
            pointBackgroundColor: "#3b82f6",
            pointBorderColor: "#fff",
            pointBorderWidth: 2,
            pointHoverRadius: 7,
            tension: 0.4,
            fill: true,
          },
        ],
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: { display: false },
          zoom: {
            zoom: {
              wheel: { enabled: false, speed: 0.1 },
              pinch: { enabled: false },
              mode: "xy",
            },
            pan: {
              enabled: true,
              mode: "xy",
            },
          },
        },
        scales: {
          x: { grid: { display: false } },
          y: {
            beginAtZero: true,
            ticks: { precision: 0 },
            grid: { color: gridColor },
          },
        },
      },
    });

    const poMonthMap = {};
    poEnriched.forEach((record) => {
      const key = monthKey(record.tanggal);
      if (!key) return;
      const [year] = key.split("-");
      if (year === state.selectedYear) {
        poMonthMap[key] = (poMonthMap[key] || 0) + 1;
      }
    });

    // Create full 12 months array for selected year
    const poMonthLabels = [];
    const poMonthData = [];
    for (let month = 1; month <= 12; month++) {
      const key = `${state.selectedYear}-${String(month).padStart(2, "0")}`;
      poMonthLabels.push(MONTHS_ID[month - 1]);
      poMonthData.push(poMonthMap[key] || 0);
    }

    destroyChart("po");
    state.charts.po = new Chart(dom.chartPo, {
      type: "line",
      data: {
        labels: poMonthLabels,
        datasets: [
          {
            data: poMonthData,
            borderColor: "#06b6d4",
            backgroundColor: "rgba(6, 182, 212, 0.1)",
            borderWidth: 3,
            pointRadius: 5,
            pointBackgroundColor: "#06b6d4",
            pointBorderColor: "#fff",
            pointBorderWidth: 2,
            pointHoverRadius: 7,
            tension: 0.4,
            fill: true,
          },
        ],
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: { display: false },
          zoom: {
            zoom: {
              wheel: { enabled: false, speed: 0.1 },
              pinch: { enabled: false },
              mode: "xy",
            },
            pan: {
              enabled: true,
              mode: "xy",
            },
          },
        },
        scales: {
          x: { grid: { display: false } },
          y: {
            beginAtZero: true,
            ticks: { precision: 0 },
            grid: { color: gridColor },
          },
        },
      },
    });

    const statusCount = { DITERIMA: 0, "ON PROSES": 0, CANCEL: 0 };
    state.lpmRows.forEach((row) => {
      const key = normKey(row.status);
      if (statusCount[key] === undefined) statusCount[key] = 0;
      statusCount[key] += 1;
    });

    destroyChart("status");
    state.charts.status = new Chart(dom.chartStatus, {
      type: "doughnut",
      data: {
        labels: Object.keys(statusCount),
        datasets: [
          {
            data: Object.values(statusCount),
            backgroundColor: ["#1f9d62", "#d28b11", "#c2415a"],
            borderColor: "#ffffff",
            borderWidth: 3,
          },
        ],
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        cutout: "68%",
        plugins: {
          legend: { display: false },
          tooltip: {
            callbacks: {
              label: (context) => `${context.label}: ${context.raw}`,
            },
          },
        },
      },
    });

    const vendorValue = {};
    poEnriched.forEach((record) => {
      if (!record.vendor) return;
      vendorValue[record.vendor] =
        (vendorValue[record.vendor] || 0) + (Number(record.total) || 0);
    });
    const vendorKeys = Object.keys(vendorValue)
      .sort((left, right) => vendorValue[right] - vendorValue[left])
      .slice(0, 10);

    destroyChart("vendor");
    state.charts.vendor = new Chart(dom.chartVendor, {
      type: "bar",
      data: {
        labels: vendorKeys,
        datasets: [
          {
            data: vendorKeys.map((key) => vendorValue[key]),
            backgroundColor: "#19a7a5",
            borderRadius: 10,
            maxBarThickness: 22,
          },
        ],
      },
      options: {
        indexAxis: "y",
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: { display: false },
          tooltip: { callbacks: { label: (context) => fmtIDR(context.raw) } },
        },
        scales: {
          x: {
            beginAtZero: true,
            grid: { color: gridColor },
            ticks: {
              callback: (value) => `${(Number(value) / 1000000).toFixed(1)} jt`,
            },
          },
          y: { grid: { display: false } },
        },
      },
    });
  }

  function buildPoDataset() {
    const detailMap = {};
    state.poDetails.forEach((detail) => {
      if (!detailMap[detail.no]) detailMap[detail.no] = [];
      detailMap[detail.no].push(detail);
    });

    return state.poRows.map((record) => {
      const items = detailMap[record.no] || [];
      const itemSummary = items.length
        ? items
            .map((item) => item.desc)
            .filter(Boolean)
            .join(", ")
        : "-";
      return {
        ...record,
        itemSummary,
        itemSummaryPlain: itemSummary,
        vendorText: record.vendor || "-",
        totalText: Number(record.total) || 0,
      };
    });
  }

  function renderTables() {
    destroyTables();

    const poData = buildPoDataset();
    const lpmData = state.lpmRows.slice();
    state.rendered.po = poData;
    state.rendered.lpm = lpmData;

    if (hasDataTable()) {
      state.tables.po = new DataTable(dom.poTable, {
        data: poData,
        pageLength: 8,
        lengthMenu: [5, 8, 15, 25],
        ordering: true,
        searching: false,
        info: true,
        autoWidth: false,
        language: {
          lengthMenu: "Tampil _MENU_ baris",
          info: "Menampilkan _START_ sampai _END_ dari _TOTAL_ baris",
          infoEmpty: "Tidak ada data",
          emptyTable: "Tidak ada data PO yang terbaca",
          paginate: { previous: "Sebelumnya", next: "Berikutnya" },
        },
        columns: [
          {
            data: "no",
            render: (data, type) =>
              type === "display"
                ? `<span class="strong">${escapeHtml(data)}</span>`
                : data,
          },
          {
            data: "tanggal",
            render: (data, type) =>
              type === "display" ? fmtDateDisplay(data) : data || "",
          },
          {
            data: "vendor",
            render: (data, type) =>
              type === "display" ? escapeHtml(data || "-") : data || "",
          },
          {
            data: "itemSummary",
            render: (data, type) => {
              if (type !== "display") return data || "";
              const text = data && data !== "-" ? data : "-";
              const clipped =
                text.length > 72 ? text.slice(0, 69) + "..." : text;
              return `<span title="${escapeHtml(text)}">${escapeHtml(clipped)}</span>`;
            },
          },
          {
            data: "totalText",
            className: "dt-body-right",
            render: (data, type) => (type === "display" ? fmtIDR(data) : data),
          },
        ],
      });

      state.tables.lpm = new DataTable(dom.lpmTable, {
        data: lpmData,
        pageLength: 8,
        lengthMenu: [5, 8, 15, 25],
        ordering: true,
        searching: false,
        info: true,
        autoWidth: false,
        language: {
          lengthMenu: "Tampil _MENU_ baris",
          info: "Menampilkan _START_ sampai _END_ dari _TOTAL_ baris",
          infoEmpty: "Tidak ada data",
          emptyTable: "Tidak ada data LPM yang terbaca",
          paginate: { previous: "Sebelumnya", next: "Berikutnya" },
        },
        columns: [
          {
            data: "no",
            render: (data, type) =>
              type === "display"
                ? `<span class="strong">${escapeHtml(data)}</span>`
                : data,
          },
          {
            data: "tanggal",
            render: (data, type) =>
              type === "display" ? fmtDateDisplay(data) : data || "",
          },
          {
            data: "dibuat",
            render: (data, type) =>
              type === "display" ? escapeHtml(data || "-") : data || "",
          },
          {
            data: "item",
            render: (data, type) =>
              type === "display" ? escapeHtml(data || "-") : data || "",
          },
          {
            data: "volume",
            className: "dt-body-right",
            render: (data, type, row) => {
              if (type !== "display") return Number(data) || 0;
              const unit = row.satuan ? ` ${escapeHtml(row.satuan)}` : "";
              return `${fmtNumber(data)}${unit}`;
            },
          },
          {
            data: "status",
            render: (data, type) => {
              if (type !== "display") return data || "";
              const label = data || "-";
              return `<span class="badge ${statusClass(label)}">${escapeHtml(label)}</span>`;
            },
          },
        ],
      });
      return;
    }

    applyPlainTableFilters();
  }

  function escapeRegex(value) {
    return String(value).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  }

  function renderDashboard() {
    if (!(state.loaded.lpm && state.loaded.po)) return;
    setStatus("Data siap", true);

    const poData = buildPoDataset();
    populateYearSelector();
    renderCharts(poData);
    renderTables();

    dom.lastUpdated.textContent =
      "Diperbarui: " +
      new Date().toLocaleTimeString("id-ID", { hour12: false });
    dom.uploadNote.textContent =
      "Data berhasil dimuat. Gunakan filter tabel untuk mempercepat monitoring.";
    dom.dashboard.classList.remove("hidden");
  }

  async function loadLpm(file, handle) {
    state.isLoading = true;
    updateLoadedState();
    try {
      const workbook = await readWorkbook(file);
      state.lpmRows = parseLpmWorkbook(workbook);
      state.lpmHandle = handle || null;
      state.fileNames.lpm = file.name;
      state.loaded.lpm = true;
      setFileLabel(dom.uploadFileLpm, file);
      setFileChip(dom.fileChipLpm, file);
      dom.uploaderLpm.classList.add("loaded");
      saveStateToStorage();
      updateLoadedState();
      renderDashboard();
    } catch (error) {
      state.loaded.lpm = false;
      state.lpmRows = [];
      state.lpmHandle = null;
      state.fileNames.lpm = "-";
      setFileLabel(dom.uploadFileLpm, null);
      setFileChip(dom.fileChipLpm, null);
      dom.uploaderLpm.classList.remove("loaded");
      alert("Gagal membaca file LPM: " + error.message);
    } finally {
      state.isLoading = false;
      updateLoadedState();
    }
  }

  async function loadPo(file, handle) {
    state.isLoading = true;
    updateLoadedState();
    try {
      const workbook = await readWorkbook(file);
      const parsed = parsePoWorkbook(workbook);
      state.poRows = parsed.poRows;
      state.poDetails = parsed.poDetails;
      state.poHandle = handle || null;
      state.fileNames.po = file.name;
      state.loaded.po = true;
      setFileLabel(dom.uploadFilePo, file);
      setFileChip(dom.fileChipPo, file);
      dom.uploaderPo.classList.add("loaded");
      saveStateToStorage();
      updateLoadedState();
      renderDashboard();
    } catch (error) {
      state.loaded.po = false;
      state.poRows = [];
      state.poDetails = [];
      state.poHandle = null;
      state.fileNames.po = "-";
      setFileLabel(dom.uploadFilePo, null);
      setFileChip(dom.fileChipPo, null);
      dom.uploaderPo.classList.remove("loaded");
      alert("Gagal membaca file PO: " + error.message);
    } finally {
      state.isLoading = false;
      updateLoadedState();
    }
  }

  function wireDropzone(dropzone, input, handler, pickerOptions) {
    dropzone.addEventListener("click", async () => {
      if (state.isLoading) return;
      if (window.showOpenFilePicker) {
        try {
          const [handle] = await window.showOpenFilePicker(pickerOptions);
          const file = await handle.getFile();
          handler(file, handle);
          return;
        } catch (error) {
          if (error && error.name === "AbortError") return;
        }
      }
      input.click();
    });

    input.addEventListener("change", (event) => {
      if (state.isLoading) return;
      const [file] = event.target.files || [];
      if (file) {
        handler(file, null);
        event.target.value = "";
      }
    });

    dropzone.addEventListener("dragover", (event) => {
      if (state.isLoading) return;
      event.preventDefault();
      dropzone.classList.add("drag");
    });
    dropzone.addEventListener("dragleave", () =>
      dropzone.classList.remove("drag"),
    );
    dropzone.addEventListener("drop", (event) => {
      if (state.isLoading) return;
      event.preventDefault();
      dropzone.classList.remove("drag");
      const [file] = event.dataTransfer.files || [];
      if (file) handler(file, null);
    });
  }

  function clearDashboard() {
    state.loaded = { lpm: false, po: false };
    state.lpmRows = [];
    state.poRows = [];
    state.poDetails = [];
    state.fileNames = { lpm: "-", po: "-" };
    state.lpmHandle = null;
    state.poHandle = null;
    dom.poSearch.value = "";
    dom.poVendorFilter.value = "";
    dom.lpmSearch.value = "";
    dom.lpmStatusFilter.value = "";
    setFileLabel(dom.uploadFileLpm, null);
    setFileLabel(dom.uploadFilePo, null);
    setFileChip(dom.fileChipLpm, null);
    setFileChip(dom.fileChipPo, null);
    dom.uploaderLpm.classList.remove("loaded");
    dom.uploaderPo.classList.remove("loaded");
    // destroyCharts();
    destroyTables();
    dom.kpiLpm.textContent = "0";
    dom.kpiPo.textContent = "0";
    dom.kpiVendor.textContent = "0";
    setKpiValue(0);
    dom.kpiValueNote.textContent = "Belum ada nilai PO yang terbaca";
    dom.uploadNote.textContent =
      "Unggah file LPM dan PO. Data diproses hanya di browser Anda.";
    dom.dashboard.classList.add("hidden");
    dom.controlBar.classList.add("hidden");
    clearStateStorage();
    updateLoadedState();
  }

  async function reloadFiles() {
    try {
      if (state.lpmHandle) {
        const permission = await state.lpmHandle.queryPermission({
          mode: "read",
        });
        if (permission !== "granted")
          await state.lpmHandle.requestPermission({ mode: "read" });
        const file = await state.lpmHandle.getFile();
        await loadLpm(file, state.lpmHandle);
      }
      if (state.poHandle) {
        const permission = await state.poHandle.queryPermission({
          mode: "read",
        });
        if (permission !== "granted")
          await state.poHandle.requestPermission({ mode: "read" });
        const file = await state.poHandle.getFile();
        await loadPo(file, state.poHandle);
      }
      if (!state.lpmHandle && !state.poHandle) {
        alert(
          "Browser ini tidak menyimpan handle file. Gunakan Ganti File untuk memilih ulang dokumen.",
        );
      }
    } catch (error) {
      alert("Gagal memuat ulang: " + error.message);
    }
  }

  function initDom() {
    dom.body = document.body;
    dom.uploadNote = $("uploadNote");
    dom.dashboard = $("dashboard");
    dom.controlBar = $("controlBar");
    dom.lastUpdated = $("lastUpdated");
    dom.controlFileLpm = $("controlFileLpm");
    dom.controlFilePo = $("controlFilePo");
    dom.statusFileLpm = $("statusFileLpm");
    dom.statusFilePo = $("statusFilePo");
    dom.uploadFileLpm = $("dzLpmName");
    dom.uploadFilePo = $("dzPoName");
    dom.fileChipLpm = $("dzLpmChip");
    dom.fileChipPo = $("dzPoChip");
    dom.statusDot = $("statusDot");
    dom.statusText = $("statusText");
    dom.statusMessage = $("statusMessage");
    dom.kpiLpm = $("kpiLpm");
    dom.kpiPo = $("kpiPo");
    dom.kpiVendor = $("kpiVendor");
    dom.kpiValue = $("kpiValue");
    dom.kpiValueNote = $("kpiValueNote");
    dom.chartLpm = $("chartLpm");
    dom.chartPo = $("chartPo");
    dom.chartStatus = $("chartStatus");
    dom.chartVendor = $("chartVendor");
    dom.poTable = $("poTable");
    dom.lpmTable = $("lpmTable");
    dom.poSearch = $("poSearch");
    dom.poVendorFilter = $("poVendorFilter");
    dom.lpmSearch = $("lpmSearch");
    dom.lpmStatusFilter = $("lpmStatusFilter");
    dom.uploaderLpm = $("uploaderLpm");
    dom.uploaderPo = $("uploaderPo");
    dom.uploadGrid = $("uploadGrid");
    dom.resetBtn = $("resetBtn");
    dom.reloadBtn = $("reloadBtn");
    dom.changeBtn = $("changeBtn");
    dom.yearSelect = $("yearSelect");
  }

  function bindActions() {
    wireDropzone(dom.uploaderLpm, $("inputLpm"), loadLpm, {
      types: [
        {
          description: "Excel",
          accept: {
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet":
              [".xlsx"],
            "application/vnd.ms-excel": [".xls"],
          },
        },
      ],
    });

    wireDropzone(dom.uploaderPo, $("inputPo"), loadPo, {
      types: [
        {
          description: "Excel",
          accept: {
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet":
              [".xlsx"],
            "application/vnd.ms-excel": [".xls"],
          },
        },
      ],
    });

    dom.reloadBtn.addEventListener("click", reloadFiles);
    dom.changeBtn.addEventListener("click", clearDashboard);
    dom.resetBtn.addEventListener("click", clearDashboard);
    const openBtn = $("openSideMenuBtn");
    const closeBtn = $("closeSideMenuBtn");
    const sidePanel = $("sidePanel");
    if (openBtn && sidePanel) {
      openBtn.addEventListener("click", () => {
        sidePanel.classList.remove("hidden");
        document.body.classList.add("no-scroll");
      });
    }
    if (closeBtn && sidePanel) {
      closeBtn.addEventListener("click", () => {
        sidePanel.classList.add("hidden");
        document.body.classList.remove("no-scroll");
      });
    }
    const toggleBtn = $("toggleUploadBtn");
    if (toggleBtn && dom.uploadGrid) {
      toggleBtn.addEventListener("click", () => {
        const hidden = dom.uploadGrid.classList.toggle("hidden");
        toggleBtn.textContent = hidden
          ? "Tampilkan Upload"
          : "Sembunyikan Upload";
      });
    }

    // Wire visible 'Pilih file' buttons to file inputs
    Array.from(document.querySelectorAll(".uploader")).forEach((u) => {
      const pick = u.querySelector(".uploader-actions .btn.primary");
      const input = u.querySelector("input[type=file]");
      if (pick && input) {
        pick.addEventListener("click", (e) => {
          e.stopPropagation();
          if (
            state.isLoading ||
            (dom.uploadGrid && dom.uploadGrid.classList.contains("hidden"))
          )
            return;
          input.click();
        });
      }
    });

    dom.yearSelect.addEventListener("change", (e) => {
      state.selectedYear = e.target.value;
      if (state.loaded.lpm && state.loaded.po) {
        renderCharts(buildPoDataset());
        dom.lastUpdated.textContent =
          "Diperbarui: " +
          new Date().toLocaleTimeString("id-ID", { hour12: false });
      }
    });

    dom.poSearch.addEventListener("input", () => {
      if (hasDataTable() && state.tables.po) {
        state.tables.po.search(dom.poSearch.value.trim()).draw();
        return;
      }
      applyPlainTableFilters();
    });

    dom.poVendorFilter.addEventListener("change", () => {
      if (hasDataTable() && state.tables.po) {
        const value = dom.poVendorFilter.value;
        state.tables.po
          .column(2)
          .search(value ? `^${escapeRegex(value)}$` : "", true, false)
          .draw();
        return;
      }
      applyPlainTableFilters();
    });

    dom.lpmSearch.addEventListener("input", () => {
      if (hasDataTable() && state.tables.lpm) {
        state.tables.lpm.search(dom.lpmSearch.value.trim()).draw();
        return;
      }
      applyPlainTableFilters();
    });

    dom.lpmStatusFilter.addEventListener("change", () => {
      if (hasDataTable() && state.tables.lpm) {
        const value = dom.lpmStatusFilter.value;
        state.tables.lpm
          .column(5)
          .search(value ? `^${escapeRegex(value)}$` : "", true, false)
          .draw();
        return;
      }
      applyPlainTableFilters();
    });

    // Chart fullscreen toggle buttons (icon-only)
    const _fsButtons = Array.from(
      document.querySelectorAll(".chart-fullscreen-btn"),
    );
    _fsButtons.forEach((btn) => setBtnToExpand(btn));

    _fsButtons.forEach((btn) => {
      btn.addEventListener("click", (e) => {
        e.preventDefault();
        const targetId = btn.getAttribute("data-target");
        if (!targetId) return;
        const canvas = document.getElementById(targetId);
        if (!canvas) return;
        const panel = canvas.closest(".panel");
        if (!panel) return;

        const entering = !panel.classList.contains("fullscreen");
        // if entering fullscreen, remove other fullscreen panels
        if (entering) {
          document.querySelectorAll(".panel.fullscreen").forEach((p) => {
            if (p !== panel) p.classList.remove("fullscreen");
          });
        }

        panel.classList.toggle("fullscreen");
        document.body.classList.toggle("no-scroll", entering);

        // also toggle chart-box class for height adjustments
        const chartBox = canvas.closest(".chart-box");
        if (chartBox) chartBox.classList.toggle("fullscreen");

        // update icons for all buttons
        _fsButtons.forEach((b) => {
          const t = b.getAttribute("data-target");
          const c = document.getElementById(t);
          const p = c && c.closest(".panel");
          if (!p) return;
          if (p.classList.contains("fullscreen")) setBtnToCollapse(b);
          else setBtnToExpand(b);
        });

        // resize chart(s)
        setTimeout(() => {
          try {
            const chartObj = Chart.getChart(canvas);
            if (chartObj && typeof chartObj.resize === "function") {
              chartObj.resize();
            }
          } catch (err) {
            // ignore
          }
        }, 60);
      });
    });

    // Allow Esc to exit fullscreen
    document.addEventListener("keydown", (ev) => {
      if (ev.key === "Escape") {
        const panels = document.querySelectorAll(".panel.fullscreen");
        if (!panels.length) return;
        panels.forEach((p) => {
          p.classList.remove("fullscreen");
          const canvas = p.querySelector("canvas");
          if (canvas) {
            const chartBox = canvas.closest(".chart-box");
            if (chartBox) chartBox.classList.remove("fullscreen");
            try {
              const chartObj = Chart.getChart(canvas);
              if (chartObj && typeof chartObj.resize === "function")
                chartObj.resize();
            } catch (e) {}
          }
        });
        document.body.classList.remove("no-scroll");
        _fsButtons.forEach((b) => setBtnToExpand(b));
      }
    });
  }

  function init() {
    initDom();
    bindActions();
    const restored = restoreStateFromStorage();
    updateLoadedState();
    if (restored && state.loaded.lpm && state.loaded.po) {
      renderDashboard();
    }
  }

  document.addEventListener("DOMContentLoaded", init);
})();
