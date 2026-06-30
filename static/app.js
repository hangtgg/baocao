const APP_CONFIG = window.APP_CONFIG || {};
const ENTRY_PAGE_SIZE = 10;

const state = {
    filters: { ...(APP_CONFIG.defaultFilters || {}) },
    contentSuggestions: {},
    services: [],
    supporters: [],
    customers: [],
    entries: [],
    entryPage: 1,
    editingEntryId: null
};

const dom = {};

document.addEventListener("DOMContentLoaded", async () => {
    captureDom();
    bindEvents();
    initializeDefaults();
    await loadData({ resetPage: true });
});

function captureDom() {
    dom.filtersForm = document.getElementById("filtersForm");
    dom.weekPicker = document.getElementById("weekPicker");
    dom.monthPicker = document.getElementById("monthPicker");
    dom.dateFrom = document.getElementById("dateFrom");
    dom.dateTo = document.getElementById("dateTo");
    dom.serviceFilter = document.getElementById("serviceFilter");
    dom.statusFilter = document.getElementById("statusFilter");
    dom.supporterFilter = document.getElementById("supporterFilter");
    dom.resetFiltersButton = document.getElementById("resetFiltersButton");
    dom.presetWeekButton = document.getElementById("presetWeekButton");
    dom.presetMonthButton = document.getElementById("presetMonthButton");
    dom.presetAllButton = document.getElementById("presetAllButton");
    dom.exportExcelButton = document.getElementById("exportExcelButton");

    dom.entryForm = document.getElementById("entryForm");
    dom.entryId = document.getElementById("entryId");
    dom.supportDate = document.getElementById("supportDate");
    dom.customerName = document.getElementById("customerName");
    dom.serviceId = document.getElementById("serviceId");
    dom.channel = document.getElementById("channel");
    dom.supporterSelect = document.getElementById("supporterSelect");
    dom.supporterOtherField = document.getElementById("supporterOtherField");
    dom.supporterOtherName = document.getElementById("supporterOtherName");
    dom.requesterName = document.getElementById("requesterName");
    dom.status = document.getElementById("status");
    dom.supportContent = document.getElementById("supportContent");
    dom.supportContentSuggestion = document.getElementById("supportContentSuggestion");
    dom.applySuggestionButton = document.getElementById("applySuggestionButton");
    dom.supportSuggestionPreview = document.getElementById("supportSuggestionPreview");
    dom.notes = document.getElementById("notes");
    dom.cancelEntryEditButton = document.getElementById("cancelEntryEditButton");
    dom.entryFormBadge = document.getElementById("entryFormBadge");
    dom.entryFormMessage = document.getElementById("entryFormMessage");
    dom.customerSuggestions = document.getElementById("customerSuggestions");
    dom.requesterSuggestions = document.getElementById("requesterSuggestions");

    dom.entriesSummary = document.getElementById("entriesSummary");
    dom.entriesTableBody = document.getElementById("entriesTableBody");
    dom.entriesPagination = document.getElementById("entriesPagination");
    dom.toastContainer = document.getElementById("toastContainer");
}

function bindEvents() {
    dom.filtersForm.addEventListener("submit", async (event) => {
        event.preventDefault();
        syncPresetPickersWithRange(dom.dateFrom.value, dom.dateTo.value);
        await loadData({ resetPage: true });
    });

    dom.weekPicker.addEventListener("change", () => {
        if (!dom.weekPicker.value) {
            return;
        }
        const [dateFrom, dateTo] = rangeFromWeekValue(dom.weekPicker.value);
        dom.dateFrom.value = dateFrom;
        dom.dateTo.value = dateTo;
        dom.monthPicker.value = "";
    });

    dom.monthPicker.addEventListener("change", () => {
        if (!dom.monthPicker.value) {
            return;
        }
        const [dateFrom, dateTo] = rangeFromMonthValue(dom.monthPicker.value);
        dom.dateFrom.value = dateFrom;
        dom.dateTo.value = dateTo;
        dom.weekPicker.value = "";
    });

    dom.resetFiltersButton.addEventListener("click", async () => {
        applyCurrentWeekPreset();
        dom.serviceFilter.value = "";
        dom.statusFilter.value = "";
        dom.supporterFilter.value = "";
        await loadData({ resetPage: true });
    });

    dom.presetWeekButton.addEventListener("click", async () => {
        applyCurrentWeekPreset();
        await loadData({ resetPage: true });
    });

    dom.presetMonthButton.addEventListener("click", async () => {
        applyCurrentMonthPreset();
        await loadData({ resetPage: true });
    });

    dom.presetAllButton.addEventListener("click", async () => {
        dom.weekPicker.value = "";
        dom.monthPicker.value = "";
        dom.dateFrom.value = "";
        dom.dateTo.value = "";
        dom.serviceFilter.value = "";
        dom.statusFilter.value = "";
        dom.supporterFilter.value = "";
        await loadData({ resetPage: true });
    });

    dom.exportExcelButton.addEventListener("click", () => {
        window.location.href = buildExportUrl("/export/report.xlsx");
    });

    dom.entryForm.addEventListener("submit", handleEntrySubmit);
    dom.serviceId.addEventListener("change", () => {
        renderSupportSuggestions();
    });
    dom.supporterSelect.addEventListener("change", syncSupporterOtherField);
    dom.supportContentSuggestion.addEventListener("change", () => {
        updateSuggestionPreview();
    });
    dom.applySuggestionButton.addEventListener("click", applySelectedSuggestion);
    dom.cancelEntryEditButton.addEventListener("click", resetEntryForm);
    dom.entriesTableBody.addEventListener("click", handleEntryTableClick);
    dom.entriesPagination.addEventListener("click", handleEntryPaginationClick);
}

function initializeDefaults() {
    if (!dom.supportDate.value) {
        dom.supportDate.value = formatDateForInput(new Date());
    }
    dom.channel.value = APP_CONFIG.channels?.[0] || "Group ĐTP-AM-Hỗ trợ";
    dom.status.value = APP_CONFIG.statuses?.[0] || "Đã xử lý";
    syncSupporterOtherField();
    syncPresetPickersWithRange(dom.dateFrom.value, dom.dateTo.value);
}

async function loadData(options = {}) {
    const { resetPage = false } = options;
    const query = buildQueryString(readFilters());
    dom.entriesSummary.textContent = "Đang tải dữ liệu...";

    try {
        const data = await fetchJson(`/api/bootstrap?${query}`);
        state.filters = data.filters;
        state.contentSuggestions = data.content_suggestions || {};
        state.services = data.services || [];
        state.supporters = data.supporters || [];
        state.customers = data.customers || [];
        state.entries = data.entries || [];
        if (resetPage) {
            state.entryPage = 1;
        }

        renderFilters();
        renderEntryFormOptions();
        renderSuggestions();
        renderEntriesTable();
        clearFormMessage();
    } catch (error) {
        dom.entriesSummary.textContent = "Không tải được dữ liệu.";
        showToast(error.message || "Không thể tải dữ liệu.", "error");
    }
}

function renderFilters() {
    fillSelect(
        dom.serviceFilter,
        [{ value: "", label: "Tất cả dịch vụ" }].concat(
            state.services.map((service) => ({
                value: String(service.id),
                label: service.is_active ? service.name : `${service.name} (ngừng dùng)`
            }))
        ),
        state.filters.service_id ? String(state.filters.service_id) : ""
    );

    fillSelect(
        dom.supporterFilter,
        [{ value: "", label: "Tất cả người hỗ trợ" }].concat(
            state.supporters.map((name) => ({ value: name, label: name }))
        ),
        state.filters.supporter_name || ""
    );

    dom.statusFilter.value = state.filters.status || "";
    dom.dateFrom.value = state.filters.date_from || "";
    dom.dateTo.value = state.filters.date_to || "";
    syncPresetPickersWithRange(dom.dateFrom.value, dom.dateTo.value);
}

function renderEntryFormOptions() {
    const currentValue = dom.serviceId.value || "";
    const availableServices = state.services.filter(
        (service) => service.is_active || String(service.id) === String(currentValue)
    );

    fillSelect(
        dom.serviceId,
        [{ value: "", label: "Chọn dịch vụ" }].concat(
            availableServices.map((service) => ({
                value: String(service.id),
                label: service.name
            }))
        ),
        currentValue
    );
    renderSupportSuggestions();
}

function renderSuggestions() {
    dom.customerSuggestions.innerHTML = state.customers
        .map((customer) => `<option value="${escapeHtml(customer)}"></option>`)
        .join("");

    dom.requesterSuggestions.innerHTML = (APP_CONFIG.requesterOptions || [])
        .map(
            (requester) => `
                <option
                    value="${escapeHtml(requester.name)}"
                    label="${escapeHtml(`${requester.name} - ${requester.area}`)}"
                ></option>
            `
        )
        .join("");
}

function renderEntriesTable() {
    const totalEntries = state.entries.length;
    const totalPages = Math.max(1, Math.ceil(totalEntries / ENTRY_PAGE_SIZE));
    state.entryPage = Math.min(Math.max(state.entryPage, 1), totalPages);

    if (!totalEntries) {
        dom.entriesSummary.textContent = `Hiển thị 0 lượt hỗ trợ ${buildPeriodText()}.`;
        dom.entriesTableBody.innerHTML = `
            <tr>
                <td colspan="10">
                    <div class="empty-state">Chưa có lượt hỗ trợ nào trong bộ lọc hiện tại.</div>
                </td>
            </tr>
        `;
        renderEntryPagination(0, 1, 1);
        return;
    }

    const startIndex = (state.entryPage - 1) * ENTRY_PAGE_SIZE;
    const pagedEntries = state.entries.slice(startIndex, startIndex + ENTRY_PAGE_SIZE);
    const endIndex = startIndex + pagedEntries.length;
    dom.entriesSummary.textContent = `Hiển thị ${startIndex + 1}-${endIndex}/${totalEntries} lượt hỗ trợ ${buildPeriodText()}.`;

    dom.entriesTableBody.innerHTML = pagedEntries
        .map(
            (entry) => `
                <tr>
                    <td>
                        <div class="table-actions">
                            <button type="button" class="link-button" data-action="edit-entry" data-id="${entry.id}">Sửa</button>
                            <button type="button" class="link-button link-button--danger" data-action="delete-entry" data-id="${entry.id}">Xóa</button>
                        </div>
                    </td>
                    <td>${escapeHtml(entry.supporter_name)}</td>
                    <td>${renderOptionalCell(entry.customer_name)}</td>
                    <td>${renderOptionalCell(entry.requester_name)}</td>
                    <td>${escapeHtml(entry.service_name)}</td>
                    <td><div class="cell-content">${renderOptionalCell(entry.support_content)}</div></td>
                    <td>${escapeHtml(entry.channel)}</td>
                    <td>${renderStatusBadge(entry.status)}</td>
                    <td><div class="cell-note">${renderOptionalCell(entry.notes)}</div></td>
                    <td>${escapeHtml(formatDisplayDate(entry.support_date))}</td>
                </tr>
            `
        )
        .join("");

    renderEntryPagination(totalEntries, totalPages, state.entryPage);
}

function renderEntryPagination(totalEntries, totalPages, currentPage) {
    if (!dom.entriesPagination) {
        return;
    }

    if (!totalEntries || totalPages <= 1) {
        dom.entriesPagination.innerHTML = "";
        dom.entriesPagination.classList.add("hidden");
        return;
    }

    const pageButtons = buildEntryPageNumbers(totalPages, currentPage)
        .map((page) => {
            const isCurrent = page === currentPage;
            return `
                <button
                    type="button"
                    class="pagination__button ${isCurrent ? "pagination__button--active" : ""}"
                    data-page="${page}"
                    ${isCurrent ? 'aria-current="page"' : ""}
                >
                    ${page}
                </button>
            `;
        })
        .join("");

    dom.entriesPagination.innerHTML = `
        <div class="pagination__summary">Trang ${currentPage}/${totalPages}</div>
        <div class="pagination__controls">
            <button
                type="button"
                class="pagination__button"
                data-page="${currentPage - 1}"
                ${currentPage === 1 ? "disabled" : ""}
            >
                Trước
            </button>
            ${pageButtons}
            <button
                type="button"
                class="pagination__button"
                data-page="${currentPage + 1}"
                ${currentPage === totalPages ? "disabled" : ""}
            >
                Sau
            </button>
        </div>
    `;
    dom.entriesPagination.classList.remove("hidden");
}

function buildEntryPageNumbers(totalPages, currentPage) {
    const startPage = Math.max(1, currentPage - 2);
    const endPage = Math.min(totalPages, startPage + 4);
    const normalizedStartPage = Math.max(1, endPage - 4);
    const pages = [];

    for (let page = normalizedStartPage; page <= endPage; page += 1) {
        pages.push(page);
    }
    return pages;
}

function handleEntryPaginationClick(event) {
    const button = event.target.closest("button[data-page]");
    if (!button || button.disabled) {
        return;
    }

    const nextPage = Number(button.dataset.page);
    if (!Number.isInteger(nextPage) || nextPage < 1 || nextPage === state.entryPage) {
        return;
    }

    state.entryPage = nextPage;
    renderEntriesTable();
    dom.entriesPagination.scrollIntoView({ behavior: "smooth", block: "nearest" });
}

async function handleEntrySubmit(event) {
    event.preventDefault();
    const payload = {
        support_date: dom.supportDate.value,
        customer_name: dom.customerName.value,
        requester_name: dom.requesterName.value.trim(),
        service_id: dom.serviceId.value,
        support_content: dom.supportContent.value,
        channel: dom.channel.value,
        supporter_name: getSelectedSupporterName(),
        status: dom.status.value,
        notes: dom.notes.value
    };

    const isEditing = Boolean(state.editingEntryId);
    const url = isEditing ? `/api/entries/${state.editingEntryId}` : "/api/entries";
    const method = isEditing ? "PUT" : "POST";

    try {
        await fetchJson(url, {
            method,
            body: JSON.stringify(payload)
        });
        showToast(isEditing ? "Đã cập nhật lượt hỗ trợ." : "Đã thêm lượt hỗ trợ.");
        resetEntryForm();
        await loadData({ resetPage: !isEditing });
    } catch (error) {
        setFormMessage(error.message, true);
        showToast(error.message || "Không lưu được lượt hỗ trợ.", "error");
    }
}

function handleEntryTableClick(event) {
    const button = event.target.closest("button[data-action]");
    if (!button) {
        return;
    }

    const action = button.dataset.action;
    const entryId = Number(button.dataset.id);

    if (action === "edit-entry") {
        const entry = state.entries.find((item) => item.id === entryId);
        if (entry) {
            startEntryEdit(entry);
        }
        return;
    }

    if (action === "delete-entry") {
        deleteEntry(entryId);
    }
}

function startEntryEdit(entry) {
    state.editingEntryId = entry.id;
    dom.entryId.value = String(entry.id);
    dom.supportDate.value = entry.support_date;
    dom.customerName.value = entry.customer_name;
    dom.serviceId.value = String(entry.service_id);
    dom.channel.value = normalizeChannelSelection(entry.channel);
    setSupporterSelection(entry.supporter_name);
    dom.requesterName.value = entry.requester_name || "";
    dom.status.value = entry.status;
    dom.supportContent.value = entry.support_content;
    dom.notes.value = entry.notes || "";
    renderSupportSuggestions(entry.support_content);
    dom.entryFormBadge.textContent = "Chỉnh sửa";
    dom.cancelEntryEditButton.classList.remove("hidden");
    setFormMessage(
        `Đang chỉnh sửa lượt hỗ trợ cho ${entry.customer_name || entry.service_name}.`,
        false
    );
    dom.entryForm.scrollIntoView({ behavior: "smooth", block: "start" });
}

function resetEntryForm() {
    state.editingEntryId = null;
    dom.entryId.value = "";
    dom.entryForm.reset();
    dom.supportDate.value = formatDateForInput(new Date());
    dom.channel.value = APP_CONFIG.channels?.[0] || "Group ĐTP-AM-Hỗ trợ";
    dom.status.value = APP_CONFIG.statuses?.[0] || "Đã xử lý";
    dom.supporterSelect.value = "";
    dom.supporterOtherName.value = "";
    dom.requesterName.value = "";
    dom.entryFormBadge.textContent = "Thêm mới";
    dom.cancelEntryEditButton.classList.add("hidden");
    syncSupporterOtherField();
    renderEntryFormOptions();
    renderSupportSuggestions();
    clearFormMessage();
}

function normalizeChannelSelection(channel) {
    const rawChannel = String(channel || "").trim();
    const availableChannels = APP_CONFIG.channels || [];
    if (availableChannels.includes(rawChannel)) {
        return rawChannel;
    }

    const legacyMap = {
        "Ultra khách": "Group ĐTP-AM-Hỗ trợ",
        "Zalo": "Nhắn tin riêng",
        "Trực tiếp": "Khác"
    };
    return legacyMap[rawChannel] || availableChannels[0] || "";
}

function getSelectedSupporterName() {
    if (dom.supporterSelect.value === "__other__") {
        return dom.supporterOtherName.value.trim();
    }
    return dom.supporterSelect.value.trim();
}

function isPredefinedSupporter(name) {
    return (APP_CONFIG.supporterOptions || []).some(
        (supporter) => supporter.name === name
    );
}

function setSupporterSelection(name) {
    const supporterName = String(name || "").trim();
    if (supporterName && isPredefinedSupporter(supporterName)) {
        dom.supporterSelect.value = supporterName;
        dom.supporterOtherName.value = "";
    } else if (supporterName) {
        dom.supporterSelect.value = "__other__";
        dom.supporterOtherName.value = supporterName;
    } else {
        dom.supporterSelect.value = "";
        dom.supporterOtherName.value = "";
    }
    syncSupporterOtherField();
}

function syncSupporterOtherField() {
    const isOtherSelected = dom.supporterSelect.value === "__other__";
    dom.supporterOtherField.classList.toggle("hidden", !isOtherSelected);
    dom.supporterOtherName.required = isOtherSelected;
    if (!isOtherSelected) {
        dom.supporterOtherName.value = "";
    }
}

function renderSupportSuggestions(preferredSuggestion = "") {
    const selectedServiceId = String(dom.serviceId.value || "");
    const suggestions = state.contentSuggestions[selectedServiceId] || [];
    const selectedValue = suggestions.includes(preferredSuggestion) ? preferredSuggestion : "";
    const placeholderLabel = suggestions.length
        ? "Chọn nội dung gợi ý theo dịch vụ"
        : "Dịch vụ này chưa có gợi ý riêng";

    fillSelect(
        dom.supportContentSuggestion,
        [{ value: "", label: placeholderLabel }].concat(
            suggestions.map((suggestion, index) => ({
                value: suggestion,
                label: `Gợi ý ${index + 1}: ${truncateText(suggestion, 110)}`
            }))
        ),
        selectedValue
    );

    dom.supportContent.placeholder = suggestions.length
        ? `Gợi ý nhanh: ${truncateText(suggestions[0], 120)}`
        : "Có thể để trống hoặc nhập khi cần...";

    updateSuggestionPreview();
}

function updateSuggestionPreview() {
    const selectedSuggestion = dom.supportContentSuggestion.value.trim();
    const selectedServiceId = String(dom.serviceId.value || "");
    const suggestions = state.contentSuggestions[selectedServiceId] || [];

    if (selectedSuggestion) {
        dom.supportSuggestionPreview.textContent = selectedSuggestion;
        dom.applySuggestionButton.disabled = false;
        return;
    }

    dom.applySuggestionButton.disabled = true;
    dom.supportSuggestionPreview.textContent = suggestions.length
        ? "Chọn một gợi ý phía trên rồi bấm “Chèn gợi ý” để đưa vào ô Nội dung hỗ trợ."
        : "Chọn dịch vụ hỗ trợ để xem các nội dung gợi ý phù hợp.";
}

function applySelectedSuggestion() {
    const selectedSuggestion = dom.supportContentSuggestion.value.trim();
    if (!selectedSuggestion) {
        showToast("Hãy chọn một nội dung gợi ý trước.", "error");
        return;
    }

    const currentContent = dom.supportContent.value.trim();
    if (!currentContent) {
        dom.supportContent.value = selectedSuggestion;
    } else if (!currentContent.includes(selectedSuggestion)) {
        dom.supportContent.value = `${currentContent}\n${selectedSuggestion}`;
    }

    dom.supportContent.focus();
    showToast("Đã chèn nội dung gợi ý.");
}

async function deleteEntry(entryId) {
    if (!window.confirm("Bạn có chắc muốn xóa lượt hỗ trợ này không?")) {
        return;
    }

    try {
        await fetchJson(`/api/entries/${entryId}`, { method: "DELETE" });
        showToast("Đã xóa lượt hỗ trợ.");
        if (state.editingEntryId === entryId) {
            resetEntryForm();
        }
        await loadData();
    } catch (error) {
        showToast(error.message || "Không xóa được lượt hỗ trợ.", "error");
    }
}

function readFilters() {
    return {
        date_from: dom.dateFrom.value,
        date_to: dom.dateTo.value,
        service_id: dom.serviceFilter.value,
        status: dom.statusFilter.value,
        supporter_name: dom.supporterFilter.value
    };
}

function buildQueryString(filters) {
    const params = new URLSearchParams();
    Object.entries(filters).forEach(([key, value]) => {
        if (value !== undefined && value !== null && String(value).trim() !== "") {
            params.set(key, value);
        }
    });
    return params.toString();
}

function buildExportUrl(path) {
    return `${path}?${buildQueryString(readFilters())}`;
}

function applyCurrentWeekPreset() {
    const now = new Date();
    const day = now.getDay() || 7;
    const monday = new Date(now);
    monday.setDate(now.getDate() - day + 1);
    const sunday = new Date(monday);
    sunday.setDate(monday.getDate() + 6);
    dom.dateFrom.value = formatDateForInput(monday);
    dom.dateTo.value = formatDateForInput(sunday);
    dom.weekPicker.value = formatWeekInput(monday);
    dom.monthPicker.value = "";
}

function applyCurrentMonthPreset() {
    const now = new Date();
    const firstDay = new Date(now.getFullYear(), now.getMonth(), 1);
    const lastDay = new Date(now.getFullYear(), now.getMonth() + 1, 0);
    dom.dateFrom.value = formatDateForInput(firstDay);
    dom.dateTo.value = formatDateForInput(lastDay);
    dom.monthPicker.value = `${firstDay.getFullYear()}-${String(firstDay.getMonth() + 1).padStart(2, "0")}`;
    dom.weekPicker.value = "";
}

function syncPresetPickersWithRange(dateFrom, dateTo) {
    if (!dateFrom || !dateTo) {
        dom.weekPicker.value = "";
        dom.monthPicker.value = "";
        return;
    }

    if (isFullWeek(dateFrom, dateTo)) {
        dom.weekPicker.value = formatWeekInput(new Date(`${dateFrom}T00:00:00`));
    } else {
        dom.weekPicker.value = "";
    }

    if (isFullMonth(dateFrom, dateTo)) {
        dom.monthPicker.value = dateFrom.slice(0, 7);
    } else {
        dom.monthPicker.value = "";
    }
}

function isFullWeek(dateFrom, dateTo) {
    const start = new Date(`${dateFrom}T00:00:00`);
    const end = new Date(`${dateTo}T00:00:00`);
    const diffDays = Math.round((end - start) / 86400000);
    const startDay = start.getDay() || 7;
    return startDay === 1 && diffDays === 6;
}

function isFullMonth(dateFrom, dateTo) {
    const start = new Date(`${dateFrom}T00:00:00`);
    const end = new Date(`${dateTo}T00:00:00`);
    const lastDay = new Date(start.getFullYear(), start.getMonth() + 1, 0).getDate();
    return start.getDate() === 1
        && end.getDate() === lastDay
        && start.getMonth() === end.getMonth()
        && start.getFullYear() === end.getFullYear();
}

function rangeFromWeekValue(weekValue) {
    const [yearPart, weekPart] = weekValue.split("-W");
    const year = Number(yearPart);
    const week = Number(weekPart);
    const januaryFourth = new Date(Date.UTC(year, 0, 4));
    const dayOfWeek = januaryFourth.getUTCDay() || 7;
    const monday = new Date(januaryFourth);
    monday.setUTCDate(januaryFourth.getUTCDate() - dayOfWeek + 1 + (week - 1) * 7);
    const sunday = new Date(monday);
    sunday.setUTCDate(monday.getUTCDate() + 6);
    return [formatDateForInput(monday, true), formatDateForInput(sunday, true)];
}

function rangeFromMonthValue(monthValue) {
    const [yearPart, monthPart] = monthValue.split("-");
    const year = Number(yearPart);
    const month = Number(monthPart);
    const firstDay = new Date(year, month - 1, 1);
    const lastDay = new Date(year, month, 0);
    return [formatDateForInput(firstDay), formatDateForInput(lastDay)];
}

function formatWeekInput(value) {
    const date = value instanceof Date ? value : new Date(value);
    const target = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
    const dayNumber = target.getUTCDay() || 7;
    target.setUTCDate(target.getUTCDate() + 4 - dayNumber);
    const yearStart = new Date(Date.UTC(target.getUTCFullYear(), 0, 1));
    const weekNumber = Math.ceil((((target - yearStart) / 86400000) + 1) / 7);
    return `${target.getUTCFullYear()}-W${String(weekNumber).padStart(2, "0")}`;
}

function formatDisplayDate(value) {
    if (!value) {
        return "";
    }
    const [year, month, day] = value.split("-");
    return `${day}/${month}/${year}`;
}

function formatDateForInput(value, utc = false) {
    const year = utc ? value.getUTCFullYear() : value.getFullYear();
    const month = utc ? value.getUTCMonth() + 1 : value.getMonth() + 1;
    const day = utc ? value.getUTCDate() : value.getDate();
    return `${year}-${String(month).padStart(2, "0")}-${String(day).padStart(2, "0")}`;
}

function fillSelect(selectElement, options, selectedValue) {
    selectElement.innerHTML = options
        .map((option) => `<option value="${escapeHtml(option.value)}">${escapeHtml(option.label)}</option>`)
        .join("");
    selectElement.value = selectedValue || "";
}

function renderStatusBadge(status) {
    const isDone = status === "Đã xử lý";
    return `<span class="badge ${isDone ? "" : "badge--inactive"}">${escapeHtml(status)}</span>`;
}

function renderOptionalCell(value) {
    if (String(value ?? "").trim()) {
        return escapeHtml(value);
    }
    return '<span class="cell-empty">Để trống</span>';
}

function buildPeriodText() {
    const dateFrom = state.filters.date_from;
    const dateTo = state.filters.date_to;
    if (dateFrom && dateTo) {
        return `trong khoảng ${formatDisplayDate(dateFrom)} - ${formatDisplayDate(dateTo)}`;
    }
    if (dateFrom) {
        return `từ ${formatDisplayDate(dateFrom)}`;
    }
    if (dateTo) {
        return `đến ${formatDisplayDate(dateTo)}`;
    }
    return "trên toàn bộ dữ liệu";
}

function setFormMessage(message, isError) {
    dom.entryFormMessage.textContent = message || "";
    dom.entryFormMessage.style.color = isError ? "var(--danger)" : "var(--muted)";
}

function clearFormMessage() {
    setFormMessage("", false);
}

function showToast(message, type = "success") {
    const toast = document.createElement("div");
    toast.className = `toast ${type === "error" ? "toast--error" : ""}`;
    toast.textContent = message;
    dom.toastContainer.appendChild(toast);
    window.setTimeout(() => {
        toast.remove();
    }, 3200);
}

async function fetchJson(url, options = {}) {
    const response = await fetch(url, {
        headers: {
            Accept: "application/json",
            "Content-Type": "application/json"
        },
        ...options
    });

    const contentType = response.headers.get("content-type") || "";
    const isJson = contentType.includes("application/json");
    const data = isJson ? await response.json() : null;

    if (!response.ok) {
        throw new Error(extractErrorMessage(data) || "Yêu cầu không thành công.");
    }

    return data;
}

function extractErrorMessage(data) {
    if (!data) {
        return "";
    }
    if (data.error) {
        return data.error;
    }
    if (data.message) {
        return data.message;
    }
    if (data.errors) {
        return Object.values(data.errors).join(" ");
    }
    return "";
}

function escapeHtml(value) {
    return String(value ?? "")
        .replaceAll("&", "&amp;")
        .replaceAll("<", "&lt;")
        .replaceAll(">", "&gt;")
        .replaceAll('"', "&quot;")
        .replaceAll("'", "&#39;");
}

function truncateText(value, limit = 120) {
    const text = String(value ?? "").trim();
    if (text.length <= limit) {
        return text;
    }
    return `${text.slice(0, limit - 3)}...`;
}
