document.addEventListener("DOMContentLoaded", function () {
    const body = document.body;
    const sidebar = document.getElementById("sidebar");
    const sidebarToggle = document.getElementById("sidebarToggle");
    const mobileMenuButton = document.getElementById("mobileMenuButton");
    const sidebarOverlay = document.getElementById("sidebarOverlay");

    if (sidebarToggle) {
        sidebarToggle.addEventListener("click", function () {
            body.classList.toggle("sidebar-collapsed");
        });
    }

    function openMobileSidebar() {
        if (!sidebar || !sidebarOverlay) return;
        sidebar.classList.add("mobile-open");
        sidebarOverlay.classList.add("visible");
    }
    function closeMobileSidebar() {
        if (!sidebar || !sidebarOverlay) return;
        sidebar.classList.remove("mobile-open");
        sidebarOverlay.classList.remove("visible");
    }
    if (mobileMenuButton) mobileMenuButton.addEventListener("click", openMobileSidebar);
    if (sidebarOverlay) sidebarOverlay.addEventListener("click", closeMobileSidebar);
    document.querySelectorAll(".sidebar .menu-link").forEach(function (link) {
        link.addEventListener("click", function () { if (window.innerWidth <= 900) closeMobileSidebar(); });
    });
    window.addEventListener("resize", function () { if (window.innerWidth > 900) closeMobileSidebar(); });

    const settingsBody = document.getElementById("settingsBody");
    const settingsSidebarToggle = document.getElementById("settingsSidebarToggle");
    if (settingsBody && localStorage.getItem("settings-sidebar-collapsed") === "1") settingsBody.classList.add("sidebar-folded");
    settingsSidebarToggle?.addEventListener("click", function () {
        settingsBody.classList.toggle("sidebar-folded");
        localStorage.setItem("settings-sidebar-collapsed", settingsBody.classList.contains("sidebar-folded") ? "1" : "0");
    });

    const settingsTabs = document.querySelectorAll(".settings-tab");
    const settingsPanels = document.querySelectorAll(".settings-panel");
    settingsTabs.forEach(function (tab) {
        tab.addEventListener("click", function () {
            settingsTabs.forEach(item => item.classList.remove("active"));
            settingsPanels.forEach(panel => panel.classList.remove("active"));
            tab.classList.add("active");
            const selectedPanel = document.getElementById(tab.dataset.tab || "");
            if (selectedPanel) selectedPanel.classList.add("active");
        });
    });

    const counters = {};
    document.querySelectorAll(".auto-code").forEach(function (input) {
        const parts = input.value.split("-");
        counters[parts[0] || "AC"] = Math.max(counters[parts[0] || "AC"] || 0, Number(parts[1] || 1));
    });
    document.querySelectorAll(".new-definition").forEach(function (button) {
        button.addEventListener("click", function () {
            const prefix = button.dataset.prefix || "AC";
            const target = document.getElementById(button.dataset.target || "");
            counters[prefix] = (counters[prefix] || 0) + 1;
            if (target) target.value = prefix + "-" + String(counters[prefix]).padStart(4, "0");
        });
    });

    const salesTable = document.getElementById("salesEntryTable");
    if (salesTable) initSalesEntry(salesTable);
    initJournalAjaxForms();
});

function createToastContainer() {
    let container = document.getElementById("globalToastContainer");
    if (!container) {
        container = document.createElement("div");
        container.id = "globalToastContainer";
        container.className = "toast-container";
        document.body.appendChild(container);
    }
    return container;
}

function showToast(message, type = "success") {
    const container = createToastContainer();
    const toast = document.createElement("div");
    toast.className = "toast " + type;
    toast.textContent = message;
    container.appendChild(toast);
    requestAnimationFrame(() => {
        toast.classList.add("visible");
    });
    setTimeout(() => {
        toast.classList.remove("visible");
        toast.addEventListener("transitionend", () => toast.remove(), { once: true });
    }, 1500);
}

function resetJournalForm(form) {
    form.reset();
    const journalIdInput = form.querySelector('input[name="journal_id"]');
    if (journalIdInput) {
        journalIdInput.value = "";
        journalIdInput.removeAttribute('value');
    }
    form.querySelectorAll('input[name="treasury_id"]').forEach((input) => input.remove());
    document.querySelectorAll('select[name="treasury_picker"], #treasuryPicker, #expenseTreasuryPicker, #otherAccountTreasuryPicker').forEach((picker) => {
        picker.value = "";
    });
    const dateInput = form.querySelector('input[name="journal_date"]');
    if (dateInput) {
        const now = new Date();
        const day = String(now.getDate()).padStart(2, "0");
        const month = String(now.getMonth() + 1).padStart(2, "0");
        const year = now.getFullYear();
        dateInput.value = `${day}/${month}/${year}`;
    }
    const tbody = form.querySelector(".journal-scroll-area tbody");
    if (tbody) {
        const rows = Array.from(tbody.querySelectorAll("tr"));
        rows.slice(1).forEach((row) => row.remove());
        const firstRow = rows[0];
        if (firstRow) {
            firstRow.querySelectorAll("input, select, textarea").forEach((field) => {
                if (field.tagName === "SELECT") {
                    field.selectedIndex = 0;
                } else if (field.type === "checkbox" || field.type === "radio") {
                    field.checked = false;
                } else {
                    field.value = "";
                }
            });
        }
        tbody.querySelectorAll("tr").forEach((row, index) => {
            const numberCell = row.querySelector(".row-number");
            if (numberCell) numberCell.textContent = index + 1;
        });
    }
    form.querySelectorAll("select, input").forEach((field) => {
        field.dispatchEvent(new Event("change", { bubbles: true }));
        field.dispatchEvent(new Event("input", { bubbles: true }));
    });
    const newJournalLink = form.querySelector('a.secondary-button[href="/sales"], a.secondary-button[href="/purchases"], a.secondary-button[href="/expenses"], a.secondary-button[href="/other-accounts"]');
    if (newJournalLink) newJournalLink.style.display = "none";
    const firstField = form.querySelector("select:not([disabled]), input:not([readonly]):not([disabled])");
    if (firstField) firstField.focus();
}

function submitFormWithEvent(form) {
    const event = new Event("submit", { cancelable: true, bubbles: true });
    if (form.dispatchEvent(event)) {
        form.submit();
    }
}

function initJournalAjaxSave(form) {
    if (!form || form.dataset.ajaxInitialized === "1") return;
    form.dataset.ajaxInitialized = "1";
    form.addEventListener("submit", async function (event) {
        if (event.defaultPrevented) return;
        if (form.dataset.readonly === "1") return;
        event.preventDefault();
        const action = form.action;
        const method = form.method || "POST";
        const payload = new FormData(form);
        try {
            const response = await fetch(action, {
                method,
                headers: {
                    "x-requested-with": "XMLHttpRequest",
                },
                body: payload,
            });
            const data = await response.json();
            if (!data || !data.success) {
                showToast(data?.message || "حدث خطأ أثناء الحفظ", "error");
                return;
            }
            document.querySelectorAll("dialog[open]").forEach((dialog) => dialog.close());
            showToast(data.message || "تم حفظ اليومية بنجاح", "success");
            resetJournalForm(form);
        } catch (error) {
            showToast("فشل الاتصال بالسيرفر", "error");
            console.error(error);
        }
    });
}

function initJournalAjaxForms() {
    document.querySelectorAll(".journal-entry-form").forEach((form) => initJournalAjaxSave(form));
}

function initSalesEntry(table) {
    const tbody = table.querySelector("tbody");
    const addButton = document.getElementById("addSalesRow");
    const branchSelect = document.getElementById("salesBranch");
    const form = document.getElementById("salesJournalForm");
    const dialog = document.getElementById("treasuryDialog");
    const treasuryPicker = document.getElementById("treasuryPicker");

    function numberValue(input) { return Number.parseFloat(input?.value || "0") || 0; }
    function updateRowNumbers() {
        tbody.querySelectorAll("tr").forEach((row, index) => { row.querySelector(".row-number").textContent = index + 1; });
    }
    function filterEmployees() {
        const branch = branchSelect?.value || "";
        tbody.querySelectorAll(".employee-select").forEach(function (select) {
            Array.from(select.options).forEach(function (option) {
                if (!option.value) return;
                option.hidden = Boolean(branch && option.dataset.branch !== branch);
            });
            if (select.selectedOptions[0] && select.selectedOptions[0].hidden) select.value = "";
        });
    }
    function calculateRow(row) {
        const shift = numberValue(row.querySelector(".shift-value"));
        const discount = numberValue(row.querySelector(".discount-value"));
        const difference = numberValue(row.querySelector(".difference-value"));
        const net = row.querySelector(".net-value");
        if (net && net.hasAttribute("readonly")) net.value = (shift - discount + difference).toFixed(2);
        updateMetrics();
    }
    function updateMetrics() {
        const rows = Array.from(tbody.querySelectorAll("tr"));
        let count = 0, shift = 0, discount = 0, net = 0, difference = 0;
        rows.forEach(function (row) {
            const employee = row.querySelector(".employee-select")?.value;
            const rowShift = numberValue(row.querySelector(".shift-value"));
            const rowDiscount = numberValue(row.querySelector(".discount-value"));
            const rowNet = numberValue(row.querySelector(".net-value"));
            const rowDifference = numberValue(row.querySelector(".difference-value"));
            if (employee || rowShift || rowDiscount || rowNet || rowDifference) count += 1;
            shift += rowShift; discount += rowDiscount; net += rowNet; difference += rowDifference;
        });
        document.getElementById("metricRows").textContent = count;
        document.getElementById("metricShift").textContent = shift.toFixed(2);
        document.getElementById("metricDiscount").textContent = discount.toFixed(2);
        document.getElementById("metricNet").textContent = net.toFixed(2);
        document.getElementById("metricDifference").textContent = difference.toFixed(2);
    }
    function bindRow(row) {
        row.querySelectorAll("input").forEach(input => input.addEventListener("input", () => calculateRow(row)));
        row.querySelector(".employee-select")?.addEventListener("change", updateMetrics);
        row.querySelector(".remove-row")?.addEventListener("click", function () {
            if (tbody.rows.length === 1) {
                row.querySelectorAll("input").forEach(input => input.value = "");
                row.querySelector("select").value = "";
            } else row.remove();
            updateRowNumbers(); updateMetrics();
        });
        row.querySelectorAll("input, select").forEach(function (field) {
            field.addEventListener("keydown", function (event) {
                if (event.key !== "Enter") return;
                event.preventDefault();
                const fields = Array.from(row.querySelectorAll("select, input:not([readonly])"));
                const index = fields.indexOf(field);
                if (index < fields.length - 1) fields[index + 1].focus();
                else {
                    const newRow = addRow();
                    newRow.querySelector(".employee-select")?.focus();
                }
            });
        });
    }
    function addRow() {
        const source = tbody.querySelector("tr");
        const row = source.cloneNode(true);
        row.querySelectorAll("input").forEach(input => input.value = "");
        row.querySelectorAll("select").forEach(select => select.value = "");
        tbody.appendChild(row);
        bindRow(row); updateRowNumbers(); filterEmployees(); updateMetrics();
        return row;
    }
    tbody.querySelectorAll("tr").forEach(bindRow);
    if (addButton) addButton.addEventListener("click", addRow);
    if (branchSelect) branchSelect.addEventListener("change", filterEmployees);
    function filterTreasuries() {
        treasuryPicker?.querySelectorAll("option[data-branch]").forEach(option => {
            option.hidden = Boolean(branchSelect?.value && option.dataset.branch !== branchSelect.value);
        });
        if (treasuryPicker?.selectedOptions[0]?.hidden) treasuryPicker.value = "";
    }
    document.getElementById("openTreasuryDialog")?.addEventListener("click", () => {
        if (!form?.reportValidity()) return;
        filterTreasuries();
        document.getElementById("dialogNet").textContent = document.getElementById("metricNet")?.textContent || "0.00";
        dialog?.showModal();
    });
    document.getElementById("closeTreasuryDialog")?.addEventListener("click", () => {
        treasuryPicker.value = "";
        dialog?.close();
    });
    document.getElementById("confirmSalesSave")?.addEventListener("click", () => {
        if (!treasuryPicker?.value) {
            showToast('اختر الخزينة', 'error');
            treasuryPicker?.focus();
            return;
        }
        let hidden = form.querySelector('input[name="treasury_id"]');
        if (!hidden) {
            hidden = document.createElement('input');
            hidden.type = 'hidden';
            hidden.name = 'treasury_id';
            form.appendChild(hidden);
        }
        hidden.value = treasuryPicker.value;
        dialog?.close();
        if (typeof form.requestSubmit === 'function') {
            form.requestSubmit();
        } else {
            submitFormWithEvent(form);
        }
    });
    branchSelect?.addEventListener("change", filterTreasuries); filterTreasuries();
    filterEmployees(); updateMetrics();
}
