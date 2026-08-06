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
});

function initSalesEntry(table) {
    const tbody = table.querySelector("tbody");
    const addButton = document.getElementById("addSalesRow");
    const branchSelect = document.getElementById("salesBranch");

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
    filterEmployees(); updateMetrics();
}
