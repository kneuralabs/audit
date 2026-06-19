/* ============================================================
   KNEURAUDIT — application logic
   ------------------------------------------------------------
   Sections:
     1.  Domain config   — standards, categories, SLA & theme policy
     2.  Core utilities  — dates, ids, html-escape, SLA math
     3.  Excel schema    — single column definition shared by export + import
     4.  ExcelJS loader  — lazy, on-demand CDN load
     5.  DOM references
     6.  Theme
     7.  Toast
     8.  UI — approval toggle + finding-row factory
     9.  Review / edit mode
     10. Excel export
     11. Excel import
     12. Report generation
     13. Wiring & init

   Behaviour is intentionally identical to the original single-file build;
   this module only removes duplication and isolates concerns.
   ============================================================ */
(function () {
    'use strict';

    /* ========================================================
       1. DOMAIN CONFIG
       ======================================================== */

    const STANDARDS = {
        iso20000: { name: "ISO 20000:2018", prefix: "KNISO20K", fullName: "ISO 20000:2018 (IT Service Management)",
            clauses: [
                { code: "4.1", desc: "Understanding the organization and its context: The organization must determine external and internal issues relevant to its service management system (SMS) that affect its ability to achieve intended outcomes." },
                { code: "4.2", desc: "Needs and expectations of interested parties: Identify interested parties relevant to the SMS and their requirements; monitor and review this information." },
                { code: "5.1", desc: "Leadership and commitment: Top management must demonstrate leadership by ensuring SMS policy and objectives, integrating SMS into business processes, and providing resources." },
                { code: "5.2", desc: "Policy: Establish an SMS policy appropriate to purpose, including commitment to satisfy requirements and continual improvement." },
                { code: "5.3", desc: "Organizational roles, responsibilities and authorities: Assign and communicate responsibilities for SMS performance and reporting." },
                { code: "6.1", desc: "Actions to address risks and opportunities: Plan actions to address risks/opportunities, integrate into SMS processes, evaluate effectiveness." },
                { code: "6.2", desc: "Service management objectives and planning: Establish measurable objectives at relevant functions, aligned with policy, and plan how to achieve them." },
                { code: "7.1", desc: "Resources: Determine and provide necessary resources for establishing, implementing, maintaining and improving the SMS." },
                { code: "7.2", desc: "Competence: Ensure personnel are competent based on education, training, experience; take actions to acquire necessary competence." },
                { code: "7.3", desc: "Awareness: Ensure persons doing work under organization's control are aware of SMS policy, relevant objectives, their contribution to effectiveness." },
                { code: "7.4", desc: "Communication: Determine internal and external communications relevant to SMS, including what, when, whom, and how to communicate." },
                { code: "8.1", desc: "Operational planning and control: Plan, implement and control processes needed to meet service requirements, including establishing criteria for processes." },
                { code: "8.2", desc: "Service portfolio: Manage the service portfolio, including service definitions, service levels, service catalog, and lifecycle of services." },
                { code: "8.3", desc: "Relationship and agreement: Manage relationships with customers and suppliers; formalize service level agreements (SLAs) and contracts." },
                { code: "8.4", desc: "Supply and demand: Manage capacity, availability, and demand for services; ensure suppliers meet agreed requirements." },
                { code: "8.5", desc: "Service design, build and transition: Design, build and transition new or changed services to meet agreed requirements and minimize risk." },
                { code: "8.6", desc: "Resolution and fulfillment: Resolve incidents, service requests, and problems effectively; manage known errors." },
                { code: "8.7", desc: "Service assurance: Ensure service continuity, availability, information security, and capacity are managed to meet agreed levels." },
                { code: "9.1", desc: "Monitoring, measurement, analysis and evaluation: Monitor, measure, analyze and evaluate SMS performance and service effectiveness." },
                { code: "9.2", desc: "Internal audit: Conduct internal audits at planned intervals to determine if SMS conforms to requirements and is effectively implemented." },
                { code: "9.3", desc: "Management review: Top management reviews the SMS at planned intervals to ensure continuing suitability, adequacy, and effectiveness." },
                { code: "10.1", desc: "Nonconformity and corrective action: When nonconformity occurs, react, control and correct it; evaluate need for action to eliminate causes." },
                { code: "10.2", desc: "Continual improvement: Continually improve the suitability, adequacy and effectiveness of the SMS." }
            ] },
        iso42001: { name: "ISO 42001:2023", prefix: "KNISO42K", fullName: "ISO 42001:2023 (AI Management System)",
            clauses: [
                { code: "4.1", desc: "Understanding organization and context related to AI: Determine internal/external issues that affect the AI management system (AIMS) and its intended outcomes." },
                { code: "4.2", desc: "Needs and expectations of interested parties (AI): Identify relevant stakeholders regarding AI systems (users, regulators, affected individuals) and their requirements." },
                { code: "4.3", desc: "Determining scope of AI management system: Define boundaries and applicability of AIMS considering organization's AI activities, products, and services." },
                { code: "4.4", desc: "AI management system: Establish, implement, maintain and continually improve AIMS including processes, documentation and governance." },
                { code: "5.1", desc: "Leadership and commitment for AI: Top management to demonstrate accountability for effectiveness of AIMS, ensure policy and resources." },
                { code: "5.2", desc: "AI policy: Establish policy appropriate to purpose, including commitment to responsible AI, fairness, transparency and continual improvement." },
                { code: "5.3", desc: "Roles, responsibilities for AI: Assign and communicate AI governance roles, including AI ethics officer or committee where relevant." },
                { code: "6.1", desc: "Actions to address AI risks & opportunities: Identify AI-specific risks (bias, safety, security) and opportunities; plan actions to address." },
                { code: "6.2", desc: "AI objectives and planning: Establish measurable AI objectives at relevant functions, aligned with AI policy, and plan actions to achieve them." },
                { code: "7.1", desc: "Resources for AI: Determine and provide resources needed for AI system development, deployment, monitoring and AIMS operation." },
                { code: "7.2", desc: "Competence for AI: Ensure AI personnel have necessary skills in data science, ethics, and risk management; maintain records of competence." },
                { code: "7.3", desc: "Awareness for AI: Ensure all relevant personnel are aware of AI policy, ethical principles, and their contribution to AIMS effectiveness." },
                { code: "7.4", desc: "Communication for AI: Determine internal/external communications related to AI system capabilities, limitations, and impacts." },
                { code: "7.5", desc: "Documented information: Maintain documented information supporting AIMS, including AI system documentation, impact assessments." },
                { code: "8.1", desc: "Operational planning & control for AI systems: Plan, implement and control processes for AI lifecycle (design, development, validation, deployment)." },
                { code: "8.2", desc: "AI system impact assessment: Conduct AI impact assessments considering fundamental rights, safety, bias, and societal effects." },
                { code: "8.3", desc: "AI system design and development: Apply risk-based approach to AI design, including data governance, model selection, and testing." },
                { code: "8.4", desc: "Data management: Ensure data quality, provenance, privacy, and representativeness to avoid bias and ensure lawful processing." },
                { code: "8.5", desc: "AI system deployment: Establish deployment criteria, including human oversight, fallback procedures, and user notification." },
                { code: "8.6", desc: "AI system monitoring and performance: Continuously monitor AI system performance, detect drift, and trigger corrective actions." },
                { code: "9.1", desc: "Monitoring, measurement for AI: Monitor AI system outputs, user satisfaction, and compliance with AI policy; analyze data for improvement." },
                { code: "9.2", desc: "Internal audit for AI: Conduct internal audits of AIMS at planned intervals to verify conformity and effectiveness." },
                { code: "9.3", desc: "Management review for AI: Top management reviews AIMS performance, AI risks, audit results, and resource needs." },
                { code: "10.1", desc: "Nonconformity & corrective action for AI: Address AI nonconformities (e.g., bias incidents, safety failures), investigate causes, prevent recurrence." },
                { code: "10.2", desc: "Continual improvement for AI: Enhance AIMS suitability, adequacy and effectiveness based on monitoring and emerging AI practices." }
            ] },
        aims_framework: { name: "AIMS Framework", prefix: "KNAIMS", fullName: "AIMS AI Governance Framework (AI & Governance)",
            clauses: [
                { code: "GOV-01", desc: "AI Governance Oversight & Board Responsibility: Establish a governance body (e.g., AI ethics board) accountable for AI strategy, risk appetite, and compliance with regulations." },
                { code: "GOV-02", desc: "AI Risk Management & Impact Assessment: Implement a systematic risk management process covering AI-specific risks (algorithmic bias, safety, security, societal impact). Mandatory impact assessments for high-risk AI." },
                { code: "GOV-03", desc: "AI Transparency & Explainability Requirements: Ensure AI systems are transparent, and decisions can be explained to stakeholders. Provide meaningful information about AI capabilities and limitations." },
                { code: "GOV-04", desc: "AI Fairness, Non-discrimination & Bias Mitigation: Implement bias detection and mitigation measures across AI lifecycle. Regularly test for discriminatory outcomes and ensure fairness metrics are defined." },
                { code: "GOV-05", desc: "AI Accountability & Traceability: Maintain traceability of AI system development, data sources, and decisions. Assign clear accountability for AI outcomes and establish escalation paths." },
                { code: "GOV-06", desc: "AI Security, Privacy & Data Protection: Apply security controls to AI models and data. Comply with data protection laws (e.g., GDPR). Ensure privacy by design in AI systems." },
                { code: "GOV-07", desc: "AI Lifecycle Management (Design to Decommission): Govern all phases: conception, data acquisition, model development, validation, deployment, operations, and retirement. Include validation gates." },
                { code: "GOV-08", desc: "Continuous Monitoring & Performance Evaluation: Establish ongoing monitoring of AI performance, including accuracy, drift, and user feedback. Define key performance indicators (KPIs) and regular reporting." },
                { code: "GOV-09", desc: "Compliance with Regulatory & Ethical Requirements: Map AI systems to applicable laws (EU AI Act, etc.) and ethical standards. Maintain evidence of compliance and conduct periodic audits." },
                { code: "GOV-10", desc: "Stakeholder Engagement & Transparency Reporting: Engage with affected stakeholders (users, employees, public). Publish transparency reports on AI usage, incidents, and corrective actions taken." }
            ] }
    };

    /* Single source of truth for every category-derived value:
       SLA days, on-screen row class, table pill colours, report colours,
       and the Excel row fill. Previously these lived in 5 separate places. */
    const NO_DEADLINE = 999; // sentinel: category never locks (Good Practice)

    const CATEGORY_CONFIG = {
        "Major Non-Compliance": {
            slaDays: 20, rowClass: 'row-major',
            report: { color: '#991b1b', bg: '#fef2f2', border: '#fecaca', dot: '#dc2626' },
            xlsxRowFill: 'FFFFE6E6'
        },
        "Minor Non-Compliance": {
            slaDays: 15, rowClass: 'row-minor',
            report: { color: '#9a3412', bg: '#fff7ed', border: '#fed7aa', dot: '#ea580c' },
            xlsxRowFill: 'FFFFF0E0'
        },
        "Observation": {
            slaDays: 10, rowClass: 'row-observation',
            report: { color: '#854d0e', bg: '#fefce8', border: '#fde68a', dot: '#ca8a04' },
            xlsxRowFill: 'FFFFFACD'
        },
        "Good Practice": {
            slaDays: NO_DEADLINE, rowClass: 'row-good',
            report: { color: '#166534', bg: '#f0fdf4', border: '#bbf7d0', dot: '#16a34a' },
            xlsxRowFill: 'FFE6F4EA'
        }
    };
    const CATEGORIES = Object.keys(CATEGORY_CONFIG);

    /* Single source of truth for approval-status presentation. */
    const STATUS_CONFIG = {
        approved: { label: 'Approved', report: { color: '#166534', bg: '#f0fdf4', border: '#bbf7d0' },
            xlsxFill: 'FFE8F5E9', xlsxFont: 'FF2E7D32' },
        rejected: { label: 'Rejected', report: { color: '#991b1b', bg: '#fef2f2', border: '#fecaca' },
            xlsxFill: 'FFFDECEA', xlsxFont: 'FFB91C1C' },
        pending:  { label: 'Pending',  report: { color: '#92400e', bg: '#fffbeb', border: '#fde68a' },
            xlsxFill: 'FFFFF9E6', xlsxFont: 'FFE67E22' }
    };
    const statusKey = (status) => (status === 'approved' || status === 'rejected') ? status : 'pending';

    /* ========================================================
       2. CORE UTILITIES
       ======================================================== */

    /** ISO-ish stamp "YYYY-MM-DD HH:MM:SS" used for timestamp + export meta. */
    function nowStamp() {
        return new Date().toISOString().slice(0, 19).replace('T', ' ');
    }

    /** Parse a "YYYY-MM-DD HH:MM:SS" stamp back into a Date. */
    function parseStamp(stamp) {
        return new Date(String(stamp).replace(' ', 'T'));
    }

    function addDays(date, days) {
        const d = new Date(date);
        d.setDate(date.getDate() + days);
        return d;
    }

    /** Escape user-supplied text before it is interpolated into report HTML. */
    function escapeHtml(value) {
        return String(value == null ? '' : value)
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;')
            .replace(/'/g, '&#39;');
    }

    function slaDaysFor(category) {
        const cfg = CATEGORY_CONFIG[category];
        return cfg ? cfg.slaDays : NO_DEADLINE;
    }

    /**
     * Shared SLA computation used by both the on-screen chip and Excel export.
     * Returns the raw facts; each caller formats its own labels.
     */
    function computeSla(baseDate, category, now = new Date()) {
        const days = slaDaysFor(category);
        const hasDeadline = days !== NO_DEADLINE;
        const validBase = baseDate instanceof Date && !isNaN(baseDate.getTime());
        const dueDate = (hasDeadline && validBase) ? addDays(baseDate, days) : null;
        return {
            days,
            hasDeadline,
            validBase,
            dueDate,
            isOverdue: !!dueDate && now > dueDate,
            // ceil-day delta used by the live chip
            diffDays: dueDate ? Math.ceil((dueDate - now) / 86400000) : null
        };
    }

    /* ========================================================
       3. EXCEL COLUMN SCHEMA  (single definition, shared by export + import)
       ======================================================== */

    // key      → logical field
    // header   → printed header text (row 7)
    // width    → column width
    // editable → cell left unlocked under sheet protection (PM-fillable)
    const EXCEL_COLUMNS = [
        { key: 'clauseCode',        header: 'Clause Code',            width: 14, editable: false },
        { key: 'clauseDisplay',     header: 'Clause',                 width: 22, editable: false },
        { key: 'clauseDescription', header: 'Clause Description',     width: 38, editable: false },
        { key: 'category',          header: 'Category',               width: 20, editable: false },
        { key: 'findingDetail',     header: 'Finding Details',        width: 38, editable: false },
        { key: 'caText',            header: 'Corrective Action (CA)', width: 38, editable: true  },
        { key: 'caStatus',          header: 'CA Status',              width: 14, editable: false },
        { key: 'paText',            header: 'Preventive Action (PA)', width: 38, editable: true  },
        { key: 'paStatus',          header: 'PA Status',              width: 14, editable: false },
        { key: 'slaDue',            header: 'SLA Due Date',           width: 14, editable: false },
        { key: 'slaStatus',         header: 'SLA Status',             width: 12, editable: false }
    ];
    const colIndex = (key) => EXCEL_COLUMNS.findIndex(c => c.key === key) + 1; // 1-based
    const HEADER_ROW = 7;
    const DATA_START_ROW = 8;
    const META_LAST_COL = EXCEL_COLUMNS.length; // meta value cells merge across all columns

    /* ========================================================
       4. EXCELJS LOADER  (lazy — avoids loading ~hundreds of KB on every visit)
       ======================================================== */

    const EXCELJS_SRC = 'https://cdn.jsdelivr.net/npm/exceljs@4.4.0/dist/exceljs.min.js';
    let _exceljsPromise = null;
    function loadExcelJS() {
        if (window.ExcelJS) return Promise.resolve(window.ExcelJS);
        if (_exceljsPromise) return _exceljsPromise;
        _exceljsPromise = new Promise((resolve, reject) => {
            const s = document.createElement('script');
            s.src = EXCELJS_SRC;
            s.async = true;
            s.onload = () => window.ExcelJS ? resolve(window.ExcelJS) : reject(new Error('ExcelJS unavailable'));
            s.onerror = () => { _exceljsPromise = null; reject(new Error('Failed to load ExcelJS')); };
            document.head.appendChild(s);
        });
        return _exceljsPromise;
    }

    /* ========================================================
       5. DOM REFERENCES
       ======================================================== */

    const auditeeInput      = document.getElementById('auditeeInput');
    const auditIdField      = document.getElementById('auditIdField');
    const timestampAuto     = document.getElementById('timestampAuto');
    const standardSelect    = document.getElementById('standardSelect');
    const findingsTbody     = document.getElementById('findingsTbody');
    const addRowBtn         = document.getElementById('addRowBtn');
    const exportExcelBtn    = document.getElementById('exportExcelBtn');
    const uploadExcelInput  = document.getElementById('uploadExcelInput');
    const generateReportBtn = document.getElementById('generateReportBtn');
    const enableEditBtn     = document.getElementById('enableEditBtn');
    const reportPreviewCard = document.getElementById('reportPreviewCard');
    const closeReportBtn    = document.getElementById('closeReportBtn');
    const emptyRowMsgDiv    = document.getElementById('emptyRowMsg');
    const toastEl           = document.getElementById('toast');
    const resetBtn          = document.getElementById('resetBtn');
    const themeBtn          = document.getElementById('themeBtn');
    const themeIcon         = document.getElementById('themeIcon');

    /* ========================================================
       6. THEME
       ======================================================== */

    const SUN  = '<circle cx="10" cy="10" r="3.5" fill="currentColor"/><g stroke="currentColor" stroke-width="1.5" stroke-linecap="round"><line x1="10" y1="1.5" x2="10" y2="3"/><line x1="10" y1="17" x2="10" y2="18.5"/><line x1="1.5" y1="10" x2="3" y2="10"/><line x1="17" y1="10" x2="18.5" y2="10"/><line x1="3.9" y1="3.9" x2="5" y2="5"/><line x1="15" y1="15" x2="16.1" y2="16.1"/><line x1="16.1" y1="3.9" x2="15" y2="5"/><line x1="5" y1="15" x2="3.9" y2="16.1"/></g>';
    const MOON = '<path d="M16.5 11.5A6.5 6.5 0 1 1 8.5 3.5a5 5 0 0 0 8 8z" fill="currentColor"/>';

    function applyTheme(dark) {
        document.body.classList.toggle('dark', dark);
        themeIcon.innerHTML = dark ? SUN : MOON;
    }

    (function initTheme() {
        const saved = localStorage.getItem('kna-theme');
        let dark;
        if (saved) {
            dark = saved === 'dark';
        } else {
            const h = new Date().getHours();
            dark = h < 7 || h >= 20;
        }
        applyTheme(dark);
    })();

    themeBtn.addEventListener('click', () => {
        const isDark = !document.body.classList.contains('dark');
        applyTheme(isDark);
        localStorage.setItem('kna-theme', isDark ? 'dark' : 'light');
    });

    /* ========================================================
       7. TOAST
       ======================================================== */

    let toastTimer = null;
    function showToast(msg, type = '') {
        toastEl.textContent = msg;
        toastEl.className = 'toast-show' + (type ? ' toast-' + type : '');
        clearTimeout(toastTimer);
        toastTimer = setTimeout(() => { toastEl.className = ''; }, 2800);
    }

    /* ========================================================
       8. UI — IDENTITY, CLAUSES, APPROVAL TOGGLE, ROW FACTORY
       ======================================================== */

    let isReviewMode = false;

    function getCurrentStandardObj() { return STANDARDS[standardSelect.value]; }

    function generateAuditId() {
        const stdObj = getCurrentStandardObj();
        const now = new Date();
        const mm  = String(now.getMonth() + 1).padStart(2, '0');
        const dd  = String(now.getDate()).padStart(2, '0');
        const yy  = String(now.getFullYear()).slice(-2);
        const seq = String(Math.floor(Math.random() * 900) + 100).slice(0, 3);
        return `${stdObj.prefix}${mm}${dd}${yy}${seq}`;
    }

    function setTimestampAndId() {
        timestampAuto.value = nowStamp();
        auditIdField.value  = generateAuditId();
    }

    function buildClauseOptions(selectedCode = null) {
        const std = getCurrentStandardObj();
        let html = '<option value="">-- Select clause --</option>';
        std.clauses.forEach(cl => {
            const display = `${std.name} - ${cl.code}`;
            html += `<option value="${cl.code}" ${selectedCode === cl.code ? 'selected' : ''}>${display}</option>`;
        });
        return html;
    }

    function getDescription(code) {
        const match = getCurrentStandardObj().clauses.find(c => c.code === code);
        return match ? match.desc : '';
    }

    function applyRowColor(row, categoryValue) {
        Object.values(CATEGORY_CONFIG).forEach(cfg => row.classList.remove(cfg.rowClass));
        const cfg = CATEGORY_CONFIG[categoryValue];
        if (cfg) row.classList.add(cfg.rowClass);
    }

    /** Tri-state Approve / Reject / Pending toggle. */
    function createRadioApproval(currentValue, onToggle) {
        const wrap = document.createElement('div');
        wrap.className = 'approval-toggle-wrap';
        const toggle = document.createElement('div');
        toggle.className = 'approval-toggle is-pending';
        toggle.dataset.isToggle = 'true';
        const track = document.createElement('div');
        track.className = 'toggle-track';
        const thumb = document.createElement('div');
        thumb.className = 'toggle-thumb';
        const approveBtn = document.createElement('div');
        approveBtn.className = 'toggle-side approve-side';
        approveBtn.textContent = 'Approve';
        const rejectBtn = document.createElement('div');
        rejectBtn.className = 'toggle-side reject-side';
        rejectBtn.textContent = 'Reject';
        track.appendChild(thumb);
        track.appendChild(approveBtn);
        track.appendChild(rejectBtn);
        toggle.appendChild(track);
        const pendingLabel = document.createElement('div');
        pendingLabel.className = 'toggle-pending-label';
        pendingLabel.textContent = 'Pending';
        wrap.appendChild(toggle);
        wrap.appendChild(pendingLabel);

        let currentState = null;

        function updateUI(value) {
            toggle.classList.remove('is-approved', 'is-rejected', 'is-pending');
            if (value === 'approved') {
                toggle.classList.add('is-approved');
                pendingLabel.style.display = 'none';
            } else if (value === 'rejected') {
                toggle.classList.add('is-rejected');
                pendingLabel.style.display = 'none';
            } else {
                toggle.classList.add('is-pending');
                pendingLabel.style.display = 'block';
            }
            currentState = value || null;
        }
        updateUI(currentValue);

        approveBtn.addEventListener('click', () => {
            const next = currentState === 'approved' ? null : 'approved';
            updateUI(next); onToggle(next);
        });
        rejectBtn.addEventListener('click', () => {
            const next = currentState === 'rejected' ? null : 'rejected';
            updateUI(next); onToggle(next);
        });

        return { container: wrap, updateUI };
    }

    function createRow(clauseCode = '', findingText = '', categoryVal = 'Observation', caText = '', paText = '', caStatus = null, paStatus = null) {
        const tr = document.createElement('tr');

        const tdClause = document.createElement('td');
        const select   = document.createElement('select');
        select.className = 'clause-select';
        select.innerHTML = buildClauseOptions(clauseCode);
        if (clauseCode) select.value = clauseCode;
        tdClause.appendChild(select);

        const tdDesc   = document.createElement('td');
        const descSpan = document.createElement('span');
        descSpan.textContent = clauseCode ? getDescription(clauseCode) : '—';
        tdDesc.appendChild(descSpan);

        const tdCat    = document.createElement('td');
        const catSelect = document.createElement('select');
        catSelect.className = 'category-select';
        CATEGORIES.forEach(c => {
            const opt = document.createElement('option');
            opt.value = c; opt.textContent = c;
            if (c === categoryVal) opt.selected = true;
            catSelect.appendChild(opt);
        });
        tdCat.appendChild(catSelect);

        const slaChip = document.createElement('div');
        slaChip.className = 'sla-chip';
        tdCat.appendChild(slaChip);

        function updateSlaChip() {
            const sla = computeSla(parseStamp(timestampAuto.value), catSelect.value);
            if (!sla.hasDeadline) { slaChip.textContent = 'No deadline'; slaChip.className = 'sla-chip sla-none'; return; }
            if (!sla.validBase)   { slaChip.textContent = 'SLA pending'; slaChip.className = 'sla-chip sla-none'; return; }
            const d = sla.diffDays;
            if (d < 0)       { slaChip.textContent = `Overdue by ${Math.abs(d)}d`; slaChip.className = 'sla-chip sla-overdue'; }
            else if (d <= 3) { slaChip.textContent = `Due in ${d}d`;               slaChip.className = 'sla-chip sla-warning'; }
            else             { slaChip.textContent = `${d}d remaining`;            slaChip.className = 'sla-chip sla-ok'; }
        }
        updateSlaChip();

        const tdFinding  = document.createElement('td');
        const findingTA  = document.createElement('textarea');
        findingTA.className = 'finding-input'; findingTA.rows = 2; findingTA.value = findingText;
        tdFinding.appendChild(findingTA);

        const tdCAText = document.createElement('td');
        const caTA     = document.createElement('textarea');
        caTA.className = 'ca-pa-input'; caTA.rows = 2;
        caTA.placeholder = 'Corrective action (PM fills)'; caTA.value = caText;
        tdCAText.appendChild(caTA);

        const tdCAApproval = document.createElement('td');
        let currentCAStatus = caStatus;
        const { container: caToggle, updateUI: updateCAToggle } = createRadioApproval(currentCAStatus, v => { currentCAStatus = v; });
        tdCAApproval.appendChild(caToggle);

        const tdPAText = document.createElement('td');
        const paTA     = document.createElement('textarea');
        paTA.className = 'ca-pa-input'; paTA.rows = 2;
        paTA.placeholder = 'Preventive action (PM fills)'; paTA.value = paText;
        tdPAText.appendChild(paTA);

        const tdPAApproval = document.createElement('td');
        let currentPAStatus = paStatus;
        const { container: paToggle, updateUI: updatePAToggle } = createRadioApproval(currentPAStatus, v => { currentPAStatus = v; });
        tdPAApproval.appendChild(paToggle);

        const tdAction = document.createElement('td');
        tdAction.style.textAlign = 'center';
        const removeBtn = document.createElement('button');
        removeBtn.innerHTML = '&times;';
        removeBtn.setAttribute('aria-label', 'Remove row');
        removeBtn.className = 'remove-row-btn';
        tdAction.appendChild(removeBtn);

        tr.appendChild(tdClause); tr.appendChild(tdDesc);   tr.appendChild(tdCat);
        tr.appendChild(tdFinding); tr.appendChild(tdCAText); tr.appendChild(tdCAApproval);
        tr.appendChild(tdPAText);  tr.appendChild(tdPAApproval); tr.appendChild(tdAction);

        const updateRowColor = () => { applyRowColor(tr, catSelect.value); updateSlaChip(); };
        catSelect.addEventListener('change', updateRowColor);
        updateRowColor();

        select.addEventListener('change', () => {
            descSpan.textContent = select.value ? getDescription(select.value) : '—';
        });

        tr.getRowData = () => ({
            clauseCode:        select.value,
            clauseDisplay:     select.options[select.selectedIndex]?.text || '',
            clauseDescription: descSpan.textContent,
            category:          catSelect.value,
            findingDetail:     findingTA.value,
            caText:            caTA.value,
            caStatus:          currentCAStatus,
            paText:            paTA.value,
            paStatus:          currentPAStatus
        });

        tr.updateFromExcel = (newCAText, newPAText, newCaStatus, newPaStatus, newClauseCode) => {
            if (newClauseCode && select.querySelector(`option[value="${newClauseCode}"]`)) {
                select.value = newClauseCode;
                descSpan.textContent = getDescription(newClauseCode);
            }
            caTA.value = newCAText; paTA.value = newPAText;
            currentCAStatus = newCaStatus || null; currentPAStatus = newPaStatus || null;
            updateCAToggle(currentCAStatus); updatePAToggle(currentPAStatus);
        };

        removeBtn.onclick = () => { tr.remove(); toggleEmptyMsg(); };
        return tr;
    }

    function toggleEmptyMsg() {
        emptyRowMsgDiv.style.display = findingsTbody.children.length === 0 ? 'block' : 'none';
    }

    function addEmptyRow() {
        findingsTbody.appendChild(createRow('', '', 'Observation', '', '', null, null));
        toggleEmptyMsg();
        applyReviewMode(isReviewMode);
    }

    function getAllRowsData() {
        return Array.from(findingsTbody.querySelectorAll('tr')).map(r => r.getRowData?.()).filter(Boolean);
    }

    /* ========================================================
       9. REVIEW / EDIT MODE
       ======================================================== */

    const REVIEW_LOCK_SELECTORS = [
        '#auditeeInput', '#standardSelect', '.clause-select', '.category-select',
        '.finding-input', '.ca-pa-input', '#addRowBtn', '#exportExcelBtn',
        '#generateReportBtn', '.remove-row-btn'
    ];

    function applyReviewMode(reviewActive) {
        REVIEW_LOCK_SELECTORS.forEach(sel => {
            document.querySelectorAll(sel).forEach(el => { el.disabled = reviewActive; });
        });
        addRowBtn.disabled         = reviewActive;
        exportExcelBtn.disabled    = reviewActive;
        generateReportBtn.disabled = reviewActive;
    }

    function enableEditMode() {
        isReviewMode = false;
        applyReviewMode(false);
        showToast('Edit mode — all fields unlocked.', 'warn');
    }

    /* ========================================================
       10. EXCEL EXPORT
       ======================================================== */

    const XLSX = {
        AMBER:      'FFFFC000',
        WHITE:      'FFFFFFFF',
        BORDER_CLR: 'FF999999',
        META_HDR:   'FFD9D9D9',
        EDIT_FILL:  'FFFFFEF0'
    };

    async function exportExcel() {
        const rowsData = getAllRowsData();
        let ExcelJS;
        try {
            ExcelJS = await loadExcelJS();
        } catch (err) {
            showToast('Could not load the Excel engine. Check your connection.', 'error');
            return;
        }

        const workbook = new ExcelJS.Workbook();
        workbook.creator = 'KNEURAUDIT';
        const ws = workbook.addWorksheet('CAPA_SLA');

        const thinBorder = {
            top: { style: 'thin', color: { argb: XLSX.BORDER_CLR } }, left:   { style: 'thin', color: { argb: XLSX.BORDER_CLR } },
            bottom: { style: 'thin', color: { argb: XLSX.BORDER_CLR } }, right: { style: 'thin', color: { argb: XLSX.BORDER_CLR } }
        };
        const wrapAlign  = { wrapText: true, vertical: 'top' };
        const centerWrap = { wrapText: true, vertical: 'middle', horizontal: 'center' };
        const fill = (argb) => ({ type: 'pattern', pattern: 'solid', fgColor: { argb } });

        function amberHeader(cell, value) {
            cell.value = value;
            cell.font  = { bold: true, name: 'Inter', size: 10, color: { argb: 'FF000000' } };
            cell.fill  = fill(XLSX.AMBER);
            cell.alignment = centerWrap; cell.border = thinBorder;
        }
        function metaLabel(cell, value) {
            cell.value = value;
            cell.font  = { bold: true, name: 'Inter', size: 9 };
            cell.fill  = fill(XLSX.META_HDR);
            cell.alignment = { wrapText: true, vertical: 'middle' }; cell.border = thinBorder;
        }

        // Sheet protection — CA/PA cells stay editable; everything else locked.
        ws.protect('KNEURAUDIT_PROTECTED', {
            sheet: true, formatCells: false, formatColumns: false, formatRows: false,
            insertRows: false, insertColumns: false, deleteRows: false, deleteColumns: false,
            selectLockedCells: true, selectUnlockedCells: true
        });

        // Meta block (rows 1-5)
        const metaFields = [
            ['Audit ID',  auditIdField.value],
            ['Auditee',   auditeeInput.value],
            ['Timestamp', timestampAuto.value],
            ['Standard',  getCurrentStandardObj().fullName],
            ['Exported',  nowStamp()]
        ];
        metaFields.forEach(([label, val], i) => {
            const row = i + 1;
            metaLabel(ws.getCell(row, 1), label);
            const vc = ws.getCell(row, 2);
            vc.value = val; vc.font = { name: 'Inter', size: 9 };
            vc.alignment = { wrapText: true, vertical: 'middle' }; vc.border = thinBorder;
            ws.mergeCells(row, 2, row, META_LAST_COL);
        });

        // Header row (row 7) + column widths — both driven by EXCEL_COLUMNS
        EXCEL_COLUMNS.forEach((col, i) => {
            amberHeader(ws.getCell(HEADER_ROW, i + 1), col.header);
            ws.getColumn(i + 1).width = col.width;
        });
        ws.getRow(HEADER_ROW).height = 36;

        const baseDate = parseStamp(timestampAuto.value);
        const today    = new Date();

        rowsData.forEach((f, idx) => {
            const r = DATA_START_ROW + idx;
            const sla = computeSla(baseDate, f.category, today);
            const dueDateStr = sla.hasDeadline ? sla.dueDate.toISOString().slice(0, 10) : 'N/A';
            const slaStatus  = !sla.hasDeadline ? 'No lock' : (sla.isOverdue ? 'OVERDUE' : 'On track');

            // Auto-fill empty CA/PA once the SLA has expired.
            let caVal = f.caText, paVal = f.paText;
            if (sla.isOverdue && (!caVal || caVal.trim() === '')) caVal = 'No response (SLA expired)';
            if (sla.isOverdue && (!paVal || paVal.trim() === '')) paVal = 'No response (SLA expired)';

            const rowFill = (CATEGORY_CONFIG[f.category] || {}).xlsxRowFill || XLSX.WHITE;

            // Pre-mark the two editable columns so empty CA/PA cells stay unlocked + tinted.
            [colIndex('caText'), colIndex('paText')].forEach(col => {
                const editCell = ws.getCell(r, col);
                editCell.protection = { locked: false }; editCell.alignment = wrapAlign;
                editCell.fill = fill(XLSX.EDIT_FILL);
            });

            const values = {
                clauseCode: f.clauseCode, clauseDisplay: f.clauseDisplay, clauseDescription: f.clauseDescription,
                category: f.category, findingDetail: f.findingDetail,
                caText: caVal, caStatus: STATUS_CONFIG[statusKey(f.caStatus)].label,
                paText: paVal, paStatus: STATUS_CONFIG[statusKey(f.paStatus)].label,
                slaDue: dueDateStr, slaStatus
            };

            EXCEL_COLUMNS.forEach((col, ci) => {
                const cell = ws.getCell(r, ci + 1);
                const val  = values[col.key];
                const isEditable = col.editable;

                if (!val && val !== 0) {
                    if (isEditable) {
                        cell.protection = { locked: false }; cell.alignment = wrapAlign;
                        cell.fill = fill(XLSX.EDIT_FILL);
                    }
                    return;
                }

                cell.value = val; cell.font = { name: 'Inter', size: 9 };
                cell.alignment = wrapAlign; cell.border = thinBorder;
                cell.protection = { locked: !isEditable };

                if (col.key === 'caStatus' || col.key === 'paStatus') {
                    const sc = STATUS_CONFIG[statusKey(col.key === 'caStatus' ? f.caStatus : f.paStatus)];
                    cell.fill = fill(sc.xlsxFill);
                    cell.alignment = { ...wrapAlign, horizontal: 'center' };
                    cell.font = { name: 'Inter', size: 9, bold: true, color: { argb: sc.xlsxFont } };
                } else if (col.key === 'slaStatus') {
                    const bg = sla.isOverdue ? 'FFFDECEA' : (!sla.hasDeadline ? 'FFF5F5F5' : 'FFE8F5E9');
                    const fg = sla.isOverdue ? 'FFB91C1C' : (!sla.hasDeadline ? 'FF888888' : 'FF2E7D32');
                    cell.fill = fill(bg);
                    cell.alignment = { ...wrapAlign, horizontal: 'center' };
                    cell.font = { name: 'Inter', size: 9, bold: sla.isOverdue, color: { argb: fg } };
                } else if (isEditable) {
                    cell.fill = fill(XLSX.EDIT_FILL);
                } else {
                    cell.fill = fill(rowFill);
                }
            });

            ws.getRow(r).height = 54;
        });

        ws.views      = [{ state: 'frozen', xSplit: 0, ySplit: HEADER_ROW, activeCell: 'A8' }];
        ws.autoFilter = { from: { row: HEADER_ROW, column: 1 }, to: { row: HEADER_ROW, column: META_LAST_COL } };

        const buffer = await workbook.xlsx.writeBuffer();
        const blob   = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        const link   = document.createElement('a');
        link.href     = URL.createObjectURL(blob);
        link.download = `KNEURAUDIT_${auditIdField.value}.xlsx`;
        link.click();
        URL.revokeObjectURL(link.href);
        showToast('Excel exported successfully.');
    }

    /* ========================================================
       11. EXCEL IMPORT
       ======================================================== */

    function normalizeStatus(raw) {
        const v = String(raw || '').toLowerCase();
        return v === 'approved' ? 'approved' : v === 'rejected' ? 'rejected' : null;
    }

    async function importExcelAndRebuild(file) {
        try {
            const ExcelJS  = await loadExcelJS();
            const workbook = new ExcelJS.Workbook();
            await workbook.xlsx.load(await file.arrayBuffer());
            const ws = workbook.getWorksheet(1);

            const auditIdCell   = ws.getCell('B1')?.value?.toString() || '';
            const auditeeCell   = ws.getCell('B2')?.value?.toString() || '';
            const timestampCell = ws.getCell('B3')?.value?.toString() || '';
            const standardCell  = ws.getCell('B4')?.value?.toString() || '';
            if (auditeeCell)   auditeeInput.value  = auditeeCell;
            if (auditIdCell)   auditIdField.value  = auditIdCell;
            if (timestampCell) timestampAuto.value = timestampCell;
            for (const key in STANDARDS) {
                if (STANDARDS[key].fullName === standardCell || standardCell.includes(STANDARDS[key].name)) {
                    standardSelect.value = key; break;
                }
            }

            const rowsData = [];
            for (let r = DATA_START_ROW; r <= ws.rowCount; r++) {
                const row        = ws.getRow(r);
                const clauseCode = row.getCell(colIndex('clauseCode')).value?.toString() || '';
                if (!clauseCode.trim()) continue;
                rowsData.push({
                    clauseCode,
                    category: row.getCell(colIndex('category')).value?.toString() || 'Observation',
                    finding:  row.getCell(colIndex('findingDetail')).value?.toString() || '',
                    caText:   row.getCell(colIndex('caText')).value?.toString() || '',
                    caStatus: normalizeStatus(row.getCell(colIndex('caStatus'))?.value?.toString()),
                    paText:   row.getCell(colIndex('paText')).value?.toString() || '',
                    paStatus: normalizeStatus(row.getCell(colIndex('paStatus'))?.value?.toString())
                });
            }

            findingsTbody.innerHTML = '';
            for (const rd of rowsData) {
                const newRow = createRow(rd.clauseCode, rd.finding, rd.category, rd.caText, rd.paText, rd.caStatus, rd.paStatus);
                findingsTbody.appendChild(newRow);
                const catSel = newRow.querySelector('.category-select');
                if (catSel) catSel.dispatchEvent(new Event('change'));
            }
            toggleEmptyMsg();
            isReviewMode = true;
            applyReviewMode(true);
            showToast(`Imported ${rowsData.length} finding${rowsData.length !== 1 ? 's' : ''} — approval mode active.`);
        } catch (err) {
            showToast('Error reading Excel file.', 'error');
        }
    }

    /* ========================================================
       12. REPORT GENERATION
       ======================================================== */

    function getCategoryMeta(cat) {
        const cfg = CATEGORY_CONFIG[cat];
        return cfg ? cfg.report : { color: '#1d1d1f', bg: '#f5f5f7', border: '#e0e0e5', dot: '#aeaeb2' };
    }

    function getStatusMeta(status) {
        const sc = STATUS_CONFIG[statusKey(status)];
        return { label: sc.label, color: sc.report.color, bg: sc.report.bg, border: sc.report.border };
    }

    function buildSummaryStats(rowsData) {
        const counts = { 'Major Non-Compliance': 0, 'Minor Non-Compliance': 0, 'Observation': 0, 'Good Practice': 0 };
        let caApproved = 0, caRejected = 0, caPending = 0, paApproved = 0, paRejected = 0, paPending = 0;
        rowsData.forEach(f => {
            if (counts[f.category] !== undefined) counts[f.category]++;
            if (f.caStatus === 'approved') caApproved++; else if (f.caStatus === 'rejected') caRejected++; else caPending++;
            if (f.paStatus === 'approved') paApproved++; else if (f.paStatus === 'rejected') paRejected++; else paPending++;
        });
        return { total: rowsData.length, counts, caApproved, caRejected, caPending, paApproved, paRejected, paPending };
    }

    function buildReportHTML(rowsData) {
        const std         = getCurrentStandardObj();
        const auditId     = auditIdField.value;
        const auditee     = auditeeInput.value;
        const ts          = timestampAuto.value;
        const stats       = buildSummaryStats(rowsData);
        const generatedAt = new Date().toLocaleString('en-US', { dateStyle: 'long', timeStyle: 'short' });
        const dash        = '<span style="color:#aeaeb2;font-style:italic;">—</span>';

        const statCards = [
            { label: 'Total Findings', value: stats.total,                         color: '#0071e3', bg: '#eff6ff' },
            { label: 'Major NC',       value: stats.counts['Major Non-Compliance'], color: '#dc2626', bg: '#fef2f2' },
            { label: 'Minor NC',       value: stats.counts['Minor Non-Compliance'], color: '#ea580c', bg: '#fff7ed' },
            { label: 'Observations',   value: stats.counts['Observation'],          color: '#ca8a04', bg: '#fefce8' },
            { label: 'Good Practice',  value: stats.counts['Good Practice'],        color: '#16a34a', bg: '#f0fdf4' },
        ].map(s => `<div style="flex:1;min-width:100px;background:${s.bg};border:1px solid ${s.color}22;border-radius:10px;padding:14px 16px;text-align:center;">
            <div style="font-size:2rem;font-weight:700;color:${s.color};line-height:1.1;">${s.value}</div>
            <div style="font-size:0.72rem;color:#6e6e73;margin-top:4px;font-weight:500;">${s.label}</div>
        </div>`).join('');

        const capaRow = (label, approved, rejected, pending) => `<tr>
            <td style="padding:10px 12px;font-weight:500;color:#3a3a3c;">${label}</td>
            <td style="padding:10px 12px;text-align:center;"><span style="background:#f0fdf4;color:#166534;border:1px solid #bbf7d0;border-radius:20px;padding:2px 10px;font-size:0.8rem;font-weight:500;">✓ ${approved} Approved</span></td>
            <td style="padding:10px 12px;text-align:center;"><span style="background:#fef2f2;color:#991b1b;border:1px solid #fecaca;border-radius:20px;padding:2px 10px;font-size:0.8rem;font-weight:500;">✗ ${rejected} Rejected</span></td>
            <td style="padding:10px 12px;text-align:center;"><span style="background:#fffbeb;color:#92400e;border:1px solid #fde68a;border-radius:20px;padding:2px 10px;font-size:0.8rem;font-weight:500;">⏳ ${pending} Pending</span></td>
        </tr>`;

        const findingRows = rowsData.length === 0
            ? `<tr><td colspan="6" style="text-align:center;padding:32px;color:#aeaeb2;font-style:italic;">No findings recorded.</td></tr>`
            : rowsData.map((f, idx) => {
                const cm  = getCategoryMeta(f.category);
                const cas = getStatusMeta(f.caStatus);
                const pas = getStatusMeta(f.paStatus);
                return `<tr style="border-left:4px solid ${cm.dot};">
                    <td style="padding:10px 12px;font-weight:600;color:#3a3a3c;font-size:0.8rem;">${idx + 1}</td>
                    <td style="padding:10px 12px;font-size:0.8rem;">
                        <div style="font-weight:600;color:#1d1d1f;">${escapeHtml(f.clauseDisplay)}</div>
                        <div style="font-size:0.72rem;color:#6e6e73;margin-top:3px;line-height:1.4;">${escapeHtml(f.clauseDescription || '')}</div>
                    </td>
                    <td style="padding:10px 12px;">
                        <span style="background:${cm.bg};color:${cm.color};border:1px solid ${cm.border};border-radius:20px;padding:2px 9px;font-size:0.7rem;font-weight:500;white-space:nowrap;">${escapeHtml(f.category)}</span>
                    </td>
                    <td style="padding:10px 12px;font-size:0.8rem;color:#3a3a3c;">${f.findingDetail ? escapeHtml(f.findingDetail) : dash}</td>
                    <td style="padding:10px 12px;">
                        <div style="font-size:0.78rem;color:#3a3a3c;margin-bottom:5px;">${f.caText ? escapeHtml(f.caText) : dash}</div>
                        <span style="background:${cas.bg};color:${cas.color};border:1px solid ${cas.border};border-radius:20px;padding:1px 8px;font-size:0.68rem;font-weight:500;">${cas.label}</span>
                    </td>
                    <td style="padding:10px 12px;">
                        <div style="font-size:0.78rem;color:#3a3a3c;margin-bottom:5px;">${f.paText ? escapeHtml(f.paText) : dash}</div>
                        <span style="background:${pas.bg};color:${pas.color};border:1px solid ${pas.border};border-radius:20px;padding:1px 8px;font-size:0.68rem;font-weight:500;">${pas.label}</span>
                    </td>
                </tr>`;
            }).join('');

        const metaGrid = [['Audit ID', auditId], ['Auditee', auditee], ['SLA Base Date', ts], ['Standard', std.fullName]]
            .map(([l, v]) => `<div><div style="font-size:0.72rem;color:#aeaeb2;font-weight:500;text-transform:uppercase;letter-spacing:0.4px;margin-bottom:3px;">${escapeHtml(l)}</div><div style="font-size:0.88rem;font-weight:600;color:#1d1d1f;">${escapeHtml(v)}</div></div>`).join('');

        return `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1.0">
<title>KNEURAUDIT™ Report · ${escapeHtml(auditId)}</title>
<link rel="preload" as="style" href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap" />
<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap" media="print" onload="this.media='all'" />
<noscript><link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap" /></noscript>
<style>
  * { margin:0; padding:0; box-sizing:border-box; }
  body { font-family:'Inter',-apple-system,BlinkMacSystemFont,sans-serif; background:#f5f5f7; color:#1d1d1f; line-height:1.6; }
  .page { max-width:960px; margin:0 auto; padding:2.5rem 2rem 4rem; }
  @media print {
    body { background:#fff; }
    .no-print { display:none !important; }
    .page { padding:1rem; max-width:100%; }
    table { page-break-inside:auto; }
    tr { page-break-inside:avoid; page-break-after:auto; }
    thead { display:table-header-group; }
  }
</style>
</head>
<body>
<div class="page">
  <div class="no-print" style="display:flex;justify-content:flex-end;margin-bottom:1.5rem;gap:10px;">
    <button onclick="window.print()" style="font-family:'Inter',sans-serif;font-size:0.875rem;font-weight:500;background:#1d1d1f;color:#fff;border:none;border-radius:8px;padding:9px 20px;cursor:pointer;">⬇ Save / Print PDF</button>
    <button onclick="window.close()" style="font-family:'Inter',sans-serif;font-size:0.875rem;font-weight:500;background:none;color:#3a3a3c;border:1px solid #c7c7cc;border-radius:8px;padding:9px 20px;cursor:pointer;">✕ Close</button>
  </div>
  <div style="background:#fff;border:1px solid #e0e0e5;border-radius:14px;padding:2rem 2.25rem 1.75rem;margin-bottom:1.5rem;">
    <div style="display:flex;align-items:flex-start;justify-content:space-between;flex-wrap:wrap;gap:1rem;">
      <div>
        <div style="font-size:2.2rem;font-weight:700;letter-spacing:-1px;color:#1d1d1f;">KNEURAUDIT™</div>
        <div style="font-size:0.9rem;color:#6e6e73;margin-top:2px;">Corrective &amp; Preventive Action (CAPA) Audit Report</div>
      </div>
      <div style="text-align:right;">
        <div style="font-size:0.72rem;color:#aeaeb2;text-transform:uppercase;letter-spacing:0.5px;font-weight:500;">Generated</div>
        <div style="font-size:0.85rem;color:#3a3a3c;font-weight:500;">${escapeHtml(generatedAt)}</div>
      </div>
    </div>
    <div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(180px,1fr));gap:1rem;margin-top:1.5rem;padding-top:1.5rem;border-top:1px solid #f0f0f2;">
      ${metaGrid}
    </div>
  </div>
  <div style="background:#fff;border:1px solid #e0e0e5;border-radius:14px;padding:1.5rem 1.75rem;margin-bottom:1.5rem;">
    <div style="font-size:0.78rem;font-weight:600;color:#6e6e73;text-transform:uppercase;letter-spacing:0.5px;margin-bottom:1rem;">Summary</div>
    <div style="display:flex;gap:12px;flex-wrap:wrap;">${statCards}</div>
  </div>
  <div style="background:#fff;border:1px solid #e0e0e5;border-radius:14px;padding:1.5rem 1.75rem;margin-bottom:1.5rem;">
    <div style="font-size:0.78rem;font-weight:600;color:#6e6e73;text-transform:uppercase;letter-spacing:0.5px;margin-bottom:1rem;">CAPA Approval Status</div>
    <table style="width:100%;border-collapse:collapse;font-size:0.85rem;">
      <thead><tr style="background:#f5f5f7;">
        <th style="padding:10px 12px;text-align:left;font-weight:500;color:#6e6e73;font-size:0.75rem;border-bottom:1px solid #e0e0e5;">Action Type</th>
        <th style="padding:10px 12px;text-align:center;font-weight:500;color:#6e6e73;font-size:0.75rem;border-bottom:1px solid #e0e0e5;">Approved</th>
        <th style="padding:10px 12px;text-align:center;font-weight:500;color:#6e6e73;font-size:0.75rem;border-bottom:1px solid #e0e0e5;">Rejected</th>
        <th style="padding:10px 12px;text-align:center;font-weight:500;color:#6e6e73;font-size:0.75rem;border-bottom:1px solid #e0e0e5;">Pending</th>
      </tr></thead>
      <tbody>
        ${capaRow('Corrective Actions (CA)', stats.caApproved, stats.caRejected, stats.caPending)}
        ${capaRow('Preventive Actions (PA)', stats.paApproved, stats.paRejected, stats.paPending)}
      </tbody>
    </table>
  </div>
  <div style="background:#fff;border:1px solid #e0e0e5;border-radius:14px;padding:1.5rem 1.75rem;margin-bottom:1.5rem;">
    <div style="font-size:0.78rem;font-weight:600;color:#6e6e73;text-transform:uppercase;letter-spacing:0.5px;margin-bottom:1rem;">Detailed Findings (${rowsData.length})</div>
    <div style="overflow-x:auto;">
      <table style="width:100%;border-collapse:collapse;font-size:0.85rem;min-width:700px;">
        <thead><tr style="background:#f5f5f7;">
          <th style="padding:10px 12px;text-align:left;font-weight:500;color:#6e6e73;font-size:0.72rem;border-bottom:1px solid #e0e0e5;width:3%;">#</th>
          <th style="padding:10px 12px;text-align:left;font-weight:500;color:#6e6e73;font-size:0.72rem;border-bottom:1px solid #e0e0e5;width:22%;">Clause</th>
          <th style="padding:10px 12px;text-align:left;font-weight:500;color:#6e6e73;font-size:0.72rem;border-bottom:1px solid #e0e0e5;width:13%;">Category</th>
          <th style="padding:10px 12px;text-align:left;font-weight:500;color:#6e6e73;font-size:0.72rem;border-bottom:1px solid #e0e0e5;width:22%;">Finding Details</th>
          <th style="padding:10px 12px;text-align:left;font-weight:500;color:#6e6e73;font-size:0.72rem;border-bottom:1px solid #e0e0e5;width:20%;">Corrective Action</th>
          <th style="padding:10px 12px;text-align:left;font-weight:500;color:#6e6e73;font-size:0.72rem;border-bottom:1px solid #e0e0e5;width:20%;">Preventive Action</th>
        </tr></thead>
        <tbody>${findingRows}</tbody>
      </table>
    </div>
  </div>
  <div style="font-size:0.75rem;color:#aeaeb2;text-align:center;padding-top:1rem;">
    KNEURAUDIT™ &nbsp;·&nbsp; Confidential Audit Document &nbsp;·&nbsp; ${escapeHtml(auditId)} &nbsp;·&nbsp; ${escapeHtml(generatedAt)}
  </div>
</div>
<script>
  window.addEventListener('load', function() { setTimeout(function() { window.print(); }, 800); });
<\/script>
</body>
</html>`;
    }

    function openReportWindow() {
        const rowsData  = getAllRowsData();
        const reportWin = window.open('', '_blank');
        if (!reportWin) { showToast('Pop-up blocked — please allow pop-ups and try again.', 'error'); return; }
        reportWin.document.open();
        reportWin.document.write(buildReportHTML(rowsData));
        reportWin.document.close();
        showToast('Report opened — PDF dialog will appear shortly.');
    }

    /* ========================================================
       13. WIRING & INIT
       ======================================================== */

    standardSelect.addEventListener('change', () => {
        if (findingsTbody.children.length > 0 && !confirm("Change standard? All current findings will be cleared.")) return;
        findingsTbody.innerHTML = '';
        const defaultClause = getCurrentStandardObj().clauses[0]?.code || '';
        if (defaultClause) findingsTbody.appendChild(createRow(defaultClause, 'Initial audit finding', 'Observation', '', '', null, null));
        else addEmptyRow();
        toggleEmptyMsg();
        auditIdField.value = generateAuditId();
        if (isReviewMode) { isReviewMode = false; applyReviewMode(false); }
    });

    addRowBtn.addEventListener('click', addEmptyRow);
    exportExcelBtn.addEventListener('click', exportExcel);
    uploadExcelInput.addEventListener('change', async e => {
        if (e.target.files.length) await importExcelAndRebuild(e.target.files[0]);
        uploadExcelInput.value = '';
    });

    resetBtn.addEventListener('click', () => {
        const count = findingsTbody.children.length;
        const msg   = count > 0 ? `Reset the form and delete all ${count} finding${count !== 1 ? 's' : ''}?` : 'Reset the form?';
        if (!confirm(msg)) return;
        findingsTbody.innerHTML = '';
        auditeeInput.value   = 'Nexa Innovations';
        standardSelect.value = 'iso20000';
        setTimestampAndId();
        isReviewMode = false;
        applyReviewMode(false);
        reportPreviewCard.style.display = 'none';
        toggleEmptyMsg();
        showToast('Form reset — all findings cleared.');
    });

    generateReportBtn.addEventListener('click', openReportWindow);
    closeReportBtn.addEventListener('click', () => { reportPreviewCard.style.display = 'none'; });
    enableEditBtn.addEventListener('click', enableEditMode);

    // Boot
    setTimestampAndId();
    const initClause = getCurrentStandardObj().clauses[0]?.code || '';
    if (initClause) findingsTbody.appendChild(createRow(initClause, 'Evidence of process documentation incomplete.', 'Observation', '', '', null, null));
    else addEmptyRow();
    toggleEmptyMsg();
    applyReviewMode(false);
})();
