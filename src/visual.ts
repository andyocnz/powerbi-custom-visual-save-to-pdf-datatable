"use strict";

import powerbi from "powerbi-visuals-api";
import { FormattingSettingsService } from "powerbi-visuals-utils-formattingmodel";
import { valueFormatter } from "powerbi-visuals-utils-formattingutils";
import { jsPDF } from "jspdf";
import autoTable from "jspdf-autotable";
import "./../style/visual.less";
import { VisualFormattingSettingsModel } from "./settings";

import DataView = powerbi.DataView;
import DataViewTable = powerbi.DataViewTable;
import DataViewObjectPropertyIdentifier = powerbi.DataViewObjectPropertyIdentifier;
import PrivilegeStatus = powerbi.PrivilegeStatus;
import IVisual = powerbi.extensibility.visual.IVisual;
import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
import VisualUpdateOptions = powerbi.extensibility.visual.VisualUpdateOptions;

const FORMAT_STRING_PROP: DataViewObjectPropertyIdentifier = {
    objectName: "general",
    propertyName: "formatString"
};

type PdfOrientation = "portrait" | "landscape";
type PdfOrientationPreference = "auto" | PdfOrientation;
type PdfPaperSize = "a4" | "letter" | "a3";

export class Visual implements IVisual {
    private root: HTMLElement;
    private toolbar: HTMLDivElement;
    private tableWrap: HTMLDivElement;
    private tableEl: HTMLTableElement;
    private prevBtn: HTMLButtonElement;
    private nextBtn: HTMLButtonElement;
    private loadAllBtn: HTMLButtonElement;
    private exportPdfBtn: HTMLButtonElement;
    private pageSizeSelect: HTMLSelectElement;
    private pageIndicator: HTMLSpanElement;
    private statusBar: HTMLDivElement;
    private columns: string[] = [];
    private rows: string[][] = [];
    private rawRows: powerbi.PrimitiveValue[][] = [];
    private colMeta: Array<{ isNumeric: boolean; isDateTime: boolean; showTotal: boolean; formatString: string }> = [];
    private colFormatters: Array<{ format(value: unknown): string }> = [];
    private pageSize: number = 25;
    private currentPage: number = 1;
    private showAll: boolean = false;

    private readonly downloadService: powerbi.extensibility.IDownloadService | undefined;
    private readonly formattingSettingsService: FormattingSettingsService;
    private formattingSettings: VisualFormattingSettingsModel;
    private dataView: DataView | undefined;

    constructor(options: VisualConstructorOptions) {
        this.downloadService = options.host.downloadService;
        this.formattingSettingsService = new FormattingSettingsService();
        this.formattingSettings = new VisualFormattingSettingsModel();

        this.root = document.createElement("div");
        this.root.className = "visual-root";

        this.toolbar = document.createElement("div");
        this.toolbar.className = "toolbar";

        this.prevBtn = this.makeButton("Prev", () => {
            if (this.currentPage > 1) {
                this.currentPage -= 1;
                this.render();
            }
        });

        this.nextBtn = this.makeButton("Next", () => {
            if (this.currentPage < this.totalPages()) {
                this.currentPage += 1;
                this.render();
            }
        });

        this.loadAllBtn = this.makeButton("Load all", () => {
            this.showAll = !this.showAll;
            this.render();
        });

        this.pageSizeSelect = document.createElement("select");
        this.pageSizeSelect.className = "page-size-select";
        [10, 20, 50, 100, 200].forEach((size) => {
            const option = document.createElement("option");
            option.value = String(size);
            option.text = `${size}/page`;
            this.pageSizeSelect.appendChild(option);
        });
        this.pageSizeSelect.addEventListener("change", () => {
            const selected = Number(this.pageSizeSelect.value);
            if (Number.isFinite(selected) && selected > 0) {
                this.pageSize = selected;
                this.currentPage = 1;
                this.showAll = false;
                this.render();
            }
        });

        this.exportPdfBtn = this.makeButton("Save PDF", () => {
            void this.handleExportPdf();
        });

        this.pageIndicator = document.createElement("span");
        this.pageIndicator.className = "page-indicator";

        this.toolbar.appendChild(this.prevBtn);
        this.toolbar.appendChild(this.nextBtn);
        this.toolbar.appendChild(this.pageSizeSelect);
        this.toolbar.appendChild(this.loadAllBtn);
        this.toolbar.appendChild(this.exportPdfBtn);
        this.toolbar.appendChild(this.pageIndicator);

        this.statusBar = document.createElement("div");
        this.statusBar.className = "status-bar";

        this.tableWrap = document.createElement("div");
        this.tableWrap.className = "table-wrap";

        this.tableEl = document.createElement("table");
        this.tableEl.className = "data-table";
        this.tableWrap.appendChild(this.tableEl);

        this.root.appendChild(this.toolbar);
        this.root.appendChild(this.statusBar);
        this.root.appendChild(this.tableWrap);
        options.element.appendChild(this.root);
        this.setStatus("Ready.");
    }

    public update(options: VisualUpdateOptions): void {
        this.dataView = options.dataViews?.[0];
        this.formattingSettings = this.formattingSettingsService.populateFormattingSettingsModel(
            VisualFormattingSettingsModel,
            this.dataView
        );

        const table: DataViewTable | undefined = this.dataView?.table;
        if (!table || !table.columns || !table.rows) {
            this.columns = [];
            this.rows = [];
            this.rawRows = [];
            this.colMeta = [];
            this.colFormatters = [];
            this.currentPage = 1;
            this.renderNoData();
            return;
        }

        this.pageSize = this.readPageSize();
        this.columns = table.columns.map((c) => c.displayName || c.queryName || "Column");

        this.colMeta = table.columns.map((c) => ({
            isNumeric: !!(c.type?.numeric || c.type?.integer),
            isDateTime: !!(c.type?.dateTime),
            showTotal: !!(c.objects?.["columnTotals"]?.["showTotal"]),
            formatString:
                valueFormatter.getFormatString(c, FORMAT_STRING_PROP) ||
                c.format ||
                "",
        }));

        this.colFormatters = this.colMeta.map((m) =>
            valueFormatter.create({ format: m.formatString })
        );

        this.rawRows = table.rows.map((row) => [...row]);

        this.rows = this.rawRows.map((row) =>
            row.map((cell, colIdx) => {
                if (cell === null || cell === undefined) return "";
                try {
                    let value: powerbi.PrimitiveValue = cell;
                    if (this.colMeta[colIdx].isDateTime && !(value instanceof Date)) {
                        const coerced = new Date(value as string | number);
                        if (!isNaN(coerced.getTime())) value = coerced;
                    }
                    return this.colFormatters[colIdx].format(value);
                } catch {
                    return this.valueToText(cell);
                }
            })
        );

        if (this.currentPage > this.totalPages()) {
            this.currentPage = Math.max(this.totalPages(), 1);
        }

        this.render();
    }

    public getFormattingModel(): powerbi.visuals.FormattingModel {
        const model = this.formattingSettingsService.buildFormattingModel(this.formattingSettings);

        // Enable fx button (dynamic measure binding) on title and subtitle slices
        for (const card of model.cards) {
            for (const group of (card as powerbi.visuals.FormattingCard).groups ?? []) {
                for (const slice of (group as powerbi.visuals.FormattingGroup).slices ?? []) {
                    const uid = (slice as powerbi.visuals.SimpleVisualFormattingSlice).uid ?? "";
                    if (uid === "exportSettings-headerText" || uid === "exportSettings-subtitleText") {
                        // instanceKind 2 = ConstantOrRule — shows the fx button in the Format pane
                        // eslint-disable-next-line @typescript-eslint/no-explicit-any
                        const descriptor = ((slice as any).control?.properties?.descriptor);
                        if (descriptor) descriptor.instanceKind = 2;
                    }
                }
            }
        }

        const columns = this.dataView?.table?.columns;
        const showTotals = !!(this.formattingSettings?.totalsCard?.showTotals?.value);
        if (showTotals && columns?.length) {
            const slices: powerbi.visuals.FormattingSlice[] = columns.map((col, idx) => ({
                uid: `columnTotals-slice-${col.queryName || idx}`,
                displayName: col.displayName || `Column ${idx + 1}`,
                control: {
                    type: "ToggleSwitch" as powerbi.visuals.FormattingComponent.ToggleSwitch,
                    properties: {
                        descriptor: {
                            objectName: "columnTotals",
                            propertyName: "showTotal",
                            selector: col.queryName ? { metadata: col.queryName } : null
                        },
                        value: !!(col.objects?.["columnTotals"]?.["showTotal"])
                    }
                }
            }));

            const columnTotalsCard: powerbi.visuals.FormattingCard = {
                uid: "columnTotals-card",
                displayName: "Column totals",
                groups: [{
                    uid: "columnTotals-group",
                    displayName: "",
                    slices
                }]
            };

            model.cards = [...model.cards, columnTotalsCard];
        }

        return model;
    }

    private renderNoData(): void {
        this.clearElement(this.tableEl);
        this.pageIndicator.textContent = "No data";
        this.prevBtn.disabled = true;
        this.nextBtn.disabled = true;
        this.pageSizeSelect.disabled = true;
        this.loadAllBtn.disabled = true;
        this.exportPdfBtn.disabled = true;
    }

    private render(): void {
        if (!this.columns.length || !this.rows.length) {
            this.renderNoData();
            return;
        }

        const rowsToRender = this.showAll ? this.rows : this.getCurrentPageRows();
        const thead = document.createElement("thead");
        const headerRow = document.createElement("tr");
        for (const col of this.columns) {
            const th = document.createElement("th");
            th.textContent = col;
            headerRow.appendChild(th);
        }
        thead.appendChild(headerRow);

        const tbody = document.createElement("tbody");
        for (const row of rowsToRender) {
            const tr = document.createElement("tr");
            for (const cell of row) {
                const td = document.createElement("td");
                td.textContent = cell;
                tr.appendChild(td);
            }
            tbody.appendChild(tr);
        }

        this.clearElement(this.tableEl);
        this.tableEl.appendChild(thead);
        this.tableEl.appendChild(tbody);

        if (this.formattingSettings?.totalsCard?.showTotals?.value) {
            const totalRow = this.computeTotalRow();
            const tfoot = document.createElement("tfoot");
            const tr = document.createElement("tr");
            totalRow.forEach((cell) => {
                const td = document.createElement("td");
                td.textContent = cell;
                tr.appendChild(td);
            });
            tfoot.appendChild(tr);
            this.tableEl.appendChild(tfoot);
        }

        const totalPages = this.totalPages();
        const rowCount = this.rows.length;
        this.prevBtn.disabled = this.showAll || this.currentPage <= 1;
        this.nextBtn.disabled = this.showAll || this.currentPage >= totalPages;
        this.pageSizeSelect.disabled = false;
        this.syncPageSizeSelect();
        this.loadAllBtn.disabled = false;
        this.exportPdfBtn.disabled = false;
        this.loadAllBtn.textContent = this.showAll ? "Restore paging" : "Load all";
        this.pageIndicator.textContent = this.showAll
            ? `Showing all ${rowCount} rows`
            : `Page ${this.currentPage}/${totalPages} (${rowCount} rows total)`;
    }

    private async handleExportPdf(): Promise<void> {
        if (!this.rows.length) {
            this.setStatus("No rows to export.", true);
            return;
        }

        if (!this.downloadService) {
            this.setStatus("Download service is unavailable in this host.", true);
            return;
        }

        this.setStatus("Checking export permission...");
        try {
            const exportStatus = await this.downloadService.exportStatus();
            if (exportStatus !== PrivilegeStatus.Allowed) {
                this.setStatus(`Export blocked: ${this.privilegeToText(exportStatus)}.`, true);
                return;
            }

            const pdfBase64 = this.buildPdfBase64();
            const fileName = this.buildExportFileName();
            this.setStatus("Saving PDF...");
            const result = await this.downloadService.exportVisualsContentExtended(
                pdfBase64,
                fileName,
                "base64",
                "Formatted data table export"
            );

            if (result?.downloadCompleted) {
                this.setStatus(`PDF saved: ${result.fileName || "datatable-export.pdf"}.`);
                return;
            }

            this.setStatus("Download API returned without completion.", true);
        } catch (error) {
            const message = error instanceof Error ? error.message : String(error);
            this.setStatus(`PDF export failed: ${message}`, true);
        }
    }

    private buildPdfBase64(): string {
        const orientationPref = this.readOrientationPreference();
        const orientation = orientationPref === "auto" ? this.pickAutoOrientation() : orientationPref;
        const paperSize = this.readPaperSize();
        const fontSize = this.readPdfFontSize();
        const titleSettings = this.readGeneralTitle();
        const subtitleSettings = this.readGeneralSubtitle();
        const headerText = titleSettings.text;

        const titleY = Math.max(18, titleSettings.fontSize + 4);
        const subtitleY = titleY + subtitleSettings.fontSize + 4;
        const dateY = (subtitleSettings.show ? subtitleY : titleY) + 12;
        const tableStartY = Math.max(50, dateY + 14);

        const now = new Date();
        const dateLabel = `Schedule as at ${now.toLocaleDateString("en-GB", { day: "numeric", month: "long", year: "numeric" })}`;

        const doc = new jsPDF({
            orientation,
            unit: "pt",
            format: paperSize
        });

        const showTotals = !!(this.formattingSettings?.totalsCard?.showTotals?.value);
        const pdfFoot = showTotals ? [this.computeTotalRow()] : undefined;

        // Compute per-column widths from data content only (ignore header length).
        // Headers will wrap within the constrained width.
        const charPt = fontSize * 0.52;
        const colPadding = 10;
        const minColWidth = 28;
        const maxColWidth = 200;
        const availableWidth = doc.internal.pageSize.getWidth() - 40; // 20pt left + 20pt right margin

        // Step 1: compute preferred width per column (data-driven, header wraps to ~3 lines max)
        const preferredWidths = this.columns.map((header, colIdx) => {
            const maxDataLen = this.rows.reduce((max, row) => Math.max(max, String(row[colIdx] ?? "").length), 4);
            const dataWidth = maxDataLen * charPt + colPadding;
            const headerFloor = Math.ceil(header.length / 3) * charPt + colPadding;
            return Math.min(maxColWidth, Math.max(minColWidth, dataWidth, headerFloor));
        });

        // Step 2: if total exceeds page width, scale all columns proportionally to fit
        const totalPreferred = preferredWidths.reduce((sum, w) => sum + w, 0);
        const scale = totalPreferred > availableWidth ? availableWidth / totalPreferred : 1;

        const columnStyles: Record<number, { cellWidth: number }> = {};
        preferredWidths.forEach((w, colIdx) => {
            columnStyles[colIdx] = { cellWidth: Math.floor(w * scale) };
        });

        autoTable(doc, {
            head: [this.columns],
            body: this.rows,
            foot: pdfFoot,
            startY: tableStartY,
            margin: { left: 20, right: 20, top: tableStartY, bottom: 24 },
            theme: "grid",
            tableWidth: "wrap",
            showHead: "everyPage",
            showFoot: showTotals ? "lastPage" : "never",
            columnStyles,
            styles: {
                fontSize,
                cellPadding: 4,
                overflow: "linebreak",
                lineColor: [210, 220, 230],
                lineWidth: 0.5
            },
            headStyles: {
                fillColor: [240, 244, 248],
                textColor: [16, 42, 67],
                fontStyle: "bold"
            },
            footStyles: {
                fillColor: [240, 244, 248],
                textColor: [16, 42, 67],
                fontStyle: "bold"
            },
            alternateRowStyles: {
                fillColor: [248, 250, 252]
            }
        });

        const pageCount = doc.getNumberOfPages();
        const pageSize = doc.internal.pageSize;
        const pageWidth = pageSize.getWidth();
        const pageHeight = pageSize.getHeight();
        for (let page = 1; page <= pageCount; page += 1) {
            doc.setPage(page);
            doc.setFont("helvetica", titleSettings.bold ? "bold" : "normal");
            doc.setFontSize(titleSettings.fontSize);
            doc.setTextColor(...titleSettings.color);
            doc.text(headerText, 20, titleY);
            if (subtitleSettings.show) {
                doc.setFont("helvetica", subtitleSettings.bold ? "bold" : "normal");
                doc.setFontSize(subtitleSettings.fontSize);
                doc.setTextColor(...subtitleSettings.color);
                doc.text(subtitleSettings.text, 20, subtitleY);
            }
            doc.setFont("helvetica", "normal");
            doc.setFontSize(8);
            doc.setTextColor(0, 0, 0);
            doc.text(dateLabel, 20, dateY);

            const pageLabel = `Page ${page} of ${pageCount}`;
            doc.text(pageLabel, pageWidth - 80, pageHeight - 10);
        }

        const dataUri = doc.output("datauristring");
        const base64 = dataUri.split(",")[1];
        return base64 || "";
    }

    private pickAutoOrientation(): PdfOrientation {
        const columnCount = this.columns.length;
        const avgColumnHeaderLen = this.columns.length
            ? this.columns.reduce((acc, c) => acc + c.length, 0) / this.columns.length
            : 0;
        return columnCount > 6 || avgColumnHeaderLen > 18 ? "landscape" : "portrait";
    }

    private computeTotalRow(): string[] {
        // Find first column that won't show a sum — use it for the "Total" label
        const labelIdx = this.colMeta.findIndex((m) => !m.showTotal || !m.isNumeric);
        return this.colMeta.map((meta, colIdx) => {
            if (colIdx === labelIdx) return "Total";
            if (!meta.showTotal || !meta.isNumeric) return "";
            let sum = 0;
            for (const row of this.rawRows) {
                const v = row[colIdx];
                if (typeof v === "number" && isFinite(v)) sum += v;
            }
            try {
                return this.colFormatters[colIdx].format(sum);
            } catch {
                return String(sum);
            }
        });
    }

    private getCurrentPageRows(): string[][] {
        const start = (this.currentPage - 1) * this.pageSize;
        const end = start + this.pageSize;
        return this.rows.slice(start, end);
    }

    private totalPages(): number {
        return Math.max(1, Math.ceil(this.rows.length / this.pageSize));
    }

    private readPageSize(): number {
        const candidate = this.formattingSettings?.pagingCard?.pageSize?.value;
        const numeric = Number(candidate);
        if (!Number.isFinite(numeric)) {
            return 25;
        }
        return Math.max(5, Math.min(1000, Math.floor(numeric)));
    }

    private syncPageSizeSelect(): void {
        const current = String(this.pageSize);
        const hasOption = Array.from(this.pageSizeSelect.options).some((o) => o.value === current);
        if (!hasOption) {
            const option = document.createElement("option");
            option.value = current;
            option.text = `${current}/page`;
            this.pageSizeSelect.appendChild(option);
        }
        this.pageSizeSelect.value = current;
    }

    private readOrientationPreference(): PdfOrientationPreference {
        const raw = String(this.formattingSettings?.exportCard?.orientation?.value ?? "");
        if (raw === "portrait" || raw === "landscape" || raw === "auto") return raw;
        return "auto";
    }

    private readPaperSize(): PdfPaperSize {
        const raw = String(this.formattingSettings?.exportCard?.paperSize?.value ?? "");
        if (raw === "letter") return "letter";
        if (raw === "a3") return "a3";
        return "a4";
    }

    private readPdfFontSize(): number {
        const candidate = Number(this.formattingSettings?.exportCard?.pdfFontSize?.value);
        if (!Number.isFinite(candidate)) {
            return 9;
        }
        return Math.max(7, Math.min(14, Math.floor(candidate)));
    }

    private readGeneralTitle(): { text: string; bold: boolean; fontSize: number; color: [number, number, number] } {
        const titleObj = this.dataView?.metadata?.objects?.["title"];
        const text = String(this.formattingSettings?.exportCard?.headerText?.value ?? "").trim()
            || String(titleObj?.["text"] ?? "").trim()
            || "Data Export";
        const bold = titleObj?.["bold"] !== false;
        const rawSize = Number(titleObj?.["fontSize"]);
        const fontSize = Number.isFinite(rawSize) && rawSize > 0 ? Math.min(rawSize, 24) : 10;
        const rawColor = String(titleObj?.["fontColor"]?.["solid"]?.["color"] ?? "").trim();
        const color = this.hexToRgb(rawColor) ?? [16, 42, 67];
        return { text, bold, fontSize, color };
    }

    private readGeneralSubtitle(): { text: string; show: boolean; bold: boolean; fontSize: number; color: [number, number, number] } {
        const subObj = this.dataView?.metadata?.objects?.["subTitle"];
        const text = String(this.formattingSettings?.exportCard?.subtitleText?.value ?? "").trim()
            || String(subObj?.["text"] ?? "").trim();
        const show = text.length > 0;
        const bold = !!(subObj?.["bold"]);
        const rawSize = Number(subObj?.["fontSize"]);
        const fontSize = Number.isFinite(rawSize) && rawSize > 0 ? Math.min(rawSize, 20) : 9;
        const rawColor = String(subObj?.["fontColor"]?.["solid"]?.["color"] ?? "").trim();
        const color = this.hexToRgb(rawColor) ?? [80, 80, 80];
        return { text, show: show && text.length > 0, bold, fontSize, color };
    }

    private hexToRgb(hex: string): [number, number, number] | null {
        const m = /^#?([0-9a-f]{2})([0-9a-f]{2})([0-9a-f]{2})$/i.exec(hex);
        if (!m) return null;
        return [parseInt(m[1], 16), parseInt(m[2], 16), parseInt(m[3], 16)];
    }

    private readHeaderText(): string {
        return this.readGeneralTitle().text;
    }

    private buildExportFileName(): string {
        const header = this.readHeaderText();
        const fallback = "datatable-export";
        const source = header || fallback;
        const sanitized = source
            .replace(/[\\/:*?"<>|]+/g, " ")
            .replace(/\s+/g, "-")
            .replace(/-+/g, "-")
            .replace(/^-|-$/g, "")
            .toLowerCase();
        const base = sanitized || fallback;
        return `${base}.pdf`;
    }

    private makeButton(label: string, onClick: () => void): HTMLButtonElement {
        const button = document.createElement("button");
        button.type = "button";
        button.textContent = label;
        button.addEventListener("click", onClick);
        return button;
    }

    private clearElement(element: HTMLElement): void {
        while (element.firstChild) {
            element.removeChild(element.firstChild);
        }
    }

    private setStatus(message: string, isError: boolean = false): void {
        this.statusBar.textContent = message;
        this.statusBar.className = isError ? "status-bar error" : "status-bar";
    }

    private privilegeToText(status: PrivilegeStatus): string {
        if (status === PrivilegeStatus.NotDeclared) {
            return "Privilege not declared in capabilities";
        }
        if (status === PrivilegeStatus.NotSupported) {
            return "Not supported by current Power BI host";
        }
        if (status === PrivilegeStatus.DisabledByAdmin) {
            return "Disabled by tenant admin";
        }
        return "Unknown privilege status";
    }

    private valueToText(value: unknown): string {
        if (value === null || value === undefined) {
            return "";
        }
        if (value instanceof Date) {
            return value.toISOString();
        }
        return String(value);
    }
}
