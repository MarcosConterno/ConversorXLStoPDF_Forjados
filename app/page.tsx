"use client";

import { useCallback, useMemo, useRef, useState } from "react";
import * as XLSX from "xlsx";
import jsPDF from "jspdf";
// opcional: só mantenha se for usar de fato
// import autoTable from "jspdf-autotable";

// ====== TEMA (preto/vermelho/branco) ======
const BRAND = {
  black: [20, 20, 20] as [number, number, number],
  red: [220, 38, 38] as [number, number, number],       // red-600
  redSoft: [248, 113, 113] as [number, number, number], // red-400
  white: [255, 255, 255] as [number, number, number],
  grayLine: [230, 230, 230] as [number, number, number],
};

// ====== IMAGENS ======
const LOGO_PATH = "/logo.jpg";      // sua logo (em /public)
const WATERMARK_PATH = "/logo.jpg"; // opcional (pode usar a mesma)

// ====== TÍTULO DO PDF ======
const HEADER_TITLE = "FICHA DE INSCRIÇÃO FORJADOS MC";

// ====== REMOÇÃO DE "Carimbo de data e hora" ======
function isTimestampHeader(h?: string) {
  if (!h) return false;
  const s = String(h).toLowerCase().trim();
  return /carimbo.*data.*hora/.test(s) || /timestamp/.test(s) || /data.*hora.*carimbo/.test(s);
}

// ====== HELPERS: datas ======
function normHeader(s: string) {
  return s
    .normalize("NFD")
    .replace(/\p{Diacritic}/gu, "")
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();
}
function isDateHeaderName(h: string) {
  const t = normHeader(h);
  return t.includes("data") || t.includes("nascimento") || t.includes("dt nasc") || /^dob$/.test(t);
}
function excelSerialToUTCDate(n: number): Date {
  const base = Date.UTC(1899, 11, 30); // 1899-12-30
  const ms = Math.round(n * 86400000);
  return new Date(base + ms);
}
function formatDateDDMMYYYY(d: Date): string {
  const dd = String(d.getUTCDate()).padStart(2, "0");
  const mm = String(d.getUTCMonth() + 1).padStart(2, "0");
  const yyyy = String(d.getUTCFullYear());
  return `${dd}/${mm}/${yyyy}`;
}
function normalizeDateCell(v: unknown): string {
  if (v == null) return "";

  if (v instanceof Date && !isNaN(v.getTime())) {
    return formatDateDDMMYYYY(v);
  }
  if (typeof v === "number" && isFinite(v)) {
    return formatDateDDMMYYYY(excelSerialToUTCDate(v));
  }

  let s = String(v).trim();
  if (!s) return "";
  s = s.split(/[ T]/)[0]; // remove horário

  // yyyy-mm-dd ou yyyy/mm/dd
  let m = s.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})$/);
  if (m) {
    const yyyy = m[1];
    const mm = m[2].padStart(2, "0");
    const dd = m[3].padStart(2, "0");
    return `${dd}/${mm}/${yyyy}`;
  }

  // dd/mm/aa(aa) OU mm/dd/aa(aa)
  m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2}|\d{4})$/);
  if (m) {
    let a = parseInt(m[1], 10); // primeiro número
    let b = parseInt(m[2], 10); // segundo número
    let yyyy =
      m[3].length === 2
        ? (parseInt(m[3], 10) >= 50 ? 1900 + parseInt(m[3], 10) : 2000 + parseInt(m[3], 10))
        : parseInt(m[3], 10);

    // Regra para decidir BR x US
    let dd: number, mm: number;
    if (b > 12 && a <= 12) {
      // US mm/dd -> inverter
      dd = b;
      mm = a;
    } else if (a > 12 && b <= 12) {
      dd = a;
      mm = b;
    } else if (a <= 12 && b <= 12) {
      // ambíguo: adota BR (dd/mm)
      dd = a;
      mm = b;
    } else {
      return s;
    }
    return `${String(dd).padStart(2, "0")}/${String(mm).padStart(2, "0")}/${String(yyyy)}`;
  }

  return s;
}

// ====== CARREGAR IMAGEM -> PNG base64 ======
async function loadImageAsPngDataUrl(path: string): Promise<{ dataUrl: string; w: number; h: number } | null> {
  try {
    const res = await fetch(path, { cache: "force-cache" });
    if (!res.ok) throw new Error(`HTTP ${res.status} ao buscar ${path}`);
    const blob = await res.blob();

    const dataUrl = await new Promise<string>((resolve, reject) => {
      const reader = new FileReader();
      reader.onerror = reject;
      reader.onload = () => resolve(reader.result as string);
      reader.readAsDataURL(blob);
    });

    const imgEl = await new Promise<HTMLImageElement>((resolve, reject) => {
      const i = new Image();
      i.onload = () => resolve(i);
      i.onerror = reject;
      i.src = dataUrl;
    });

    const canvas = document.createElement("canvas");
    canvas.width = Math.max(1, imgEl.naturalWidth);
    canvas.height = Math.max(1, imgEl.naturalHeight);
    const ctx = canvas.getContext("2d");
    if (!ctx) throw new Error("Canvas 2D context not available");
    ctx.drawImage(imgEl, 0, 0);
    const pngDataUrl = canvas.toDataURL("image/png");
    return { dataUrl: pngDataUrl, w: imgEl.naturalWidth, h: imgEl.naturalHeight };
  } catch (err) {
    console.error("Falha ao carregar imagem:", path, err);
    return null;
  }
}

export default function ExcelToPdfPage() {
  const [rows, setRows] = useState<any[]>([]);
  const [headers, setHeaders] = useState<string[]>([]);
  const [fileName, setFileName] = useState<string>("");
  const [sheetName, setSheetName] = useState<string>("");
  const [logo, setLogo] = useState<{ dataUrl: string; w: number; h: number } | null>(null);
  const [watermark, setWatermark] = useState<{ dataUrl: string; w: number; h: number } | null>(null);

  const inputRef = useRef<HTMLInputElement | null>(null);

  // ===== Dados limpos (remove "Carimbo de data e hora" + normaliza datas)
  const cleanHeaders = useMemo(() => headers.filter((h) => !isTimestampHeader(h)), [headers]);
  const sanitizedRows = useMemo(() => {
    if (!rows.length) return [];
    return rows.map((r) => {
      const out: Record<string, any> = {};
      for (const h of cleanHeaders) {
        const val = (r as any)[h];
        out[h] = isDateHeaderName(h) ? normalizeDateCell(val) : val;
      }
      return out;
    });
  }, [rows, cleanHeaders]);

  const hasCleanData = sanitizedRows.length > 0 && cleanHeaders.length > 0;

  // ===== Ler Excel
  const parseExcel = useCallback((file: File) => {
    setFileName(file.name);
    const reader = new FileReader();
    reader.onload = async (e) => {
      const data = new Uint8Array(e.target?.result as ArrayBuffer);

      const wb = XLSX.read(data, {
        type: "array",
        cellDates: true, // deixa o XLSX nos dar Date quando possível
        cellNF: true,
        cellText: false,
      });

      const wsName = wb.SheetNames[0];
      const ws = wb.Sheets[wsName];

      // traz valores crus; normalizamos datas depois
      const json = XLSX.utils.sheet_to_json(ws, { defval: "", raw: true });
      const hdrs = (XLSX.utils.sheet_to_json<string[]>(ws, { header: 1 })[0] || []) as string[];

      setSheetName(wsName);
      setHeaders(hdrs.map((h) => String(h)));
      setRows(json as any[]);

      // Carrega imagens (logo + watermark) em paralelo
      const [logoImg, wmkImg] = await Promise.all([
        loadImageAsPngDataUrl(LOGO_PATH),
        loadImageAsPngDataUrl(WATERMARK_PATH).catch(() => null),
      ]);
      setLogo(logoImg);
      setWatermark(wmkImg || null);
    };
    reader.readAsArrayBuffer(file);
  }, []);

  const onDrop = useCallback((e: React.DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    const file = e.dataTransfer.files?.[0];
    if (file) parseExcel(file);
  }, [parseExcel]);

  const onFileChange = useCallback((e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) parseExcel(file);
  }, [parseExcel]);

  // ===== Preview elegante
  const tablePreview = useMemo(() => {
    if (!hasCleanData) return null;
    const previewCount = Math.min(sanitizedRows.length, 1000);
    return (
      <div className="mt-10 rounded-3xl border border-gray-200 bg-white shadow-sm">
        <div className="flex flex-col items-center gap-1 px-6 py-6 text-center">
          <h3 className="text-lg font-semibold text-gray-900">Pré-visualização dos dados</h3>
          <p className="text-sm text-gray-500">
            <span className="font-medium text-gray-700">{sheetName}</span> •{" "}
            {sanitizedRows.length.toLocaleString()} linhas • {cleanHeaders.length} colunas
          </p>
          <span className="text-xs text-gray-400">Mostrando {previewCount.toLocaleString()} linhas</span>
        </div>
        <div className="overflow-auto max-h-[60vh]">
          <table className="min-w-full table-fixed text-sm">
            <thead className="sticky top-0 bg-black/90 text-white">
              <tr>
                {cleanHeaders.map((h) => (
                  <th key={h} className="px-4 py-3 text-left font-semibold">
                    {h || "(vazio)"}
                  </th>
                ))}
              </tr>
            </thead>
            <tbody className="[&>tr:nth-child(odd)]:bg-white [&>tr:nth-child(even)]:bg-gray-50">
              {sanitizedRows.slice(0, previewCount).map((r, i) => (
                <tr key={i} className="border-b border-gray-100">
                  {cleanHeaders.map((h) => (
                    <td key={h} className="px-4 py-2 align-top text-gray-800">
                      {String((r as any)[h] ?? "")}
                    </td>
                  ))}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    );
  }, [hasCleanData, cleanHeaders, sanitizedRows, sheetName]);

  // ===== PDF: 1 página por linha (centralizado, separadores, watermark e assinatura)
  function renderSingleRowPage(
    doc: jsPDF,
    row: any,
    opts: { pageW: number; pageH: number; marginX: number; bottomMargin: number }
  ) {
    const { pageW, pageH, marginX, bottomMargin } = opts;

    // Cabeçalho
    const headerH = 70;
    const logoSize = 40;

    doc.setFillColor(...BRAND.black);
    doc.rect(0, 0, pageW, headerH, "F");

    if (logo?.dataUrl) {
      try {
        doc.addImage(logo.dataUrl, "PNG", marginX, (headerH - logoSize) / 2, logoSize, logoSize, undefined, "FAST");
      } catch {}
    }

    doc.setFont("helvetica", "bold");
    doc.setFontSize(15);
    doc.setTextColor(...BRAND.white);
    doc.text(HEADER_TITLE, pageW / 2, headerH / 2 + 5, { align: "center" });

    doc.setFillColor(...BRAND.red);
    doc.rect(0, headerH, pageW, 4, "F");

    // Marca d'água (por baixo)
    if (watermark?.dataUrl) {
      const ratio = watermark.w / watermark.h;
      let w = pageW * 0.6;
      let h = w / ratio;
      if (h > pageH * 0.6) {
        h = pageH * 0.6;
        w = h * ratio;
      }
      const x = (pageW - w) / 2;
      const y = (pageH - h) / 2;
      const anyDoc: any = doc as any;
      if (anyDoc.setGState && anyDoc.GState) {
        const gs = new anyDoc.GState({ opacity: 0.07 });
        anyDoc.setGState(gs);
        doc.addImage(watermark.dataUrl, "PNG", x, y, w, h);
        const gsReset = new anyDoc.GState({ opacity: 1 });
        anyDoc.setGState(gsReset);
      } else {
        doc.addImage(watermark.dataUrl, "PNG", x, y, w, h);
      }
    }

    // ===== Pares (normais x emergência) — nomes únicos e tipagem de tupla
    const pairsAll: [string, string][] = cleanHeaders
      .map((h): [string, string] => {
        const key = String(h ?? "");
        const val = String((row as Record<string, unknown>)[key] ?? "").trim();
        return [key, val];
      })
      .filter(([, v]) => v !== "");

    const pairsNormal: [string, string][] = pairsAll.filter(([k]) => !/emerg/i.test(k));
    const pairsEmergency: [string, string][] = pairsAll.filter(([k]) => /emerg/i.test(k));

    // Helpers
    const labelColW = 200;
    const valueColW = pageW - marginX * 2 - labelColW;
    const sepColor = BRAND.grayLine;

    const renderKV = (label: string, value: string, y: number, fs = 11) => {
      doc.setFontSize(fs);
      const labelLines = doc.splitTextToSize(`${label}:`, labelColW);
      const valueLines = doc.splitTextToSize(value, valueColW);

      doc.setFont("helvetica", "bold");
      doc.setTextColor(30);
      doc.text(labelLines, marginX, y);

      doc.setFont("helvetica", "normal");
      doc.setTextColor(20);
      doc.text(valueLines, marginX + labelColW, y);

      const lineGap = fs + 8;
      const linesUsed = Math.max(labelLines.length, valueLines.length);
      const nextY = y + lineGap * linesUsed;

      // linha separadora
      doc.setDrawColor(...(sepColor as any));
      doc.setLineWidth(0.3);
      doc.line(marginX, nextY + 2, pageW - marginX, nextY + 2);

      return nextY + 6;
    };

    // Calcular altura total p/ centralizar
    const fakeDoc = new jsPDF({ unit: "pt", format: "a4" });
    let fakeY = 0;
    for (const [label, value] of pairsNormal) {
      const lbl = fakeDoc.splitTextToSize(`${label}:`, labelColW);
      const val = fakeDoc.splitTextToSize(value, valueColW);
      const lineGap = 11 + 8;
      fakeY += lineGap * Math.max(lbl.length, val.length) + 6;
    }
    if (pairsEmergency.length > 0) {
      fakeY += 40;
      for (const [label, value] of pairsEmergency) {
        const lbl = fakeDoc.splitTextToSize(`${label}:`, labelColW);
        const val = fakeDoc.splitTextToSize(value, valueColW);
        const lineGap = 11 + 8;
        fakeY += lineGap * Math.max(lbl.length, val.length) + 6;
      }
    }
    const headerHOffset = 40;
    const availableH = pageH - headerH - bottomMargin - (headerHOffset + 140);
    const startOffset = Math.max((availableH - fakeY) / 2, 0);

    // Render
    let y = headerH + headerHOffset + startOffset;
    for (const [label, value] of pairsNormal) {
      y = renderKV(label, value, y, 11);
    }

    if (pairsEmergency.length > 0) {
      y += 18;
      doc.setFont("helvetica", "bold");
      doc.setFontSize(12);
      doc.setTextColor(...BRAND.red);
      doc.text("CONTATOS DE EMERGÊNCIA", marginX, y);
      y += 14;
      for (const [label, value] of pairsEmergency) {
        y = renderKV(label, value, y, 11);
      }
    }

    // Assinatura
    const sigY = pageH - bottomMargin - 60;
    const center = pageW / 2;
    const lineW = 280;
    doc.setDrawColor(...BRAND.red);
    doc.setLineWidth(1.3);
    doc.line(center - lineW / 2, sigY, center + lineW / 2, sigY);
    doc.setFont("helvetica", "normal");
    doc.setFontSize(11);
    doc.setTextColor(60);
    doc.text("Assinatura Membro", center, sigY + 14, { align: "center" });

    // Rodapé
    doc.setTextColor(120);
    doc.setFontSize(9);
    doc.text(`Página ${doc.getCurrentPageInfo().pageNumber}`, pageW - marginX, pageH - 20, { align: "right" });
  }

  // ===== Exportar PDF
  const handleExportPdf = useCallback(async () => {
    if (!hasCleanData) return;

    // garante imagens (caso usuário exporte muito rápido)
    if (!logo) {
      const l = await loadImageAsPngDataUrl(LOGO_PATH);
      if (l) setLogo(l);
    }
    if (!watermark) {
      const w = await loadImageAsPngDataUrl(WATERMARK_PATH);
      if (w) setWatermark(w);
    }

    const doc = new jsPDF({ orientation: "portrait", unit: "pt", format: "a4" });
    const pageW = doc.internal.pageSize.getWidth();
    const pageH = doc.internal.pageSize.getHeight();
    const marginX = 44;
    const bottomMargin = 32;

    sanitizedRows.forEach((row, idx) => {
      if (idx > 0) doc.addPage();
      renderSingleRowPage(doc, row, { pageW, pageH, marginX, bottomMargin });
    });

    const pdfName = (fileName?.replace(/\.[^.]+$/, "") || "relatorio") + ".pdf";
    doc.save(pdfName);
  }, [hasCleanData, sanitizedRows, fileName, logo, watermark]);

  // ===== UI (mais bonita + centralizada)
  return (
    <main className="min-h-screen bg-gradient-to-b from-white to-gray-50">
      <div className="mx-auto w-full max-w-6xl px-4 py-10">
        {/* HERO */}
        <section className="relative overflow-hidden rounded-3xl border border-gray-200 bg-white shadow-sm">
          <div className="absolute inset-0 opacity-5 bg-[radial-gradient(1200px_400px_at_80%_-10%,#ef4444,transparent_60%)]" />
          <div className="relative flex flex-col items-center gap-3 p-10 text-center">
            <h1 className="text-3xl font-semibold tracking-tight text-gray-900">
              Ficha de Inscrição Forjados MC
            </h1>
            <p className="max-w-2xl text-sm text-gray-600">
              Envie um arquivo Excel, veja os dados e gere um PDF estilizado — 1 página por linha, com cabeçalho, marca d’água e assinatura.
            </p>
            <div className="mt-2 h-1 w-full max-w-md rounded-full bg-gradient-to-r from-black to-red-600" />
          </div>
        </section>

        {/* Uploader */}
        <section className="mt-8">
          <div
            onDrop={onDrop}
            onDragOver={(e) => e.preventDefault()}
            className="group relative flex flex-col items-center justify-center rounded-3xl border-2 border-dashed border-gray-300 bg-white p-10 text-center shadow-sm transition hover:border-red-400"
          >
            <input ref={inputRef} type="file" accept=".xlsx,.xls" className="hidden" onChange={onFileChange} />
            <div className="pointer-events-none select-none">
              <div className="mx-auto mb-4 flex h-16 w-16 items-center justify-center rounded-2xl border border-gray-200 bg-gray-50">
                <div className="h-8 w-8 rounded-full bg-red-600" />
              </div>
              <p className="mt-1 text-xs text-gray-500">
                Formatos: .xlsx, .xls • Tamanho recomendado &lt; 10MB
              </p>
            </div>
          </div>

          {/* Actions */}
          <div className="mt-6 flex flex-col items-center gap-3">
            <div className="text-sm text-gray-700">
              {fileName ? (
                <span>
                  <span className="font-medium text-gray-900">Arquivo:</span> {fileName}
                </span>
              ) : (
                <span className="text-gray-400">Nenhum arquivo selecionado</span>
              )}
            </div>

            <div className="flex flex-wrap justify-center gap-3">
              <button
                type="button"
                disabled={!hasCleanData}
                onClick={handleExportPdf}
                className="inline-flex items-center gap-2 rounded-2xl bg-black px-6 py-3 text-sm font-semibold text-white shadow-sm transition hover:bg-red-600 disabled:cursor-not-allowed disabled:opacity-40"
              >
                Exportar PDF (1 pág/linha)
              </button>
              <button
                type="button"
                onClick={() => {
                  setRows([]); setHeaders([]); setFileName(""); setSheetName("");
                  setLogo(null); setWatermark(null);
                  if (inputRef.current) (inputRef.current as any).value = "";
                }}
                className="inline-flex items-center gap-2 rounded-2xl border border-gray-300 bg-white px-6 py-3 text-sm font-semibold text-gray-900 shadow-sm transition hover:bg-gray-50"
              >
                Limpar
              </button>
            </div>
          </div>
        </section>

        {/* Preview */}
        {tablePreview}
      </div>
    </main>
  );
}
