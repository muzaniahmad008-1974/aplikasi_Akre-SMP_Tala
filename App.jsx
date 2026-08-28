import React, { useState, useEffect, useCallback, useRef, useMemo } from 'react';
import {
  LayoutDashboard, Building2, BarChart3, ShieldCheck, ShieldAlert, ChevronDown, ChevronRight,
  Search, Save, Check, X, AlertCircle, FolderOpen, Info, Lock, Unlock, Filter,
  ClipboardList, Printer, RotateCcw, ArrowLeft, GraduationCap, UserCog, Users2, Plus, Trash2, Download, FileText, LogOut
} from 'lucide-react';
import { BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer, Legend, Cell } from 'recharts';
import INDICATOR_DATA from './data/indicators.json';
import {
  Document as DocxDocument, Packer, Paragraph, TextRun, Table as DocxTable, TableRow as DocxTableRow,
  TableCell as DocxTableCell, WidthType, BorderStyle, AlignmentType, ShadingType, VerticalAlign,
} from 'docx';
import { jsPDF } from 'jspdf';
import autoTable from 'jspdf-autotable';

/* =========================================================================
   DATA INDIKATOR — 59 indikator (19 Guru + 21 Kepala Sekolah + 19 Iklim)
   Sumber: Bedah Indikator Rubrik Instrumen Akreditasi Sekolah 2026 — dimuat
   dari ./data/indicators.json (lihat file tersebut untuk detail lengkap)
   ========================================================================= */

const BAGIAN_META = {
  kp: { key: 'kp', label: 'Guru', fullLabel: 'Bagian I — Kinerja Pendidik', peran: 'Setiap Guru Mata Pelajaran', maks: 76, icon: 'GraduationCap', color: '2E5395' },
  kk: { key: 'kk', label: 'Kepala Sekolah', fullLabel: 'Bagian II — Kepemimpinan Kepala Sekolah', peran: 'Kepala Sekolah', maks: 84, icon: 'UserCog', color: '0F6B5C' },
  il: { key: 'il', label: 'Iklim Lingkungan', fullLabel: 'Bagian III — Iklim Lingkungan Belajar', peran: 'Kepala Sekolah bersama Tim Sekolah', maks: 76, icon: 'Users2', color: 'B5651D' },
};
const TOTAL_MAKS = 236;

const SCHOOL_SEED = [
  { id: 's1', name: 'SMPN 1 Jorong', npsn: '', kepsek: '', code: '1001' },
  { id: 's2', name: 'SMPN 1 Kintap', npsn: '', kepsek: '', code: '1002' },
  { id: 's3', name: 'SMPN 3 Panyipatan', npsn: '', kepsek: '', code: '1003' },
  { id: 's4', name: 'SMPN 1 Tambang Ulang', npsn: '', kepsek: '', code: '1004' },
  { id: 's5', name: 'SMPN 4 Bajuin', npsn: '', kepsek: '', code: '1005' },
  { id: 's6', name: 'SMPN 4 Pelaihari', npsn: '', kepsek: '', code: '1006' },
  { id: 's7', name: 'SMPN 5 Pelaihari', npsn: '', kepsek: '', code: '1007' },
  { id: 's8', name: 'SMPS IT Sirajul Huda', npsn: '', kepsek: '', code: '1008' },
  { id: 's9', name: 'SMPS Muhammadiyah', npsn: '', kepsek: '', code: '1009' },
  { id: 's10', name: 'SMP Tahfizh Bilingual Daarul Qur\u2019an Istiqomah', npsn: '', kepsek: '', code: '1010' },
];

// Bagian yang boleh diakses tiap peran di sekolah yang terkunci untuknya
const ROLE_TABS = { guru: ['kp'], kepsek: ['kk', 'il'], pengawas: ['kp', 'kk', 'il'] };
const ROLE_LABEL = { guru: 'Guru', kepsek: 'Kepala Sekolah', pengawas: 'Pengawas' };

const CATEGORY = (pct) => {
  if (pct === null || pct === undefined || isNaN(pct)) return { label: '\u2014', color: '9CA3AF' };
  if (pct >= 0.85) return { label: 'Sangat Baik', color: '0F6B5C' };
  if (pct >= 0.70) return { label: 'Baik', color: '2E5395' };
  if (pct >= 0.55) return { label: 'Cukup', color: 'B5651D' };
  return { label: 'Kurang', color: 'B0413E' };
};

function emptyEvidence() {
  const ev = {};
  for (const k of ['kp', 'kk', 'il']) {
    INDICATOR_DATA[k].butir.forEach(b => b.items.forEach(it => {
      ev[it.code] = { skor: 0, buktiTersedia: '', buktiPerlu: '', lokasi: '', catatan: '', updatedAt: null };
    }));
  }
  return ev;
}

function calcBagianSkor(evidence, key) {
  let total = 0;
  INDICATOR_DATA[key].butir.forEach(b => b.items.forEach(it => {
    total += (evidence[it.code]?.skor || 0);
  }));
  return total;
}
function calcBagianFilled(evidence, key) {
  let filled = 0, count = 0;
  INDICATOR_DATA[key].butir.forEach(b => b.items.forEach(it => {
    count++;
    if ((evidence[it.code]?.skor || 0) > 0) filled++;
  }));
  return { filled, count };
}
function calcTotalSkor(evidence) {
  return calcBagianSkor(evidence, 'kp') + calcBagianSkor(evidence, 'kk') + calcBagianSkor(evidence, 'il');
}
function calcTotalFilled(evidence) {
  let filled = 0, count = 0;
  for (const k of ['kp', 'kk', 'il']) {
    const r = calcBagianFilled(evidence, k);
    filled += r.filled; count += r.count;
  }
  return { filled, count };
}

/* =========================================================================
   LAPORAN REKAP PER SEKOLAH — dibuat sebagai file HTML yang bisa diunduh
   dan dicetak/disimpan sebagai PDF lewat dialog cetak browser (Ctrl/Cmd+P
   -> "Simpan sebagai PDF"), tanpa perlu pustaka PDF tambahan.
   ========================================================================= */
function escapeHtml(s) {
  return (s || '').toString()
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function buildReportHTML(school, evidence) {
  const now = new Date();
  const tanggal = now.toLocaleDateString('id-ID', { day: 'numeric', month: 'long', year: 'numeric' });
  const skorTotal = calcTotalSkor(evidence);
  const pctTotal = skorTotal / TOTAL_MAKS;
  const katTotal = CATEGORY(pctTotal);

  const ringkasanRows = ['kp', 'kk', 'il'].map(k => {
    const m = BAGIAN_META[k];
    const s = calcBagianSkor(evidence, k);
    const p = s / m.maks;
    const kat = CATEGORY(p);
    return `<tr>
      <td>${escapeHtml(m.fullLabel)}</td>
      <td style="text-align:center">${s}</td>
      <td style="text-align:center">${m.maks}</td>
      <td style="text-align:center">${Math.round(p * 100)}%</td>
      <td style="text-align:center"><span class="badge" style="background:#${kat.color}1A;color:#${kat.color}">${escapeHtml(kat.label)}</span></td>
    </tr>`;
  }).join('');

  const bagianSections = ['kp', 'kk', 'il'].map(k => {
    const meta = BAGIAN_META[k];
    const butirHtml = INDICATOR_DATA[k].butir.map(b => {
      const rows = b.items.map(it => {
        const d = evidence[it.code] || {};
        const skor = d.skor || 0;
        const skorColor = skor === 0 ? '9CA3AF' : SKOR_LABELS[skor].color;
        return `<tr>
          <td style="white-space:nowrap;font-weight:700">${escapeHtml(it.code)}</td>
          <td>${escapeHtml(it.name)}</td>
          <td style="text-align:center"><span class="badge" style="background:#${skorColor}1A;color:#${skorColor}">${skor === 0 ? 'Belum' : 'Skor ' + skor}</span></td>
          <td>${escapeHtml(d.buktiTersedia)}</td>
          <td>${escapeHtml(d.buktiPerlu)}</td>
          <td>${escapeHtml(d.lokasi)}</td>
          <td>${escapeHtml(d.catatan)}</td>
        </tr>`;
      }).join('');
      return `
        <h4>BUTIR ${escapeHtml(b.no)} — ${escapeHtml(b.rumusan)}</h4>
        <table class="detail">
          <thead><tr><th>Kode</th><th>Indikator</th><th>Skor</th><th>Bukti Tersedia</th><th>Bukti Perlu Dilengkapi</th><th>Lokasi Berkas</th><th>Catatan</th></tr></thead>
          <tbody>${rows}</tbody>
        </table>`;
    }).join('');
    const bagianSkor = calcBagianSkor(evidence, k);
    return `
      <div class="bagian-section">
        <h2 style="color:#${meta.color}">${escapeHtml(meta.fullLabel)} <span style="font-weight:400;font-size:14px;color:#555">— ${bagianSkor}/${meta.maks} (${Math.round((bagianSkor / meta.maks) * 100)}%)</span></h2>
        <p style="color:#555;font-size:13px;margin-top:-4px">Penanggung jawab: ${escapeHtml(meta.peran)}</p>
        ${butirHtml}
      </div>`;
  }).join('');

  return `<!doctype html>
<html lang="id">
<head>
<meta charset="utf-8" />
<title>Laporan Rekap - ${escapeHtml(school.name)}</title>
<style>
  * { box-sizing: border-box; }
  body { font-family: 'Segoe UI', Arial, sans-serif; color: #1C2530; margin: 0; padding: 32px; background: #fff; }
  h1 { font-size: 22px; color: #1F3864; margin-bottom: 4px; }
  h2 { font-size: 17px; margin: 28px 0 6px; }
  h4 { font-size: 13px; background: #DCE6F1; color: #1F3864; padding: 6px 10px; border-radius: 6px; margin: 16px 0 6px; }
  .header { border-bottom: 3px solid #1F3864; padding-bottom: 14px; margin-bottom: 18px; }
  .meta { color: #444; font-size: 13px; margin: 2px 0; }
  table { width: 100%; border-collapse: collapse; margin-bottom: 10px; }
  table.summary th, table.summary td { border: 1px solid #ccc; padding: 7px 10px; font-size: 13px; }
  table.summary th { background: #1F3864; color: #fff; text-align: left; }
  table.detail { font-size: 11.5px; }
  table.detail th, table.detail td { border: 1px solid #ddd; padding: 5px 7px; vertical-align: top; }
  table.detail th { background: #2E5395; color: #fff; text-align: left; }
  .badge { display: inline-block; padding: 2px 8px; border-radius: 999px; font-weight: 700; font-size: 11px; }
  .total-box { display: inline-block; background: #0F6B5C; color: #fff; padding: 10px 18px; border-radius: 10px; font-weight: 700; margin-top: 6px; }
  .footer { margin-top: 30px; font-size: 11px; color: #888; border-top: 1px solid #eee; padding-top: 10px; }
  @media print { .bagian-section { page-break-before: always; } body { padding: 12px; } }
</style>
</head>
<body>
  <div class="header">
    <h1>Laporan Rekap Bukti Dukung Akreditasi SMP 2026</h1>
    <p class="meta"><strong>${escapeHtml(school.name)}</strong></p>
    <p class="meta">NPSN: ${escapeHtml(school.npsn) || '-'} &middot; Kepala Sekolah: ${escapeHtml(school.kepsek) || '-'}</p>
    <p class="meta">Diunduh pada: ${tanggal}</p>
  </div>

  <h2>Ringkasan Skor</h2>
  <table class="summary">
    <thead><tr><th>Komponen</th><th>Skor</th><th>Maks</th><th>%</th><th>Kategori</th></tr></thead>
    <tbody>${ringkasanRows}</tbody>
  </table>
  <div class="total-box">TOTAL: ${skorTotal} / ${TOTAL_MAKS} (${Math.round(pctTotal * 100)}%) — ${escapeHtml(katTotal.label)}</div>

  ${bagianSections}

  <div class="footer">
    Dihasilkan otomatis oleh Aplikasi Bukti Dukung Akreditasi SMP 2026 &middot; Kab. Tanah Laut. Data mengikuti kondisi terakhir pada saat laporan ini diunduh.
  </div>
</body>
</html>`;
}

function triggerBlobDownload(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 2000);
}

function safeFileName(school) {
  return school.name.replace(/[^a-z0-9]+/gi, '_');
}

function downloadReportHTML(school, evidence) {
  const html = buildReportHTML(school, evidence);
  triggerBlobDownload(new Blob([html], { type: 'text/html;charset=utf-8' }), `Laporan_${safeFileName(school)}.html`);
}

/* ---------- Laporan Word (.docx) asli, memakai pustaka "docx" ---------- */
function docxCellHeader(text, width) {
  return new DocxTableCell({
    width: { size: width, type: WidthType.PERCENTAGE },
    shading: { type: ShadingType.CLEAR, fill: '2E5395' },
    verticalAlign: VerticalAlign.CENTER,
    margins: { top: 60, bottom: 60, left: 80, right: 80 },
    children: [new Paragraph({ children: [new TextRun({ text, bold: true, size: 16, color: 'FFFFFF' })] })],
  });
}
function docxCell(text, width, opts = {}) {
  return new DocxTableCell({
    width: { size: width, type: WidthType.PERCENTAGE },
    shading: opts.fill ? { type: ShadingType.CLEAR, fill: opts.fill } : undefined,
    verticalAlign: VerticalAlign.TOP,
    margins: { top: 60, bottom: 60, left: 80, right: 80 },
    children: [new Paragraph({ alignment: opts.center ? AlignmentType.CENTER : AlignmentType.LEFT, children: [new TextRun({ text: text || '', size: 16, bold: !!opts.bold })] })],
  });
}

async function buildReportDocx(school, evidence) {
  const now = new Date();
  const tanggal = now.toLocaleDateString('id-ID', { day: 'numeric', month: 'long', year: 'numeric' });
  const skorTotal = calcTotalSkor(evidence);
  const pctTotal = skorTotal / TOTAL_MAKS;
  const katTotal = CATEGORY(pctTotal);

  const children = [];
  children.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 100 }, children: [new TextRun({ text: 'LAPORAN REKAP BUKTI DUKUNG AKREDITASI SMP 2026', bold: true, size: 32, color: '1F3864' })] }));
  children.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 40 }, children: [new TextRun({ text: school.name, bold: true, size: 24 })] }));
  children.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 40 }, children: [new TextRun({ text: `NPSN: ${school.npsn || '-'}  \u00b7  Kepala Sekolah: ${school.kepsek || '-'}`, size: 20 })] }));
  children.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 240 }, children: [new TextRun({ text: `Diunduh pada: ${tanggal}`, italics: true, size: 18, color: '666666' })] }));

  children.push(new Paragraph({ spacing: { before: 100, after: 100 }, children: [new TextRun({ text: 'Ringkasan Skor', bold: true, size: 24, color: '1F3864' })] }));
  const summaryHeader = new DocxTableRow({ children: [docxCellHeader('Komponen', 40), docxCellHeader('Skor', 15), docxCellHeader('Maks', 15), docxCellHeader('%', 15), docxCellHeader('Kategori', 15)] });
  const summaryRows = ['kp', 'kk', 'il'].map((k) => {
    const m = BAGIAN_META[k];
    const s = calcBagianSkor(evidence, k);
    const p = s / m.maks;
    const kat = CATEGORY(p);
    return new DocxTableRow({ children: [docxCell(m.fullLabel, 40), docxCell(String(s), 15, { center: true }), docxCell(String(m.maks), 15, { center: true }), docxCell(Math.round(p * 100) + '%', 15, { center: true }), docxCell(kat.label, 15, { center: true, bold: true })] });
  });
  const totalRow = new DocxTableRow({ children: [docxCell('TOTAL', 40, { bold: true, fill: 'DCE6F1' }), docxCell(String(skorTotal), 15, { center: true, bold: true, fill: 'DCE6F1' }), docxCell(String(TOTAL_MAKS), 15, { center: true, bold: true, fill: 'DCE6F1' }), docxCell(Math.round(pctTotal * 100) + '%', 15, { center: true, bold: true, fill: 'DCE6F1' }), docxCell(katTotal.label, 15, { center: true, bold: true, fill: 'DCE6F1' })] });
  children.push(new DocxTable({ width: { size: 100, type: WidthType.PERCENTAGE }, rows: [summaryHeader, ...summaryRows, totalRow] }));

  ['kp', 'kk', 'il'].forEach((k) => {
    const meta = BAGIAN_META[k];
    children.push(new Paragraph({ pageBreakBefore: true, spacing: { before: 120, after: 60 }, children: [new TextRun({ text: meta.fullLabel, bold: true, size: 26, color: '1F3864' })] }));
    children.push(new Paragraph({ spacing: { after: 160 }, children: [new TextRun({ text: `Penanggung jawab: ${meta.peran}`, italics: true, size: 18, color: '555555' })] }));
    INDICATOR_DATA[k].butir.forEach((b) => {
      children.push(new Paragraph({ shading: { type: ShadingType.CLEAR, fill: 'DCE6F1' }, spacing: { before: 140, after: 80 }, children: [new TextRun({ text: `BUTIR ${b.no}  ${b.rumusan}`, bold: true, size: 19, color: '1F3864' })] }));
      const header = new DocxTableRow({ children: [docxCellHeader('Kode', 8), docxCellHeader('Indikator', 24), docxCellHeader('Skor', 8), docxCellHeader('Bukti Tersedia', 18), docxCellHeader('Bukti Perlu Dilengkapi', 18), docxCellHeader('Lokasi', 12), docxCellHeader('Catatan', 12)] });
      const rows = b.items.map((it) => {
        const d = evidence[it.code] || {};
        const skorTxt = d.skor ? 'Skor ' + d.skor : 'Belum';
        return new DocxTableRow({ children: [docxCell(it.code, 8, { bold: true }), docxCell(it.name, 24), docxCell(skorTxt, 8, { center: true }), docxCell(d.buktiTersedia, 18), docxCell(d.buktiPerlu, 18), docxCell(d.lokasi, 12), docxCell(d.catatan, 12)] });
      });
      children.push(new DocxTable({ width: { size: 100, type: WidthType.PERCENTAGE }, rows: [header, ...rows] }));
      children.push(new Paragraph({ spacing: { after: 80 }, children: [] }));
    });
  });

  const doc = new DocxDocument({
    styles: { default: { document: { run: { font: 'Calibri', size: 18 } } } },
    sections: [{ properties: { page: { margin: { top: 700, bottom: 700, left: 700, right: 700 } } }, children }],
  });
  return Packer.toBlob(doc);
}

async function downloadReportDocx(school, evidence) {
  const blob = await buildReportDocx(school, evidence);
  triggerBlobDownload(blob, `Laporan_${safeFileName(school)}.docx`);
}

/* ---------- Laporan PDF asli, memakai pustaka "jspdf" + "jspdf-autotable" ---------- */
function buildReportPdf(school, evidence) {
  const now = new Date();
  const tanggal = now.toLocaleDateString('id-ID', { day: 'numeric', month: 'long', year: 'numeric' });
  const skorTotal = calcTotalSkor(evidence);
  const pctTotal = skorTotal / TOTAL_MAKS;
  const katTotal = CATEGORY(pctTotal);

  const docPdf = new jsPDF({ orientation: 'landscape', unit: 'pt', format: 'a4' });
  const pageWidth = docPdf.internal.pageSize.getWidth();

  docPdf.setFont('helvetica', 'bold');
  docPdf.setFontSize(16);
  docPdf.setTextColor(31, 56, 100);
  docPdf.text('LAPORAN REKAP BUKTI DUKUNG AKREDITASI SMP 2026', pageWidth / 2, 40, { align: 'center' });
  docPdf.setFontSize(12);
  docPdf.text(school.name, pageWidth / 2, 60, { align: 'center' });
  docPdf.setFont('helvetica', 'normal');
  docPdf.setFontSize(9);
  docPdf.setTextColor(80, 80, 80);
  docPdf.text(`NPSN: ${school.npsn || '-'}   |   Kepala Sekolah: ${school.kepsek || '-'}   |   Diunduh: ${tanggal}`, pageWidth / 2, 76, { align: 'center' });

  const summaryBody = ['kp', 'kk', 'il'].map((k) => {
    const m = BAGIAN_META[k];
    const s = calcBagianSkor(evidence, k);
    const p = s / m.maks;
    return [m.fullLabel, String(s), String(m.maks), Math.round(p * 100) + '%', CATEGORY(p).label];
  });
  summaryBody.push(['TOTAL', String(skorTotal), String(TOTAL_MAKS), Math.round(pctTotal * 100) + '%', katTotal.label]);

  autoTable(docPdf, {
    startY: 92,
    head: [['Komponen', 'Skor', 'Maks', '%', 'Kategori']],
    body: summaryBody,
    theme: 'grid',
    headStyles: { fillColor: [31, 56, 100], textColor: 255, fontSize: 9 },
    bodyStyles: { fontSize: 9 },
    margin: { left: 40, right: 40 },
  });

  ['kp', 'kk', 'il'].forEach((k) => {
    const meta = BAGIAN_META[k];
    docPdf.addPage();
    docPdf.setFont('helvetica', 'bold');
    docPdf.setFontSize(13);
    docPdf.setTextColor(31, 56, 100);
    docPdf.text(meta.fullLabel, 40, 36);
    docPdf.setFont('helvetica', 'italic');
    docPdf.setFontSize(9);
    docPdf.setTextColor(90, 90, 90);
    docPdf.text(`Penanggung jawab: ${meta.peran}`, 40, 50);

    let startY = 62;
    INDICATOR_DATA[k].butir.forEach((b) => {
      const body = b.items.map((it) => {
        const d = evidence[it.code] || {};
        const skorTxt = d.skor ? 'Skor ' + d.skor : 'Belum';
        return [it.code, it.name, skorTxt, d.buktiTersedia || '', d.buktiPerlu || '', d.lokasi || '', d.catatan || ''];
      });
      autoTable(docPdf, {
        startY,
        head: [[{ content: `BUTIR ${b.no}  ${b.rumusan}`, colSpan: 7, styles: { fillColor: [220, 230, 241], textColor: [31, 56, 100], fontStyle: 'bold', fontSize: 8.5 } }], ['Kode', 'Indikator', 'Skor', 'Bukti Tersedia', 'Bukti Perlu Dilengkapi', 'Lokasi', 'Catatan']],
        body,
        theme: 'grid',
        headStyles: { fillColor: [46, 83, 149], textColor: 255, fontSize: 7.5 },
        bodyStyles: { fontSize: 7, cellPadding: 3 },
        columnStyles: { 0: { cellWidth: 35 }, 2: { cellWidth: 30 } },
        margin: { left: 40, right: 40 },
        didParseCell: (data) => {
          if (data.row.index === 0 && data.section === 'head' && data.table.head.length > 1) {
            // baris judul butir yang di-colspan sudah diberi style lewat konten di atas
          }
        },
      });
      startY = docPdf.lastAutoTable.finalY + 14;
    });
  });

  return docPdf;
}

function downloadReportPdf(school, evidence) {
  const docPdf = buildReportPdf(school, evidence);
  docPdf.save(`Laporan_${safeFileName(school)}.pdf`);
}

/* =========================================================================
   PENYIMPANAN — panggilan ke Netlify Functions (/api/schools, /api/evidence)
   yang disokong oleh Netlify Blobs sebagai database sesungguhnya.
   ========================================================================= */
async function apiGetSchools() {
  const res = await fetch('/api/schools');
  if (!res.ok) {
    const data = await res.json().catch(() => ({}));
    const err = new Error(data.error || 'Gagal memuat data sekolah');
    err.code = data.code;
    throw err;
  }
  return res.json();
}
async function apiUpdateSchool(id, patch, adminCode) {
  const res = await fetch('/api/schools', {
    method: 'PUT',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ id, patch, adminCode }),
  });
  const data = await res.json().catch(() => ({}));
  if (!res.ok) throw new Error(data.error || 'Gagal menyimpan data sekolah');
  return data;
}
async function apiAddSchool(newSchool, adminCode) {
  const res = await fetch('/api/schools', {
    method: 'PUT',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ newSchool, adminCode }),
  });
  const data = await res.json().catch(() => ({}));
  if (!res.ok) throw new Error(data.error || 'Gagal menambah sekolah');
  return data;
}
async function apiDeleteSchool(id, adminCode) {
  const res = await fetch('/api/schools', {
    method: 'DELETE',
    headers: { 'Content-Type': 'application/json', 'x-admin-code': adminCode },
    body: JSON.stringify({ id, adminCode }),
  });
  const data = await res.json().catch(() => ({}));
  if (!res.ok) throw new Error(data.error || 'Gagal menghapus sekolah');
  return data;
}
async function apiVerifyAdmin(adminCode) {
  const res = await fetch('/api/schools', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ adminCode: (adminCode || '').trim() }),
  });
  const data = await res.json().catch(() => ({ ok: false }));
  return {
    ok: !!data.ok,
    usingDefault: !!data.usingDefault,
    providedLength: data.providedLength,
    expectedLength: data.expectedLength,
  };
}
async function apiGetEvidence(schoolId) {
  const res = await fetch('/api/evidence?schoolId=' + encodeURIComponent(schoolId));
  if (!res.ok) {
    const data = await res.json().catch(() => ({}));
    const err = new Error(data.error || 'Gagal memuat bukti dukung');
    err.code = data.code;
    throw err;
  }
  return res.json();
}
async function apiSaveEvidence(schoolId, evidence) {
  const res = await fetch('/api/evidence', {
    method: 'PUT',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ schoolId, evidence }),
  });
  return res.ok;
}
async function apiResetEvidence(schoolId, adminCode) {
  const res = await fetch('/api/evidence?schoolId=' + encodeURIComponent(schoolId), {
    method: 'DELETE',
    headers: { 'Content-Type': 'application/json', 'x-admin-code': adminCode },
    body: JSON.stringify({ schoolId, adminCode }),
  });
  const data = await res.json().catch(() => ({}));
  if (!res.ok) throw new Error(data.error || 'Gagal mereset data');
  return data;
}

/* =========================================================================
   KOMPONEN KECIL
   ========================================================================= */
function ProgressRing({ pct, size = 64, stroke = 9, color = '1F3864', label }) {
  const r = (size - stroke) / 2;
  const c = 2 * Math.PI * r;
  const safePct = isNaN(pct) ? 0 : Math.max(0, Math.min(1, pct));
  const dash = c * safePct;
  return (
    <div className="relative inline-flex items-center justify-center" style={{ width: size, height: size }}>
      <svg width={size} height={size} className="-rotate-90">
        <circle cx={size / 2} cy={size / 2} r={r} fill="none" stroke="#E5E9F0" strokeWidth={stroke} />
        <circle
          cx={size / 2} cy={size / 2} r={r} fill="none" stroke={`#${color}`} strokeWidth={stroke}
          strokeDasharray={`${dash} ${c - dash}`} strokeLinecap="round"
          style={{ transition: 'stroke-dasharray 0.6s ease' }}
        />
      </svg>
      <div className="absolute inset-0 flex flex-col items-center justify-center">
        <span className="font-serif font-extrabold text-[#1C2530]" style={{ fontSize: size * 0.32 }}>{Math.round(safePct * 100)}</span>
        {label && <span className="text-[11px] text-slate-600 -mt-0.5">{label}</span>}
      </div>
    </div>
  );
}

function Modal({ open, onClose, title, children, footer }) {
  if (!open) return null;
  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/40 backdrop-blur-sm" onClick={onClose}>
      <div className="bg-white rounded-2xl shadow-2xl w-full max-w-md overflow-hidden" onClick={e => e.stopPropagation()}>
        <div className="flex items-center justify-between px-5 py-4 border-b border-slate-300">
          <h3 className="font-bold text-[#1F3864]">{title}</h3>
          <button onClick={onClose} className="text-slate-600 hover:text-slate-800 p-1 rounded-full hover:bg-slate-100">
            <X size={18} />
          </button>
        </div>
        <div className="px-5 py-4">{children}</div>
        {footer && <div className="px-5 py-4 border-t border-slate-300 flex justify-end gap-2">{footer}</div>}
      </div>
    </div>
  );
}

function Badge({ children, color = '1F3864', soft = true }) {
  return (
    <span
      className="inline-flex items-center px-2 py-0.5 rounded-full text-sm font-bold"
      style={soft ? { backgroundColor: `#${color}1A`, color: `#${color}` } : { backgroundColor: `#${color}`, color: '#fff' }}
    >
      {children}
    </span>
  );
}

/* =========================================================================
   GERBANG PERAN — Guru & Kepala Sekolah hanya bisa membuka sekolahnya sendiri
   ========================================================================= */
function RoleGateScreen({ schools, onLogin, error, checking }) {
  const [step, setStep] = useState('pilih'); // 'pilih' | 'guru' | 'kepsek' | 'pengawas'
  const [schoolId, setSchoolId] = useState('');
  const [code, setCode] = useState('');

  const roleCard = (role, title, desc, Icon) => (
    <button
      onClick={() => setStep(role)}
      className="w-full flex items-center gap-3.5 rounded-2xl border-2 border-slate-300 bg-white p-4 text-left hover:border-[#1F3864] hover:shadow-md transition-all"
    >
      <div className="w-11 h-11 rounded-xl bg-[#1F3864]/10 flex items-center justify-center shrink-0">
        <Icon size={22} className="text-[#1F3864]" />
      </div>
      <div className="min-w-0">
        <p className="font-bold text-base text-slate-900">{title}</p>
        <p className="text-sm text-slate-700 font-semibold">{desc}</p>
      </div>
      <ChevronRight size={18} className="text-slate-500 ml-auto shrink-0" />
    </button>
  );

  const submit = () => {
    if (step === 'pengawas') { if (code) onLogin('pengawas', null, code.trim()); return; }
    if (schoolId && code) onLogin(step, schoolId, code.trim());
  };

  return (
    <div className="min-h-screen bg-[#6FA8DC] flex items-center justify-center p-4">
      <div className="w-full max-w-md">
        <div className="text-center mb-6">
          <div className="w-14 h-14 rounded-2xl bg-[#1F3864] flex items-center justify-center mx-auto mb-3 shadow-lg">
            <ClipboardList size={26} className="text-white" />
          </div>
          <h1 className="font-serif font-extrabold text-2xl text-slate-900">Bukti Dukung Akreditasi</h1>
          <p className="text-sm text-slate-800 font-semibold">SMP 2026 &middot; Kab. Tanah Laut</p>
        </div>

        <div className="bg-white/60 backdrop-blur rounded-3xl p-4">
          {step === 'pilih' && (
            <div className="space-y-2.5">
              <p className="text-center text-sm font-bold text-slate-900 mb-3">Masuk sebagai:</p>
              {roleCard('guru', 'Guru Mata Pelajaran', 'Isi bukti dukung Bagian I sekolah Anda', GraduationCap)}
              {roleCard('kepsek', 'Kepala Sekolah', 'Isi bukti dukung Bagian II & III sekolah Anda', UserCog)}
              {roleCard('pengawas', 'Pengawas', 'Lihat & kelola seluruh sekolah binaan', ShieldCheck)}
            </div>
          )}

          {(step === 'guru' || step === 'kepsek' || step === 'pengawas') && (
            <div className="bg-white rounded-2xl p-4">
              <button onClick={() => { setStep('pilih'); setCode(''); setSchoolId(''); }} className="flex items-center gap-1 text-sm font-bold text-slate-700 mb-3">
                <ArrowLeft size={15} /> Kembali
              </button>
              <p className="font-bold text-base text-slate-900 mb-3">
                Masuk sebagai {step === 'guru' ? 'Guru' : step === 'kepsek' ? 'Kepala Sekolah' : 'Pengawas'}
              </p>
              {step !== 'pengawas' && (
                <>
                  <label className="text-xs font-bold text-slate-700 uppercase tracking-wide">Nama Sekolah</label>
                  <select
                    value={schoolId} onChange={e => setSchoolId(e.target.value)}
                    className="mt-1 w-full rounded-lg border-2 border-slate-400 px-3 py-2.5 text-base font-semibold text-slate-900 focus:outline-none focus:ring-2 focus:ring-[#1F3864]/40 mb-3"
                  >
                    <option value="" disabled>Pilih sekolah Anda...</option>
                    {schools.map(s => <option key={s.id} value={s.id}>{s.name}</option>)}
                  </select>
                </>
              )}
              <label className="text-xs font-bold text-slate-700 uppercase tracking-wide">{step === 'pengawas' ? 'Kode Akses Pengawas' : 'Kode Akses Sekolah'}</label>
              <input
                type={step === 'pengawas' ? 'password' : 'text'}
                value={code} onChange={e => setCode(e.target.value)}
                onKeyDown={e => e.key === 'Enter' && submit()}
                placeholder={step === 'pengawas' ? 'Kode akses Pengawas' : 'mis. 1001'} inputMode={step === 'pengawas' ? 'text' : 'numeric'}
                className="mt-1 w-full rounded-lg border-2 border-slate-400 px-3 py-2.5 text-base font-semibold text-slate-900 focus:outline-none focus:ring-2 focus:ring-[#1F3864]/40"
              />
              <p className="text-xs text-slate-600 font-semibold mt-1.5">
                {step === 'pengawas' ? 'Kode Pengawas Pembina (ADMIN_CODE).' : 'Minta kode akses sekolah Anda ke Pengawas Pembina.'}
              </p>
              {error && <p className="text-sm text-rose-600 font-bold mt-2 flex items-center gap-1"><AlertCircle size={14} /> {error}</p>}
              <button
                onClick={submit}
                disabled={(step !== 'pengawas' && !schoolId) || !code || checking}
                className="w-full mt-4 py-2.5 rounded-xl text-base font-bold text-white bg-[#1F3864] hover:bg-[#162a4d] disabled:opacity-50"
              >
                {checking ? 'Memeriksa...' : 'Masuk'}
              </button>
            </div>
          )}
        </div>
      </div>
    </div>
  );
}

const SKOR_LABELS = {
  0: { label: 'Belum', color: '9CA3AF', desc: 'Belum dinilai' },
  1: { label: '1', color: 'B0413E', desc: 'Belum Ada Sistem — belum dijalankan sama sekali atau tidak ada bukti pelaksanaan' },
  2: { label: '2', color: 'B5651D', desc: 'Mulai Ada tetapi Belum Konsisten — sudah ada upaya, tapi masih sporadis' },
  3: { label: '3', color: '2E5395', desc: 'Berjalan Cukup Baik dan Terencana — teratur, konsisten, mulai terdokumentasi' },
  4: { label: '4', color: '0F6B5C', desc: 'Konsisten, Sistematis, Berdampak — terdokumentasi lengkap, dievaluasi berkala' },
};

function SkorPicker({ value, onChange }) {
  return (
    <div className="flex gap-1.5">
      {[0, 1, 2, 3, 4].map(v => {
        const active = (value || 0) === v;
        const meta = SKOR_LABELS[v];
        return (
          <button
            key={v}
            onClick={() => onChange(v)}
            className={`w-11 h-11 rounded-lg text-lg font-extrabold border transition-all ${active ? 'scale-105 shadow-sm' : 'opacity-90 hover:opacity-100'}`}
            style={active
              ? { backgroundColor: `#${meta.color}`, color: '#fff', borderColor: `#${meta.color}` }
              : { backgroundColor: '#fff', color: `#${meta.color}`, borderColor: `#${meta.color}99` }}
            title={v === 0 ? 'Belum dinilai' : `Skor ${v}`}
          >
            {v === 0 ? '\u2014' : v}
          </button>
        );
      })}
    </div>
  );
}

/* =========================================================================
   KARTU INDIKATOR (input bukti dukung)
   ========================================================================= */
function IndicatorCard({ item, data, onChange }) {
  const [showInfo, setShowInfo] = useState(false);
  const skor = data?.skor || 0;
  const filled = skor > 0;

  return (
    <div className={`rounded-xl border bg-white transition-colors ${filled ? 'border-slate-400' : 'border-dashed border-slate-300'}`}>
      <div className="flex items-start gap-3 p-3.5">
        <div className="flex-1 min-w-0">
          <div className="flex items-center gap-2 flex-wrap">
            <Badge color="1F3864">{item.code}</Badge>
            {filled && <Badge color={SKOR_LABELS[skor].color}>Skor {skor}</Badge>}
            {!filled && <span className="text-sm text-amber-600 font-semibold">Belum diisi</span>}
          </div>
          <p className="text-base font-bold text-[#1C2530] mt-1 leading-snug">{item.name}</p>
          <button
            onClick={() => setShowInfo(s => !s)}
            className="mt-1.5 inline-flex items-center gap-1 text-sm text-[#0F6B5C] font-semibold hover:underline"
          >
            <Info size={13} /> {showInfo ? 'Sembunyikan panduan' : 'Lihat panduan skor & sumber bukti'}
          </button>
        </div>
      </div>

      {showInfo && (
        <div className="mx-3.5 mb-3 rounded-lg bg-[#F5F8FC] border border-[#DCE6F1] p-3 text-sm text-slate-800 space-y-2">
          <div>
            <span className="font-bold text-[#1F3864]">Panduan Skor Mandiri:</span>
            <div className="mt-1.5 space-y-1">
              {[1, 2, 3, 4].map((v) => (
                <div key={v} className="flex items-start gap-2">
                  <span
                    className="shrink-0 w-5 h-5 rounded-full flex items-center justify-center text-xs font-bold text-white"
                    style={{ backgroundColor: `#${SKOR_LABELS[v].color}` }}
                  >
                    {v}
                  </span>
                  <span className="text-sm text-slate-700">{SKOR_LABELS[v].desc}</span>
                </div>
              ))}
            </div>
          </div>
          {item.penjelasan && <p><span className="font-bold text-[#1F3864]">Penjelasan: </span>{item.penjelasan}</p>}
          {item.tindakan && item.tindakan.length > 0 && (
            <div>
              <span className="font-bold text-[#1F3864]">Yang perlu dilakukan:</span>
              <ul className="list-disc ml-4 mt-0.5 space-y-0.5">
                {item.tindakan.map((t, i) => <li key={i}>{t}</li>)}
              </ul>
            </div>
          )}
          {item.telaah && <p><span className="font-bold text-[#1F3864]">Telaah Dokumen: </span>{item.telaah}</p>}
          {item.wawancara && <p><span className="font-bold text-[#1F3864]">Wawancara: </span>{item.wawancara}</p>}
          {item.observasi && <p><span className="font-bold text-[#1F3864]">Observasi: </span>{item.observasi}</p>}
        </div>
      )}

      <div className="px-3.5 pb-3.5 space-y-2.5">
        <div>
          <label className="text-sm font-bold text-slate-700 uppercase tracking-wide">Skor Mandiri</label>
          <div className="mt-1"><SkorPicker value={skor} onChange={(v) => onChange({ skor: v })} /></div>
        </div>
        <div className="grid sm:grid-cols-2 gap-2.5">
          <div>
            <label className="text-sm font-bold text-slate-700 uppercase tracking-wide">Bukti yang Sudah Tersedia</label>
            <textarea
              value={data?.buktiTersedia || ''} onChange={e => onChange({ buktiTersedia: e.target.value })}
              rows={2} placeholder="Contoh: SK Tim TPPK, jurnal refleksi guru..."
              className="mt-1 w-full text-base rounded-lg border border-slate-400 px-2.5 py-2 focus:outline-none focus:ring-2 focus:ring-[#1F3864]/30 focus:border-[#1F3864] resize-none"
            />
          </div>
          <div>
            <label className="text-sm font-bold text-slate-700 uppercase tracking-wide">Bukti yang Perlu Dilengkapi</label>
            <textarea
              value={data?.buktiPerlu || ''} onChange={e => onChange({ buktiPerlu: e.target.value })}
              rows={2} placeholder="Contoh: belum ada dokumentasi foto kegiatan..."
              className="mt-1 w-full text-base rounded-lg border border-slate-400 px-2.5 py-2 focus:outline-none focus:ring-2 focus:ring-[#1F3864]/30 focus:border-[#1F3864] resize-none"
            />
          </div>
        </div>
        <div className="grid sm:grid-cols-2 gap-2.5">
          <div>
            <label className="text-sm font-bold text-slate-700 uppercase tracking-wide">Lokasi / Nama Berkas</label>
            <div className="mt-1 flex items-center gap-1.5 rounded-lg border border-slate-400 px-2.5 py-2 focus-within:ring-2 focus-within:ring-[#1F3864]/30 focus-within:border-[#1F3864]">
              <FolderOpen size={14} className="text-slate-600 shrink-0" />
              <input
                value={data?.lokasi || ''} onChange={e => onChange({ lokasi: e.target.value })}
                placeholder="mis. Drive/Bagian I/1.1.1.pdf"
                className="w-full text-base focus:outline-none"
              />
            </div>
          </div>
          <div>
            <label className="text-sm font-bold text-slate-700 uppercase tracking-wide">Catatan / Rencana Perbaikan</label>
            <input
              value={data?.catatan || ''} onChange={e => onChange({ catatan: e.target.value })}
              placeholder="Catatan singkat..."
              className="mt-1 w-full text-base rounded-lg border border-slate-400 px-2.5 py-2 focus:outline-none focus:ring-2 focus:ring-[#1F3864]/30 focus:border-[#1F3864]"
            />
          </div>
        </div>
      </div>
    </div>
  );
}

/* =========================================================================
   BUTIR ACCORDION
   ========================================================================= */
function ButirAccordion({ butir, evidence, onChangeItem, defaultOpen, filterMode }) {
  const [open, setOpen] = useState(!!defaultOpen);
  const items = butir.items.filter(it => {
    if (filterMode === 'belum') return (evidence[it.code]?.skor || 0) === 0;
    if (filterMode === 'ada') return (evidence[it.code]?.skor || 0) > 0;
    return true;
  });
  const filledCount = butir.items.filter(it => (evidence[it.code]?.skor || 0) > 0).length;

  if (filterMode !== 'semua' && items.length === 0) return null;

  return (
    <div className="rounded-xl border border-slate-400 overflow-hidden bg-white">
      <button onClick={() => setOpen(o => !o)} className="w-full flex items-center justify-between gap-3 px-4 py-3 bg-[#DCE6F1]/60 hover:bg-[#DCE6F1] transition-colors text-left">
        <div className="flex items-center gap-2 min-w-0">
          {open ? <ChevronDown size={16} className="text-[#1F3864] shrink-0" /> : <ChevronRight size={16} className="text-[#1F3864] shrink-0" />}
          <span className="text-sm font-bold text-[#1F3864] shrink-0">BUTIR {butir.no}</span>
          <span className="text-base text-slate-900 truncate">{butir.rumusan}</span>
        </div>
        <span className="text-sm font-bold text-slate-700 shrink-0">{filledCount}/{butir.items.length}</span>
      </button>
      {open && (
        <div className="p-3 space-y-3 bg-slate-50">
          {items.map(it => (
            <IndicatorCard
              key={it.code} item={it} data={evidence[it.code]}
              onChange={(patch) => onChangeItem(it.code, patch)}
            />
          ))}
          {items.length === 0 && <p className="text-sm text-slate-600 text-center py-3">Tidak ada indikator pada filter ini.</p>}
        </div>
      )}
    </div>
  );
}

/* =========================================================================
   TAMPILAN: DETAIL SEKOLAH (input bukti dukung)
   ========================================================================= */
function SchoolDetailView({ school, evidence, onChangeItem, saveStatus, isAdmin, onUpdateSchoolMeta, onBack, onResetRequest, onDeleteSchoolRequest, allowedTabs, lockedRole }) {
  const tabs = allowedTabs && allowedTabs.length ? allowedTabs : ['kp', 'kk', 'il'];
  const [tab, setTab] = useState(tabs[0]);
  const [filterMode, setFilterMode] = useState('semua');
  const [query, setQuery] = useState('');

  const skorTotal = calcTotalSkor(evidence);
  const pctTotal = skorTotal / TOTAL_MAKS;
  const kat = CATEGORY(pctTotal);
  const { filled: filledTotal, count: countTotal } = calcTotalFilled(evidence);

  const handleChangeItem = (code, patch) => onChangeItem(code, patch);

  const bagianData = INDICATOR_DATA[tab];
  const bagianMeta = BAGIAN_META[tab];
  const bagianSkor = calcBagianSkor(evidence, tab);
  const bagianPct = bagianSkor / bagianMeta.maks;

  const filteredButir = useMemo(() => {
    if (!query.trim()) return bagianData.butir;
    const q = query.toLowerCase();
    return bagianData.butir.map(b => ({
      ...b,
      items: b.items.filter(it => it.code.includes(q) || it.name.toLowerCase().includes(q))
    })).filter(b => b.items.length > 0);
  }, [bagianData, query]);

  return (
    <div className="max-w-4xl mx-auto pb-24 md:pb-8">
      <div className="flex items-center gap-2 mb-3">
        {!lockedRole && (
          <button onClick={onBack} className="p-1.5 rounded-lg hover:bg-slate-100 text-slate-700 md:hidden">
            <ArrowLeft size={18} />
          </button>
        )}
        <div className="flex-1 min-w-0">
          <h2 className="font-serif font-bold text-xl text-[#1C2530] truncate">{school.name}</h2>
          <p className="text-sm text-slate-900 font-semibold">{school.npsn ? `NPSN ${school.npsn}` : 'NPSN belum diisi'} \u00b7 {school.kepsek || 'Kepala sekolah belum diisi'}</p>
        </div>
        {lockedRole && <Badge color={lockedRole === 'guru' ? '2E5395' : '0F6B5C'}>{ROLE_LABEL[lockedRole]}</Badge>}
        <SaveIndicator status={saveStatus} />
      </div>

      <div className="mb-4">
        <DownloadReportMenu school={school} evidence={evidence} />
      </div>

      {isAdmin && <SchoolMetaEditor school={school} onSave={onUpdateSchoolMeta} onDeleteRequest={onDeleteSchoolRequest} />}

      <div className="grid grid-cols-3 gap-2 my-4">
        {['kp', 'kk', 'il'].map(k => {
          const m = BAGIAN_META[k];
          const s = calcBagianSkor(evidence, k);
          const p = s / m.maks;
          return (
            <div key={k} className="rounded-xl border border-slate-400 bg-white p-2.5 text-center">
              <p className="text-xs font-bold text-slate-600 uppercase truncate">{m.label}</p>
              <p className="font-serif font-extrabold text-2xl" style={{ color: `#${m.color}` }}>{Math.round(p * 100)}%</p>
              <p className="text-xs text-slate-600">{s}/{m.maks}</p>
            </div>
          );
        })}
      </div>

      <div className="flex items-center justify-between mb-3 flex-wrap gap-2">
        <Badge color={kat.color} soft={false}>{kat.label} \u00b7 {skorTotal}/{TOTAL_MAKS} ({Math.round(pctTotal * 100)}%)</Badge>
        <span className="text-sm text-slate-900 font-semibold">{filledTotal}/{countTotal} indikator terisi</span>
      </div>

      <div className="flex gap-1.5 mb-3 overflow-x-auto no-scrollbar">
        {tabs.map(k => {
          const m = BAGIAN_META[k];
          const active = tab === k;
          return (
            <button
              key={k} onClick={() => setTab(k)}
              className={`shrink-0 px-3.5 py-2 rounded-lg text-base font-bold transition-colors ${active ? 'text-white' : 'text-slate-700 bg-slate-100 hover:bg-slate-200'}`}
              style={active ? { backgroundColor: `#${m.color}` } : {}}
            >
              {m.label}
            </button>
          );
        })}
      </div>

      <p className="text-sm text-slate-900 font-semibold mb-3">Penanggung jawab: <span className="font-bold text-slate-900">{bagianMeta.peran}</span></p>

      <div className="flex items-center gap-2 mb-3">
        <div className="flex-1 flex items-center gap-1.5 rounded-lg border border-slate-400 bg-white px-2.5 py-1.5">
          <Search size={14} className="text-slate-600" />
          <input value={query} onChange={e => setQuery(e.target.value)} placeholder="Cari kode / kata kunci indikator..." className="w-full text-base focus:outline-none" />
        </div>
        <select value={filterMode} onChange={e => setFilterMode(e.target.value)} className="text-sm font-semibold rounded-lg border border-slate-400 px-2 py-1.5 bg-white text-slate-800">
          <option value="semua">Semua</option>
          <option value="belum">Belum diisi</option>
          <option value="ada">Sudah diisi</option>
        </select>
      </div>

      <div className="space-y-2.5">
        {filteredButir.map((b, idx) => (
          <ButirAccordion key={b.no} butir={b} evidence={evidence} onChangeItem={handleChangeItem} filterMode={filterMode} defaultOpen={idx === 0 || query.trim().length > 0} />
        ))}
        {filteredButir.length === 0 && <p className="text-base text-slate-600 text-center py-8">Tidak ditemukan indikator yang cocok.</p>}
      </div>

      {isAdmin && (
        <button onClick={onResetRequest} className="mt-6 flex items-center gap-1.5 text-sm text-rose-500 hover:text-rose-600 font-semibold">
          <RotateCcw size={13} /> Reset seluruh data bukti dukung sekolah ini
        </button>
      )}
    </div>
  );
}

function SaveIndicator({ status }) {
  if (status === 'saving') return <span className="flex items-center gap-1 text-sm text-slate-600 shrink-0"><Save size={12} className="animate-pulse" /> Menyimpan...</span>;
  if (status === 'saved') return <span className="flex items-center gap-1 text-sm text-[#0F6B5C] shrink-0"><Check size={12} /> Tersimpan</span>;
  if (status === 'error') return <span className="flex items-center gap-1 text-sm text-rose-500 shrink-0"><AlertCircle size={12} /> Gagal menyimpan</span>;
  return null;
}

function DownloadReportMenu({ school, evidence }) {
  const [open, setOpen] = useState(false);
  const [busy, setBusy] = useState(false);

  const runAndClose = async (fn) => {
    setBusy(true);
    try {
      await fn(school, evidence);
    } catch (e) {
      console.error('Gagal membuat laporan:', e);
    }
    setBusy(false);
    setOpen(false);
  };

  return (
    <div className="relative inline-block">
      <button
        onClick={() => setOpen((o) => !o)}
        disabled={busy}
        className="flex items-center gap-1.5 px-3 py-2 rounded-lg text-sm font-bold text-white bg-[#1F3864] hover:bg-[#162a4d] disabled:opacity-60"
      >
        <Download size={15} /> {busy ? 'Menyiapkan...' : 'Unduh Laporan Rekap Sekolah Ini'} <ChevronDown size={14} />
      </button>
      {open && (
        <>
          <div className="fixed inset-0 z-10" onClick={() => setOpen(false)} />
          <div className="absolute left-0 mt-1 w-56 bg-white rounded-xl shadow-lg border border-slate-200 z-20 overflow-hidden">
            <button onClick={() => runAndClose(downloadReportDocx)} className="w-full text-left px-4 py-2.5 text-sm font-semibold text-slate-800 hover:bg-slate-50 flex items-center gap-2">
              <FileText size={15} className="text-[#2E5395]" /> Word (.docx)
            </button>
            <button onClick={() => runAndClose((s, e) => downloadReportPdf(s, e))} className="w-full text-left px-4 py-2.5 text-sm font-semibold text-slate-800 hover:bg-slate-50 flex items-center gap-2 border-t border-slate-100">
              <FileText size={15} className="text-[#B0413E]" /> PDF (.pdf)
            </button>
            <button onClick={() => runAndClose((s, e) => downloadReportHTML(s, e))} className="w-full text-left px-4 py-2.5 text-sm font-semibold text-slate-800 hover:bg-slate-50 flex items-center gap-2 border-t border-slate-100">
              <FileText size={15} className="text-slate-500" /> HTML (bisa dibuka di browser)
            </button>
          </div>
        </>
      )}
    </div>
  );
}

function SchoolMetaEditor({ school, onSave, onDeleteRequest }) {
  const [npsn, setNpsn] = useState(school.npsn || '');
  const [kepsek, setKepsek] = useState(school.kepsek || '');
  const [code, setCode] = useState(school.code || '');
  useEffect(() => { setNpsn(school.npsn || ''); setKepsek(school.kepsek || ''); setCode(school.code || ''); }, [school.id]);
  return (
    <div className="rounded-xl border border-amber-200 bg-amber-50 p-3 mb-3">
      <div className="grid sm:grid-cols-3 gap-2">
        <div>
          <label className="text-xs font-bold text-amber-700 uppercase">NPSN</label>
          <input value={npsn} onChange={e => setNpsn(e.target.value)} onBlur={() => onSave(school.id, { npsn })} className="mt-0.5 w-full text-base rounded-lg border border-amber-200 px-2 py-1.5 focus:outline-none" />
        </div>
        <div>
          <label className="text-xs font-bold text-amber-700 uppercase">Nama Kepala Sekolah</label>
          <input value={kepsek} onChange={e => setKepsek(e.target.value)} onBlur={() => onSave(school.id, { kepsek })} className="mt-0.5 w-full text-base rounded-lg border border-amber-200 px-2 py-1.5 focus:outline-none" />
        </div>
        <div>
          <label className="text-xs font-bold text-amber-700 uppercase">Kode Akses Sekolah</label>
          <input value={code} onChange={e => setCode(e.target.value)} onBlur={() => onSave(school.id, { code })} placeholder="mis. 1001" className="mt-0.5 w-full text-base rounded-lg border border-amber-200 px-2 py-1.5 focus:outline-none font-mono" />
          <p className="text-[11px] text-amber-700 mt-0.5">Dibagikan ke guru/kepala sekolah ini untuk masuk.</p>
        </div>
      </div>
      <button onClick={() => onDeleteRequest(school)} className="mt-3 flex items-center gap-1.5 text-sm font-bold text-rose-600 hover:text-rose-700">
        <Trash2 size={14} /> Hapus Sekolah Ini
      </button>
    </div>
  );
}

/* =========================================================================
   TAMPILAN: BERANDA (dashboard)
   ========================================================================= */
function BerandaView({ schools, evidenceMap, onSelectSchool, isAdmin, onOpenAddSchool }) {
  const stats = schools.map(s => {
    const ev = evidenceMap[s.id] || emptyEvidence();
    const skor = calcTotalSkor(ev);
    const pct = skor / TOTAL_MAKS;
    const { filled, count } = calcTotalFilled(ev);
    return { school: s, skor, pct, filled, count, kat: CATEGORY(pct) };
  });
  const rataRata = stats.reduce((a, s) => a + s.pct, 0) / (stats.length || 1);
  const belumMulai = stats.filter(s => s.filled === 0).length;
  const sangatBaik = stats.filter(s => s.pct >= 0.85).length;

  return (
    <div className="max-w-5xl mx-auto pb-24 md:pb-8">
      <div className="mb-5 flex items-start justify-between gap-3 flex-wrap">
        <div>
          <h2 className="font-serif font-bold text-3xl text-[#1C2530]">Beranda</h2>
          <p className="text-base text-slate-900 font-semibold mt-0.5">Ringkasan kelengkapan bukti dukung akreditasi — {schools.length} sekolah binaan.</p>
        </div>
        {isAdmin && (
          <button
            onClick={onOpenAddSchool}
            className="flex items-center gap-1.5 px-3.5 py-2.5 rounded-xl text-sm font-bold text-white bg-[#0F6B5C] hover:bg-[#0b5347] shrink-0"
          >
            <Plus size={16} /> Tambah Sekolah
          </button>
        )}
      </div>

      <div className="grid grid-cols-3 gap-2.5 mb-6">
        <div className="rounded-2xl bg-[#1F3864] text-white p-4">
          <p className="text-sm uppercase tracking-wide text-white/70">Rata-rata Binaan</p>
          <p className="font-serif font-extrabold text-5xl mt-1">{Math.round(rataRata * 100)}%</p>
        </div>
        <div className="rounded-2xl bg-white border border-slate-400 p-4">
          <p className="text-sm uppercase tracking-wide text-slate-600">Sangat Baik</p>
          <p className="font-serif font-extrabold text-5xl mt-1 text-[#0F6B5C]">{sangatBaik}<span className="text-lg text-slate-700 font-sans font-bold"> /{schools.length}</span></p>
        </div>
        <div className="rounded-2xl bg-white border border-slate-400 p-4">
          <p className="text-sm uppercase tracking-wide text-slate-600">Belum Mulai</p>
          <p className="font-serif font-extrabold text-5xl mt-1 text-rose-500">{belumMulai}<span className="text-lg text-slate-700 font-sans font-bold"> /{schools.length}</span></p>
        </div>
      </div>

      <div className="grid sm:grid-cols-2 gap-3">
        {stats.map(({ school, pct, filled, count, kat }) => (
          <button
            key={school.id} onClick={() => onSelectSchool(school.id)}
            className="flex items-center gap-3.5 rounded-2xl border border-slate-400 bg-white p-4 text-left hover:border-[#1F3864]/40 hover:shadow-sm transition-all"
          >
            <ProgressRing pct={pct} color={kat.color} size={58} />
            <div className="min-w-0 flex-1">
              <p className="font-bold text-base text-[#1C2530] truncate">{school.name}</p>
              <p className="text-sm text-slate-600 mt-0.5">{filled}/{count} indikator terisi</p>
              <Badge color={kat.color}>{kat.label}</Badge>
            </div>
            <ChevronRight size={16} className="text-slate-500 shrink-0" />
          </button>
        ))}
      </div>
    </div>
  );
}

/* =========================================================================
   TAMPILAN: REKAP ANTARSEKOLAH
   ========================================================================= */
function RekapView({ schools, evidenceMap }) {
  const rows = schools.map(s => {
    const ev = evidenceMap[s.id] || emptyEvidence();
    const kp = calcBagianSkor(ev, 'kp'), kk = calcBagianSkor(ev, 'kk'), il = calcBagianSkor(ev, 'il');
    const total = kp + kk + il;
    const pct = total / TOTAL_MAKS;
    return { school: s, kp, kk, il, total, pct, kat: CATEGORY(pct) };
  });

  const chartData = rows.map(r => ({
    name: r.school.name.replace('SMP', '').replace('Negeri', 'N').trim().slice(0, 14),
    'Guru': r.kp, 'Kepala Sekolah': r.kk, 'Iklim': r.il,
  }));

  const butirAvg = useMemo(() => {
    const acc = {};
    for (const k of ['kp', 'kk', 'il']) {
      INDICATOR_DATA[k].butir.forEach(b => {
        let sum = 0, cnt = 0;
        schools.forEach(s => {
          const ev = evidenceMap[s.id] || emptyEvidence();
          b.items.forEach(it => { sum += (ev[it.code]?.skor || 0); cnt++; });
        });
        acc[b.no] = { rumusan: b.rumusan, avgPct: cnt ? sum / (cnt * 4) : 0, bagian: BAGIAN_META[k].label };
      });
    }
    return Object.entries(acc).sort((a, b) => a[1].avgPct - b[1].avgPct);
  }, [schools, evidenceMap]);

  return (
    <div className="max-w-5xl mx-auto pb-24 md:pb-8">
      <div className="mb-5">
        <h2 className="font-serif font-bold text-3xl text-[#1C2530]">Rekap Antarsekolah</h2>
        <p className="text-base text-slate-900 font-semibold mt-0.5">Perbandingan skor bukti dukung {schools.length} sekolah binaan.</p>
      </div>

      <div className="rounded-2xl border border-slate-400 bg-white p-4 mb-5 overflow-x-auto">
        <ResponsiveContainer width="100%" height={280}>
          <BarChart data={chartData} margin={{ top: 5, right: 10, left: -10, bottom: 40 }}>
            <CartesianGrid strokeDasharray="3 3" stroke="#EEF1F5" />
            <XAxis dataKey="name" tick={{ fontSize: 12, fontWeight: 600 }} angle={-35} textAnchor="end" interval={0} height={60} />
            <YAxis tick={{ fontSize: 12, fontWeight: 600 }} />
            <Tooltip contentStyle={{ fontSize: 13, borderRadius: 8, fontWeight: 600 }} />
            <Legend wrapperStyle={{ fontSize: 13, fontWeight: 600 }} />
            <Bar dataKey="Guru" stackId="a" fill="#2E5395" radius={[0, 0, 0, 0]} />
            <Bar dataKey="Kepala Sekolah" stackId="a" fill="#0F6B5C" />
            <Bar dataKey="Iklim" stackId="a" fill="#B5651D" radius={[3, 3, 0, 0]} />
          </BarChart>
        </ResponsiveContainer>
      </div>

      <div className="rounded-2xl border border-slate-400 bg-white overflow-hidden mb-5">
        <table className="w-full text-base">
          <thead>
            <tr className="bg-[#1F3864] text-white text-sm uppercase">
              <th className="text-left px-3 py-2.5 font-bold">Sekolah</th>
              <th className="text-center px-2 py-2.5 font-bold">Guru</th>
              <th className="text-center px-2 py-2.5 font-bold">Kepsek</th>
              <th className="text-center px-2 py-2.5 font-bold">Iklim</th>
              <th className="text-center px-2 py-2.5 font-bold">Total</th>
              <th className="text-center px-2 py-2.5 font-bold">%</th>
              <th className="text-center px-3 py-2.5 font-bold">Kategori</th>
            </tr>
          </thead>
          <tbody>
            {rows.map(({ school, kp, kk, il, total, pct, kat }, i) => (
              <tr key={school.id} className={i % 2 ? 'bg-slate-50' : 'bg-white'}>
                <td className="px-3 py-2 font-semibold text-[#1C2530]">{school.name}</td>
                <td className="text-center px-2 py-2 text-slate-700">{kp}</td>
                <td className="text-center px-2 py-2 text-slate-700">{kk}</td>
                <td className="text-center px-2 py-2 text-slate-700">{il}</td>
                <td className="text-center px-2 py-2 font-bold text-[#1C2530]">{total}</td>
                <td className="text-center px-2 py-2 text-slate-700">{Math.round(pct * 100)}%</td>
                <td className="text-center px-3 py-2"><Badge color={kat.color}>{kat.label}</Badge></td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>

      <div className="rounded-2xl border border-slate-400 bg-white p-4">
        <h3 className="font-bold text-base text-[#1C2530] mb-3 flex items-center gap-1.5"><Filter size={14} className="text-[#B5651D]" /> Butir dengan Skor Rata-rata Terendah (Prioritas Pendampingan)</h3>
        <div className="space-y-2">
          {butirAvg.slice(0, 5).map(([no, d]) => (
            <div key={no} className="flex items-center gap-3">
              <span className="text-sm font-bold text-[#1F3864] w-10 shrink-0">{no}</span>
              <div className="flex-1 min-w-0">
                <p className="text-sm text-slate-800 truncate">{d.rumusan}</p>
                <div className="w-full h-1.5 bg-slate-100 rounded-full mt-1 overflow-hidden">
                  <div className="h-full rounded-full" style={{ width: `${d.avgPct * 100}%`, backgroundColor: `#${CATEGORY(d.avgPct).color}` }} />
                </div>
              </div>
              <span className="text-sm font-bold text-slate-600 w-10 text-right shrink-0">{Math.round(d.avgPct * 100)}%</span>
            </div>
          ))}
        </div>
      </div>
    </div>
  );
}

/* =========================================================================
   NAVIGASI
   ========================================================================= */
function NavItem({ icon: Icon, label, active, onClick }) {
  return (
    <button
      onClick={onClick}
      className={`flex items-center gap-2.5 px-3.5 py-2.5 rounded-xl text-base font-bold w-full transition-colors ${active ? 'bg-[#1F3864] text-white' : 'text-slate-700 hover:bg-slate-100'}`}
    >
      <Icon size={17} /> {label}
    </button>
  );
}
function BottomNavItem({ icon: Icon, label, active, onClick }) {
  return (
    <button onClick={onClick} className="flex flex-col items-center justify-center gap-0.5 flex-1 py-2">
      <Icon size={19} className={active ? 'text-[#1F3864]' : 'text-slate-600'} />
      <span className={`text-xs font-bold ${active ? 'text-[#1F3864]' : 'text-slate-600'}`}>{label}</span>
    </button>
  );
}

/* =========================================================================
   APLIKASI UTAMA
   ========================================================================= */
export default function AkreditasiApp() {
  const [schools, setSchools] = useState(SCHOOL_SEED);
  const [evidenceMap, setEvidenceMap] = useState({});
  const [loaded, setLoaded] = useState(false);
  const [loadError, setLoadError] = useState('');
  const [view, setView] = useState('beranda');
  const [selectedSchoolId, setSelectedSchoolId] = useState(null);
  const [isAdmin, setIsAdmin] = useState(false);
  const [adminCode, setAdminCode] = useState('');
  const [adminModalOpen, setAdminModalOpen] = useState(false);
  const [adminCodeInput, setAdminCodeInput] = useState('');
  const [adminError, setAdminError] = useState('');
  const [adminChecking, setAdminChecking] = useState(false);
  const [resetModalOpen, setResetModalOpen] = useState(false);
  const [addSchoolModalOpen, setAddSchoolModalOpen] = useState(false);
  const [newSchoolName, setNewSchoolName] = useState('');
  const [newSchoolNpsn, setNewSchoolNpsn] = useState('');
  const [newSchoolKepsek, setNewSchoolKepsek] = useState('');
  const [newSchoolCode, setNewSchoolCode] = useState('');
  const [addSchoolError, setAddSchoolError] = useState('');
  const [addSchoolChecking, setAddSchoolChecking] = useState(false);
  const [deleteSchoolTarget, setDeleteSchoolTarget] = useState(null);
  const [deleteSchoolChecking, setDeleteSchoolChecking] = useState(false);
  const [deleteSchoolError, setDeleteSchoolError] = useState('');
  const [saveStatus, setSaveStatus] = useState('idle');
  const [myRole, setMyRole] = useState(null); // null | 'guru' | 'kepsek' | 'pengawas'
  const [roleError, setRoleError] = useState('');
  const [roleChecking, setRoleChecking] = useState(false);
  const saveTimer = useRef(null);

  useEffect(() => {
    (async () => {
      try {
        const finalSchools = await apiGetSchools();
        setSchools(finalSchools);
        const evPairs = await Promise.all(finalSchools.map(async s => {
          const ev = await apiGetEvidence(s.id);
          return [s.id, ev || emptyEvidence()];
        }));
        const evMap = {};
        evPairs.forEach(([id, ev]) => { evMap[id] = ev; });
        setEvidenceMap(evMap);
        setLoaded(true);
      } catch (e) {
        setLoadError(e.code === 'REDIS_NOT_CONFIGURED'
          ? 'Server belum bisa menyimpan/membaca data (database Redis/Upstash belum terkonfigurasi). Data sekolah & bukti dukung TIDAK akan tersimpan sampai ini diperbaiki — lihat README bagian Troubleshooting (pasang integrasi "Upstash for Redis" lewat tab Storage di dashboard Vercel).'
          : 'Gagal memuat data dari server. Coba muat ulang halaman. Data yang ditampilkan mungkin belum tersinkron.');
        setLoaded(true);
      }
    })();
  }, []);

  const handleChangeItem = useCallback((schoolId, code, patch) => {
    setEvidenceMap(prev => {
      const prevEv = prev[schoolId] || emptyEvidence();
      const updatedEv = { ...prevEv, [code]: { ...prevEv[code], ...patch, updatedAt: Date.now() } };
      const next = { ...prev, [schoolId]: updatedEv };
      setSaveStatus('saving');
      if (saveTimer.current) clearTimeout(saveTimer.current);
      saveTimer.current = setTimeout(async () => {
        const ok = await apiSaveEvidence(schoolId, updatedEv);
        setSaveStatus(ok ? 'saved' : 'error');
      }, 650);
      return next;
    });
  }, []);

  const handleUpdateSchoolMeta = useCallback((schoolId, patch) => {
    setSchools(prev => prev.map(s => s.id === schoolId ? { ...s, ...patch } : s));
    apiUpdateSchool(schoolId, patch, adminCode).catch(() => {
      // rollback tidak kritis di sini karena field ini jarang bentrok; cukup beri tahu lewat status simpan
      setSaveStatus('error');
    });
  }, [adminCode]);

  const handleSelectSchool = (id) => { setSelectedSchoolId(id); setView('sekolah'); };

  const handleRoleLogin = async (role, schoolId, code) => {
    setRoleChecking(true);
    setRoleError('');
    if (role === 'pengawas') {
      const { ok, usingDefault, providedLength, expectedLength } = await apiVerifyAdmin(code);
      setRoleChecking(false);
      if (ok) {
        setMyRole('pengawas');
        setIsAdmin(true);
        setAdminCode(code.trim());
        setView('beranda');
      } else if (usingDefault && code.trim() !== 'akreditasi2026') {
        setRoleError('Kode akses Pengawas salah. Catatan: server ini sepertinya masih memakai kode bawaan ("akreditasi2026"), bukan ADMIN_CODE kustom yang sudah Anda atur. Jika Anda sudah mengubah ADMIN_CODE di Netlify, pastikan sudah klik "Trigger deploy" ulang dan environment variable-nya mencakup scope "Functions".');
      } else if (typeof providedLength === 'number' && typeof expectedLength === 'number') {
        setRoleError(`Kode akses Pengawas salah. (Panjang kode yang Anda masukkan: ${providedLength} karakter; panjang yang tersimpan di server: ${expectedLength} karakter. Kalau angkanya beda, ada karakter yang tidak sengaja ikut/hilang saat menyimpan atau mengetik kodenya — coba salin-tempel ulang dari Netlify, jangan diketik manual.)`);
      } else {
        setRoleError('Kode akses Pengawas salah. Periksa kembali kode ADMIN_CODE Anda.');
      }
      return;
    }
    const school = schools.find(s => s.id === schoolId);
    setTimeout(() => {
      setRoleChecking(false);
      if (school && school.code && school.code === code) {
        setMyRole(role);
        setSelectedSchoolId(schoolId);
        setView('sekolah');
      } else {
        setRoleError('Sekolah atau kode akses salah. Periksa kembali dengan Pengawas Pembina.');
      }
    }, 300);
  };

  const handleKeluar = () => {
    setMyRole(null);
    setSelectedSchoolId(null);
    setIsAdmin(false);
    setAdminCode('');
    setView('beranda');
  };

  const handleAdminSubmit = async () => {
    setAdminChecking(true);
    setAdminError('');
    const { ok, usingDefault, providedLength, expectedLength } = await apiVerifyAdmin(adminCodeInput.trim());
    setAdminChecking(false);
    if (ok) {
      setIsAdmin(true); setAdminCode(adminCodeInput.trim()); setAdminModalOpen(false); setAdminCodeInput('');
    } else if (usingDefault && adminCodeInput.trim() !== 'akreditasi2026') {
      setAdminError('Kode salah. Catatan: server sepertinya masih memakai kode bawaan, bukan ADMIN_CODE kustom Anda — cek kembali pengaturan environment variable & deploy ulang di Netlify.');
    } else if (typeof providedLength === 'number' && typeof expectedLength === 'number') {
      setAdminError(`Kode admin salah. (Panjang yang dimasukkan: ${providedLength}, panjang tersimpan: ${expectedLength} karakter — kalau beda, coba salin-tempel ulang kodenya.)`);
    } else {
      setAdminError('Kode admin salah. Coba lagi.');
    }
  };

  const handleResetConfirm = async () => {
    try {
      const empty = await apiResetEvidence(selectedSchoolId, adminCode);
      setEvidenceMap(prev => ({ ...prev, [selectedSchoolId]: empty }));
      setResetModalOpen(false);
    } catch (e) {
      setAdminError(e.message);
    }
  };

  const generateSchoolCode = () => String(Math.floor(1000 + Math.random() * 9000));

  const handleOpenAddSchool = () => {
    setNewSchoolName(''); setNewSchoolNpsn(''); setNewSchoolKepsek('');
    setNewSchoolCode(generateSchoolCode());
    setAddSchoolError('');
    setAddSchoolModalOpen(true);
  };

  const handleAddSchoolSubmit = async () => {
    const name = newSchoolName.trim();
    const code = newSchoolCode.trim();
    if (!name) { setAddSchoolError('Nama sekolah wajib diisi.'); return; }
    if (!code) { setAddSchoolError('Kode akses wajib diisi.'); return; }
    setAddSchoolChecking(true);
    setAddSchoolError('');
    try {
      const result = await apiAddSchool({ name, npsn: newSchoolNpsn.trim(), kepsek: newSchoolKepsek.trim(), code }, adminCode);
      setSchools(result.list);
      setEvidenceMap(prev => ({ ...prev, [result.school.id]: emptyEvidence() }));
      setAddSchoolModalOpen(false);
    } catch (e) {
      setAddSchoolError(e.message);
    }
    setAddSchoolChecking(false);
  };

  const handleDeleteSchoolConfirm = async () => {
    if (!deleteSchoolTarget) return;
    if (schools.length <= 1) { setDeleteSchoolTarget(null); return; }
    setDeleteSchoolChecking(true);
    setDeleteSchoolError('');
    try {
      const result = await apiDeleteSchool(deleteSchoolTarget.id, adminCode);
      setSchools(result.list);
      setEvidenceMap(prev => {
        const next = { ...prev };
        delete next[deleteSchoolTarget.id];
        return next;
      });
      if (selectedSchoolId === deleteSchoolTarget.id) {
        setSelectedSchoolId(null);
        setView('beranda');
      }
      setDeleteSchoolTarget(null);
    } catch (e) {
      setDeleteSchoolError(e.message);
    }
    setDeleteSchoolChecking(false);
  };

  const selectedSchool = schools.find(s => s.id === selectedSchoolId);
  const currentEvidence = selectedSchoolId ? (evidenceMap[selectedSchoolId] || emptyEvidence()) : null;

  if (!loaded) {
    return (
      <div className="min-h-screen flex items-center justify-center bg-[#6FA8DC]">
        <div className="text-center">
          <div className="w-10 h-10 border-3 border-[#1F3864] border-t-transparent rounded-full animate-spin mx-auto mb-3" />
          <p className="text-base text-slate-900 font-semibold">Memuat data bukti dukung...</p>
        </div>
      </div>
    );
  }

  if (!myRole) {
    return (
      <>
        {loadError && (
          <div className="fixed top-0 inset-x-0 z-50 bg-rose-600 text-white text-sm text-center py-2 px-4">
            {loadError}
          </div>
        )}
        <RoleGateScreen schools={schools} onLogin={handleRoleLogin} error={roleError} checking={roleChecking} />
      </>
    );
  }

  const isLocked = myRole !== 'pengawas';
  const allowedTabs = ROLE_TABS[myRole];

  return (
    <div className="min-h-screen bg-[#6FA8DC]">
      {loadError && (
        <div className="bg-rose-50 border-b border-rose-200 text-rose-600 text-sm text-center py-2 px-4">
          {loadError}
        </div>
      )}

      {/* Header */}
      <header className="sticky top-0 z-30 bg-white/90 backdrop-blur border-b border-slate-400">
        <div className="max-w-6xl mx-auto px-4 py-3 flex items-center justify-between">
          <div className="flex items-center gap-2.5">
            <div className="w-9 h-9 rounded-xl bg-[#1F3864] flex items-center justify-center shrink-0">
              <ClipboardList size={18} className="text-white" />
            </div>
            <div className="min-w-0">
              <p className="font-serif font-bold text-lg text-[#1C2530] leading-tight truncate">Bukti Dukung Akreditasi</p>
              <p className="text-xs text-slate-600 leading-tight">{isLocked ? `${ROLE_LABEL[myRole]} \u00b7 ${selectedSchool?.name || ''}` : 'SMP 2026 \u00b7 Kab. Tanah Laut'}</p>
            </div>
          </div>
          {isLocked ? (
            <button
              onClick={handleKeluar}
              className="flex items-center gap-1.5 px-2.5 py-1.5 rounded-lg text-sm font-bold shrink-0 bg-slate-100 text-slate-700"
            >
              <ArrowLeft size={13} /> Keluar
            </button>
          ) : (
            <div className="flex items-center gap-1.5 shrink-0">
              <button
                onClick={() => isAdmin ? (setIsAdmin(false), setAdminCode('')) : setAdminModalOpen(true)}
                className={`flex items-center gap-1.5 px-2.5 py-1.5 rounded-lg text-sm font-bold ${isAdmin ? 'bg-[#0F6B5C]/10 text-[#0F6B5C]' : 'bg-slate-100 text-slate-700'}`}
              >
                {isAdmin ? <Unlock size={13} /> : <Lock size={13} />} {isAdmin ? 'Mode Pengawas' : 'Masuk Pengawas'}
              </button>
              <button
                onClick={handleKeluar}
                className="flex items-center gap-1.5 px-2.5 py-1.5 rounded-lg text-sm font-bold bg-slate-100 text-slate-700"
                title="Keluar dari peran Pengawas"
              >
                <LogOut size={13} /> Keluar
              </button>
            </div>
          )}
        </div>
      </header>

      <div className="max-w-6xl mx-auto flex">
        {/* Sidebar (desktop) */}
        <aside className="hidden md:flex flex-col gap-1.5 w-56 shrink-0 p-4">
          {isLocked ? (
            <div className="rounded-xl bg-white p-3.5">
              <p className="text-xs font-bold text-slate-600 uppercase">Masuk sebagai</p>
              <p className="font-bold text-base text-slate-900">{ROLE_LABEL[myRole]}</p>
              <p className="text-sm text-slate-700 font-semibold truncate">{selectedSchool?.name}</p>
              <button onClick={handleKeluar} className="mt-3 flex items-center gap-1.5 text-sm font-bold text-rose-600">
                <ArrowLeft size={14} /> Keluar
              </button>
            </div>
          ) : (
            <>
              <NavItem icon={LayoutDashboard} label="Beranda" active={view === 'beranda'} onClick={() => setView('beranda')} />
              <NavItem icon={Building2} label="Pilih Sekolah" active={view === 'sekolah'} onClick={() => { if (!selectedSchoolId) setSelectedSchoolId(schools[0]?.id); setView('sekolah'); }} />
              <NavItem icon={BarChart3} label="Rekap Antarsekolah" active={view === 'rekap'} onClick={() => setView('rekap')} />
              {isAdmin && (
                <div className="mt-4 rounded-xl bg-[#0F6B5C]/10 p-3 text-sm text-[#0F6B5C] flex items-start gap-1.5">
                  <ShieldCheck size={14} className="shrink-0 mt-0.5" />
                  <span>Mode Pengawas aktif — Anda dapat menyunting data sekolah & mereset bukti dukung.</span>
                </div>
              )}
              <button onClick={handleKeluar} className="mt-4 flex items-center gap-1.5 px-3.5 py-2.5 rounded-xl text-sm font-bold text-rose-600 hover:bg-rose-50 w-full">
                <LogOut size={16} /> Keluar
              </button>
            </>
          )}
        </aside>

        {/* Konten */}
        <main className="flex-1 p-4 md:p-6 min-w-0">
          {!isLocked && view === 'beranda' && <BerandaView schools={schools} evidenceMap={evidenceMap} onSelectSchool={handleSelectSchool} isAdmin={isAdmin} onOpenAddSchool={handleOpenAddSchool} />}
          {(isLocked || view === 'sekolah') && selectedSchool && (
            <>
              {!isLocked && (
                <div className="hidden md:block mb-4">
                  <SchoolPicker schools={schools} value={selectedSchoolId} onChange={setSelectedSchoolId} />
                </div>
              )}
              <SchoolDetailView
                school={selectedSchool}
                evidence={currentEvidence}
                onChangeItem={(code, patch) => handleChangeItem(selectedSchoolId, code, patch)}
                saveStatus={saveStatus}
                isAdmin={isAdmin}
                onUpdateSchoolMeta={handleUpdateSchoolMeta}
                onBack={() => setView('beranda')}
                onResetRequest={() => setResetModalOpen(true)}
                onDeleteSchoolRequest={(s) => setDeleteSchoolTarget(s)}
                allowedTabs={allowedTabs}
                lockedRole={isLocked ? myRole : null}
              />
            </>
          )}
          {!isLocked && view === 'sekolah' && !selectedSchool && (
            <div className="max-w-md mx-auto mt-10">
              <SchoolPicker schools={schools} value={selectedSchoolId} onChange={setSelectedSchoolId} />
            </div>
          )}
          {!isLocked && view === 'rekap' && <RekapView schools={schools} evidenceMap={evidenceMap} />}
        </main>
      </div>

      {/* Bottom nav (mobile) */}
      <nav className="md:hidden fixed bottom-0 inset-x-0 z-30 bg-white border-t border-slate-400 flex">
        <BottomNavItem icon={LayoutDashboard} label="Beranda" active={view === 'beranda'} onClick={() => setView('beranda')} />
        <BottomNavItem icon={Building2} label="Sekolah" active={view === 'sekolah'} onClick={() => { if (!selectedSchoolId) setSelectedSchoolId(schools[0]?.id); setView('sekolah'); }} />
        <BottomNavItem icon={BarChart3} label="Rekap" active={view === 'rekap'} onClick={() => setView('rekap')} />
      </nav>

      {/* Modal: Masuk Mode Pengawas */}
      <Modal
        open={adminModalOpen} onClose={() => { setAdminModalOpen(false); setAdminError(''); setAdminCodeInput(''); }}
        title="Masuk Mode Pengawas"
        footer={<>
          <button onClick={() => setAdminModalOpen(false)} className="px-3.5 py-2 rounded-lg text-base font-semibold text-slate-700 hover:bg-slate-100">Batal</button>
          <button onClick={handleAdminSubmit} disabled={adminChecking} className="px-3.5 py-2 rounded-lg text-base font-bold text-white bg-[#1F3864] hover:bg-[#162a4d] disabled:opacity-50">{adminChecking ? 'Memeriksa...' : 'Masuk'}</button>
        </>}
      >
        <p className="text-base text-slate-700 mb-3">Masukkan kode akses Pengawas untuk menyunting data induk sekolah dan mereset data bukti dukung.</p>
        <input
          type="password" value={adminCodeInput} onChange={e => setAdminCodeInput(e.target.value)}
          onKeyDown={e => e.key === 'Enter' && handleAdminSubmit()}
          placeholder="Kode akses" autoFocus
          className="w-full rounded-lg border border-slate-400 px-3 py-2 text-base focus:outline-none focus:ring-2 focus:ring-[#1F3864]/30"
        />
        {adminError && <p className="text-sm text-rose-500 mt-2 flex items-center gap-1"><AlertCircle size={12} /> {adminError}</p>}
      </Modal>

      {/* Modal: Konfirmasi Reset */}
      <Modal
        open={resetModalOpen} onClose={() => setResetModalOpen(false)}
        title="Reset Data Bukti Dukung?"
        footer={<>
          <button onClick={() => setResetModalOpen(false)} className="px-3.5 py-2 rounded-lg text-base font-semibold text-slate-700 hover:bg-slate-100">Batal</button>
          <button onClick={handleResetConfirm} className="px-3.5 py-2 rounded-lg text-base font-bold text-white bg-rose-500 hover:bg-rose-600">Ya, Reset</button>
        </>}
      >
        <p className="text-base text-slate-700">Seluruh skor, catatan, dan bukti dukung untuk <span className="font-bold text-[#1C2530]">{selectedSchool?.name}</span> akan dihapus dan tidak dapat dikembalikan. Lanjutkan?</p>
      </Modal>

      {/* Modal: Konfirmasi Hapus Sekolah */}
      <Modal
        open={!!deleteSchoolTarget} onClose={() => { setDeleteSchoolTarget(null); setDeleteSchoolError(''); }}
        title="Hapus Sekolah Ini?"
        footer={<>
          <button onClick={() => { setDeleteSchoolTarget(null); setDeleteSchoolError(''); }} className="px-3.5 py-2 rounded-lg text-base font-semibold text-slate-700 hover:bg-slate-100">Batal</button>
          {schools.length > 1 && (
            <button onClick={handleDeleteSchoolConfirm} disabled={deleteSchoolChecking} className="px-3.5 py-2 rounded-lg text-base font-bold text-white bg-rose-500 hover:bg-rose-600 disabled:opacity-50">
              {deleteSchoolChecking ? 'Menghapus...' : 'Ya, Hapus Sekolah'}
            </button>
          )}
        </>}
      >
        {schools.length > 1 ? (
          <p className="text-base text-slate-700">
            Sekolah <span className="font-bold text-[#1C2530]">{deleteSchoolTarget?.name}</span> beserta seluruh data bukti dukungnya akan
            dihapus permanen dari daftar sekolah binaan dan tidak dapat dikembalikan. Kode aksesnya juga tidak akan berfungsi lagi. Lanjutkan?
          </p>
        ) : (
          <p className="text-base text-rose-600 font-semibold">Tidak bisa menghapus — minimal harus ada 1 sekolah binaan tersisa.</p>
        )}
        {deleteSchoolError && <p className="text-sm text-rose-500 mt-2 flex items-center gap-1"><AlertCircle size={12} /> {deleteSchoolError}</p>}
      </Modal>

      {/* Modal: Tambah Sekolah (khusus Pengawas) */}
      <Modal
        open={addSchoolModalOpen}
        onClose={() => { setAddSchoolModalOpen(false); setAddSchoolError(''); setNewSchoolName(''); setNewSchoolNpsn(''); setNewSchoolKepsek(''); }}
        title="Tambah Sekolah Binaan"
        footer={<>
          <button onClick={() => setAddSchoolModalOpen(false)} className="px-3.5 py-2 rounded-lg text-base font-semibold text-slate-700 hover:bg-slate-100">Batal</button>
          <button onClick={handleAddSchoolSubmit} disabled={addSchoolChecking} className="px-3.5 py-2 rounded-lg text-base font-bold text-white bg-[#0F6B5C] hover:bg-[#0b5347] disabled:opacity-50">
            {addSchoolChecking ? 'Menyimpan...' : 'Tambah'}
          </button>
        </>}
      >
        <div className="space-y-3">
          <div>
            <label className="text-xs font-bold text-slate-700 uppercase tracking-wide">Nama Sekolah *</label>
            <input value={newSchoolName} onChange={e => setNewSchoolName(e.target.value)} autoFocus placeholder="mis. SMPN 2 Pelaihari" className="mt-1 w-full rounded-lg border border-slate-400 px-3 py-2 text-base focus:outline-none focus:ring-2 focus:ring-[#1F3864]/30" />
          </div>
          <div>
            <label className="text-xs font-bold text-slate-700 uppercase tracking-wide">NPSN (opsional)</label>
            <input value={newSchoolNpsn} onChange={e => setNewSchoolNpsn(e.target.value)} className="mt-1 w-full rounded-lg border border-slate-400 px-3 py-2 text-base focus:outline-none focus:ring-2 focus:ring-[#1F3864]/30" />
          </div>
          <div>
            <label className="text-xs font-bold text-slate-700 uppercase tracking-wide">Nama Kepala Sekolah (opsional)</label>
            <input value={newSchoolKepsek} onChange={e => setNewSchoolKepsek(e.target.value)} className="mt-1 w-full rounded-lg border border-slate-400 px-3 py-2 text-base focus:outline-none focus:ring-2 focus:ring-[#1F3864]/30" />
          </div>
          <div>
            <label className="text-xs font-bold text-slate-700 uppercase tracking-wide">Kode Akses Sekolah</label>
            <input value={newSchoolCode} onChange={e => setNewSchoolCode(e.target.value)} className="mt-1 w-full rounded-lg border border-slate-400 px-3 py-2 text-base font-mono focus:outline-none focus:ring-2 focus:ring-[#1F3864]/30" />
            <p className="text-[11px] text-slate-600 mt-1">Terisi acak, boleh diganti. Bagikan kode ini ke guru/kepala sekolahnya.</p>
          </div>
          {addSchoolError && <p className="text-sm text-rose-500 flex items-center gap-1"><AlertCircle size={12} /> {addSchoolError}</p>}
        </div>
      </Modal>
    </div>
  );
}

function SchoolPicker({ schools, value, onChange }) {
  return (
    <div className="flex items-center gap-2 rounded-xl border border-slate-400 bg-white px-3 py-2">
      <Building2 size={16} className="text-slate-600 shrink-0" />
      <select value={value || ''} onChange={e => onChange(e.target.value)} className="w-full text-base font-semibold text-[#1C2530] focus:outline-none bg-transparent">
        <option value="" disabled>Pilih sekolah...</option>
        {schools.map(s => <option key={s.id} value={s.id}>{s.name}</option>)}
      </select>
    </div>
  );
}
