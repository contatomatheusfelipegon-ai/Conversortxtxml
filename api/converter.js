/**
 * Converter.js — Amazon Seller Central
 * Suporta:
 *  1. XLSX mal-separado (relatório Amazon com \xa0 como separador) → Excel formatado
 *  2. TXT → XML DocBook
 */
(function (global) {
  'use strict';

  // Colunas do modelo padrão (MODELO_AMAZON.xlsx) — sem Ref., sem colunas extras da Amazon
  const MODEL_COLS = [
    'amazon-order-id',
    'merchant-order-id',
    'purchase-date',
    'last-updated-date',
    'order-status',
    'fulfillment-channel',
    'sales-channel',
    'order-channel',
    'url',
    'ship-service-level',
    'product-name',
    'sku',
    'asin',
    'item-status',
    'quantity',
    'currency',
    'item-price',
    'item-tax',
    'shipping-price',
    'shipping-tax',
    'gift-wrap-price',
    'gift-wrap-tax',
    'item-promotion-discount',
    'ship-promotion-discount',
    'ship-city',
    'ship-state',
    'ship-postal-code',
    'ship-country',
    'promotion-ids',
    'payment-method-details',
    'cpf',
    'ship-county',
    'item-extensions-data',
    'is-business-order',
    'purchase-order-number',
    'price-designation',
  ];

  // Separador usado pelo relatório Amazon mal-exportado
  const SEP = '\xa0 \xa0 \xa0 \xa0 ';

  /**
   * Detecta se um workbook SheetJS tem o formato Amazon mal-separado
   * (única coluna cujo nome contém o separador especial)
   */
  function isAmazonMalSeparated(wb) {
    const ws = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
    if (!rows.length) return false;
    const firstRow = rows[0];
    // Tem apenas 1 coluna e o header contém o separador
    return firstRow.length === 1 && String(firstRow[0]).includes(SEP);
  }

  /**
   * Faz o parse do XLSX Amazon mal-separado.
   * Retorna { headers: string[], rows: object[] }
   * headers = ['Ref.', ...MODEL_COLS]
   */
  function parseXlsx(arrayBuffer) {
    const wb = XLSX.read(arrayBuffer, { type: 'array' });
    const ws = wb.Sheets[wb.SheetNames[0]];
    const rawRows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });

    if (!rawRows.length) throw new Error('Arquivo vazio.');

    const headerRow = String(rawRows[0][0] || '');
    const isMalSep = headerRow.includes(SEP);

    let sourceHeaders, dataRows;

    if (isMalSep) {
      // Extrai headers originais da primeira linha
      sourceHeaders = headerRow.split(SEP).map(h => h.trim());
      // Linhas de dados (a partir da linha 1)
      dataRows = rawRows.slice(1).map(row => {
        const parts = String(row[0] || '').split(SEP).map(v => v.trim());
        const obj = {};
        sourceHeaders.forEach((h, i) => { obj[h] = parts[i] !== undefined ? parts[i] : ''; });
        return obj;
      });
    } else {
      // Arquivo já tem múltiplas colunas — lê normalmente
      const jsonRows = XLSX.utils.sheet_to_json(ws, { defval: '' });
      sourceHeaders = Object.keys(jsonRows[0] || {});
      dataRows = jsonRows;
    }

    // Mapeia para colunas do modelo (descarta colunas extras da Amazon)
    const finalHeaders = ['Ref.', ...MODEL_COLS];
    const finalRows = dataRows.map((src, idx) => {
      const out = { 'Ref.': idx + 1 };
      MODEL_COLS.forEach(col => {
        // Tenta encontrar a coluna por nome exato, depois por trim
        let val = src[col];
        if (val === undefined) {
          // fallback: busca case-insensitive
          const found = Object.keys(src).find(k => k.trim().toLowerCase() === col.toLowerCase());
          val = found ? src[found] : '';
        }
        out[col] = val !== undefined && val !== null ? String(val).trim() : '';
      });
      return out;
    });

    return { headers: finalHeaders, rows: finalRows };
  }

  /**
   * Constrói o workbook SheetJS a partir dos dados parseados,
   * aplicando a formatação visual do modelo Amazon.
   */
  function buildWorkbook(headers, rows) {
    const wb = XLSX.utils.book_new();

    // Monta array de arrays para a planilha
    const sheetData = [headers];
    rows.forEach(row => {
      sheetData.push(headers.map(h => row[h] !== undefined ? row[h] : ''));
    });

    const ws = XLSX.utils.aoa_to_sheet(sheetData);

    // Larguras das colunas (estimativa)
    ws['!cols'] = headers.map(h => {
      const maxLen = Math.max(h.length, 14);
      if (h === 'product-name') return { wch: 50 };
      if (h === 'Ref.') return { wch: 6 };
      if (['purchase-date', 'last-updated-date'].includes(h)) return { wch: 22 };
      if (['cpf', 'amazon-order-id'].includes(h)) return { wch: 20 };
      return { wch: Math.min(maxLen + 4, 30) };
    });

    XLSX.utils.book_append_sheet(wb, ws, 'Planilha1');
    return wb;
  }

  /**
   * Converte TXT (Amazon Seller Central) para XML DocBook
   */
  function toXml(rawBinary) {
    const lines = rawBinary.split('\n');
    if (lines.length && lines[lines.length - 1].trim() === '') lines.pop();

    const paras = lines.map(line => {
      const hasCR = line.endsWith('\r');
      const content = hasCR ? line.slice(0, -1) : line;
      const spaced = content.replace(/\t/g, '        ');
      const escaped = escapeXml(spaced);
      return '  <para>' + escaped + (hasCR ? '&#13;' : '') + '</para>';
    }).join('\n');

    return '<?xml version="1.0" encoding="UTF-8"?>\n' +
      '<!DOCTYPE article PUBLIC "-//OASIS//DTD DocBook XML V4.1.2//EN" ' +
      '"http://www.oasis-open.org/docbook/xml/4.1.2/docbookx.dtd">\n' +
      '<article lang="">\n' +
      paras + '\n' +
      '</article>\n';
  }

  function escapeXml(str) {
    return str
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;');
  }

  global.AmazonConverter = {
    parseXlsx,
    buildWorkbook,
    toXml,
    MODEL_COLS,
  };

})(window);
