/**
 * Converter.js — Amazon Seller Central
 * Suporta:
 * 1. TXT (Tab-Separated) → Excel Padrão (NOVO)
 * 2. XLSX mal-separado → Excel Padrão
 * 3. TXT → XML DocBook
 */
(function (global) {
  'use strict';

  // Colunas do modelo padrão (MODELO_AMAZON.xlsx)
  const MODEL_COLS = [
    'amazon-order-id', 'merchant-order-id', 'purchase-date', 'last-updated-date',
    'order-status', 'fulfillment-channel', 'sales-channel', 'order-channel',
    'url', 'ship-service-level', 'product-name', 'sku', 'asin', 'item-status',
    'quantity', 'currency', 'item-price', 'item-tax', 'shipping-price',
    'shipping-tax', 'gift-wrap-price', 'gift-wrap-tax', 'item-promotion-discount',
    'ship-promotion-discount', 'ship-city', 'ship-state', 'ship-postal-code',
    'ship-country', 'promotion-ids', 'payment-method-details', 'cpf',
    'ship-county', 'item-extensions-data', 'is-business-order',
    'purchase-order-number', 'price-designation',
  ];

  const SEP_XLSX = '\xa0 \xa0 \xa0 \xa0 ';

  /**
   * Faz o parse de arquivo TXT (separado por tabs \t) da Amazon.
   */
  function parseTxt(textContent) {
    // Quebra por quebras de linha preservando linhas preenchidas
    const lines = textContent.split(/\r?\n/).filter(line => line.trim() !== '');
    if (!lines.length) throw new Error('O arquivo TXT está vazio.');

    // A primeira linha são os cabeçalhos originais separados por Tab
    const sourceHeaders = lines[0].split('\t').map(h => h.trim());
    
    // As próximas linhas são os dados
    const dataRows = lines.slice(1).map(line => {
      const parts = line.split('\t').map(v => v.trim());
      const obj = {};
      sourceHeaders.forEach((h, i) => { 
        obj[h] = parts[i] !== undefined ? parts[i] : ''; 
      });
      return obj;
    });

    return mapToModel(dataRows);
  }

  /**
   * Faz o parse do XLSX Amazon mal-separado.
   */
  function parseXlsx(arrayBuffer) {
    const wb = XLSX.read(arrayBuffer, { type: 'array' });
    const ws = wb.Sheets[wb.SheetNames[0]];
    const rawRows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });

    if (!rawRows.length) throw new Error('O arquivo XLSX está vazio.');

    const headerRow = String(rawRows[0][0] || '');
    const isMalSep = headerRow.includes(SEP_XLSX);

    let sourceHeaders, dataRows;

    if (isMalSep) {
      sourceHeaders = headerRow.split(SEP_XLSX).map(h => h.trim());
      dataRows = rawRows.slice(1).map(row => {
        const parts = String(row[0] || '').split(SEP_XLSX).map(v => v.trim());
        const obj = {};
        sourceHeaders.forEach((h, i) => { obj[h] = parts[i] !== undefined ? parts[i] : ''; });
        return obj;
      });
    } else {
      const jsonRows = XLSX.utils.sheet_to_json(ws, { defval: '' });
      sourceHeaders = Object.keys(jsonRows[0] || {});
      dataRows = jsonRows;
    }

    return mapToModel(dataRows);
  }

  /**
   * Função auxiliar que pega as linhas brutas extraídas e as força a obedecer 
   * as colunas rígidas do modelo da Amazon e insere a coluna Ref.
   */
  function mapToModel(dataRows) {
    const finalHeaders = ['Ref.', ...MODEL_COLS];
    const finalRows = dataRows.map((src, idx) => {
      const out = { 'Ref.': idx + 1 };
      MODEL_COLS.forEach(col => {
        let val = src[col];
        if (val === undefined) {
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
   * Constrói o workbook SheetJS a partir dos dados parseados.
   */
  function buildWorkbook(headers, rows) {
    const wb = XLSX.utils.book_new();
    const sheetData = [headers];
    
    rows.forEach(row => {
      sheetData.push(headers.map(h => row[h] !== undefined ? row[h] : ''));
    });

    const ws = XLSX.utils.aoa_to_sheet(sheetData);

    // Ajusta as larguras das colunas para visualização mais limpa
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
   * Converte TXT (Amazon Seller Central) para XML DocBook (Modo original)
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
    parseTxt,
    parseXlsx,
    buildWorkbook,
    toXml,
    MODEL_COLS,
  };

})(window);

