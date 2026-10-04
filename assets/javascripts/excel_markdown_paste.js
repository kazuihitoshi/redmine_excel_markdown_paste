/*
 * Copyright 2026 kazuihitoshi@gmail.com
 *
 * This program is free software; you can redistribute it and/or
 * modify it under the terms of the GNU General Public License
 * as published by the Free Software Foundation; either version 2
 * of the License, or (at your option) any later version.
 *
 * Excel Markdown Paste
 * Converts a pasted Excel (or other spreadsheet) cell range into a
 * CommonMark / GitHub-flavored Markdown table.
 *
 * Image paste is not handled here. Redmine inserts those as ![]().
 * Paste inside a fenced code block (``` or ~~~) is left unchanged.
 */
(function (root, factory) {
  var api = factory();
  if (typeof module === 'object' && module.exports) {
    module.exports = api;
  }
  if (root) {
    root.RedmineExcelMarkdownPaste = api;
  }
  if (typeof document !== 'undefined') {
    api.install();
  }
})(typeof globalThis !== 'undefined' ? globalThis : this, function () {
  'use strict';

  var installed = false;

  function install() {
    if (installed || typeof document === 'undefined') return;
    installed = true;
    document.addEventListener('paste', onPaste, true);
  }

  function onPaste(event) {
    try {
      var textarea = event.target;
      if (!shouldHandle(textarea)) return;

      var clipboard = event.clipboardData;
      if (!clipboard) return;

      // Screenshots and copied images belong to Redmine (![]()).
      // Do not read, cancel, or rewrite that paste.
      if (isPureImagePaste(clipboard)) return;

      // check if the selection is inside a code fence
      if (insideFence(textarea.value, textarea.selectionStart || 0)) return;

      var html = readClipboard(clipboard, 'text/html');
      var plain = readClipboard(clipboard, 'text/plain');
      if (shouldIgnorePaste(clipboardHasImage(clipboard), html, plain)) return;

      var markdown = convertClipboard(html, plain);
      if (!markdown) return;

      event.preventDefault();
      event.stopPropagation();

      var start = textarea.selectionStart || 0;
      var end = textarea.selectionEnd == null ? start : textarea.selectionEnd;
      var block = ensureBlock(textarea.value.slice(0, start), textarea.value.slice(end), markdown);
      insertText(textarea, block, start, end);
    } catch (error) {
      if (typeof console !== 'undefined' && console.error) console.error(error);
    }
  }

  function shouldHandle(textarea) {
    if (!textarea || !textarea.tagName || textarea.tagName !== 'TEXTAREA') return false;
    if (textarea.readOnly || textarea.disabled) return false;
    if (!textarea.classList || !textarea.classList.contains('wiki-edit')) return false;
    var body = document.body;
    if (!body || !body.classList.contains('controller-issues')) return false;
    return body.getAttribute('data-text-formatting') === 'common_mark';
  }

  function clipboardTypes(clipboard) {
    try {
      return Array.prototype.slice.call(clipboard.types || []);
    } catch (e) {
      return [];
    }
  }

  function clipboardHasImage(clipboard) {
    var i;
    if (clipboard.files) {
      for (i = 0; i < clipboard.files.length; i++) {
        if (/^image\//i.test((clipboard.files[i] && clipboard.files[i].type) || '')) return true;
      }
    }
    if (clipboard.items) {
      for (i = 0; i < clipboard.items.length; i++) {
        var item = clipboard.items[i];
        if (item && item.kind === 'file' && /^image\//i.test(item.type || '')) return true;
      }
    }
    return false;
  }

  function isPureImagePaste(clipboard) {
    if (!clipboardHasImage(clipboard)) return false;
    var types = clipboardTypes(clipboard);
    return types.indexOf('text/plain') === -1 && types.indexOf('text/html') === -1;
  }

  function isImageMarkup(html) {
    return /<img[\s/>]/i.test(html || '') ||
      /<picture[\s>]/i.test(html || '') ||
      /<v:imagedata[\s/>]/i.test(html || '');
  }

  function clipboardHasTextTable(html, plain) {
    if (isSpreadsheetHtml(html) || plainIsRectangularTable(plain)) return true;
    var parsed = html ? parseHtmlDocument(normalizeClipboardHtml(html)) : null;
    return !!(parsed && parsed.tableOnly && isUsefulGrid(trimGrid(parsed.grid)));
  }

  function shouldIgnorePaste(hasImageFile, html, plain) {
    if (!hasImageFile && !isImageMarkup(html)) return false;
    return !clipboardHasTextTable(html, plain);
  }

  function readClipboard(clipboard, type) {
    try {
      return clipboard.getData(type) || '';
    } catch (e) {
      return '';
    }
  }

  function convertClipboard(html, plain) {
    var sourceHtml = normalizeClipboardHtml(html || '');
    var parsed = sourceHtml ? parseHtmlDocument(sourceHtml) : null;
    var spreadsheet = isSpreadsheetHtml(html || '');

    if (parsed && parsed.grid && (spreadsheet || parsed.tableOnly)) {
      var fromHtml = trimGrid(parsed.grid);
      if (isUsefulGrid(fromHtml)) return gridToMarkdown(fromHtml);
    }

    if (plainIsRectangularTable(plain)) {
      var fromTsv = trimGrid(parseTsv(plain));
      if (isUsefulGrid(fromTsv) && fromTsv[0].length >= 2) return gridToMarkdown(fromTsv);
    }

    return null;
  }

  function isSpreadsheetHtml(html) {
    if (!html) return false;
    return /schemas-microsoft-com:office:excel/i.test(html) ||
      /Excel\.Sheet/i.test(html) ||
      /schemas-microsoft-com:office:spreadsheet/i.test(html) ||
      /google-sheets-html-origin/i.test(html) ||
      /data-sheets-root\s*=/i.test(html) ||
      /meta[^>]+content\s*=\s*["'][^"']*Microsoft Excel/i.test(html);
  }

  function normalizeClipboardHtml(html) {
    if (!html) return '';
    // Windows puts a Version/StartHTML header before the markup.
    // Keep the <html> document so <head> markers and text outside the table survive.
    var htmlStart = html.search(/<html[\s>]/i);
    if (htmlStart >= 0) return html.slice(htmlStart);
    var fragmentStart = html.indexOf('<!--StartFragment-->');
    if (fragmentStart >= 0) {
      var fragmentEnd = html.indexOf('<!--EndFragment-->');
      return html.slice(fragmentStart + '<!--StartFragment-->'.length, fragmentEnd === -1 ? html.length : fragmentEnd);
    }
    var tagStart = html.search(/</);
    if (tagStart > 0) return html.slice(tagStart);
    return html;
  }

  function parseHtmlDocument(html) {
    if (typeof DOMParser === 'undefined') return null;
    var doc;
    try {
      doc = new DOMParser().parseFromString(html, 'text/html');
    } catch (e) {
      return null;
    }
    if (!doc) return null;

    var tables = doc.querySelectorAll('table');
    if (!tables.length) return null;

    var best = null;
    var bestCount = 0;
    each(tables, function (table) {
      var count = tableCellCount(table);
      if (count > bestCount) {
        best = table;
        bestCount = count;
      }
    });
    if (!best || bestCount < 1) return null;

    return {
      grid: htmlTableToGrid(best),
      tableOnly: isTableOnly(doc.body)
    };
  }

  function each(list, fn) {
    Array.prototype.forEach.call(list || [], fn);
  }

  function directRows(table) {
    var rows = [];
    function collect(parent) {
      each(parent.children, function (child) {
        var name = String(child.tagName || '').toUpperCase();
        if (name === 'TR') rows.push(child);
        else if (name === 'THEAD' || name === 'TBODY' || name === 'TFOOT') collect(child);
      });
    }
    collect(table);
    return rows;
  }

  function directCells(row) {
    var cells = [];
    each(row.children, function (child) {
      var name = String(child.tagName || '').toUpperCase();
      if (name === 'TD' || name === 'TH') cells.push(child);
    });
    return cells;
  }

  function tableCellCount(table) {
    var count = 0;
    directRows(table).forEach(function (row) {
      count += directCells(row).length;
    });
    return count;
  }

  function isTableOnly(body) {
    if (!body) return false;
    var clone = body.cloneNode(true);
    each(clone.querySelectorAll('table, style, script, link, meta, title, colgroup, col'), function (el) {
      el.remove();
    });
    return ((clone.textContent || '').replace(/\s+/g, '')).length === 0;
  }

  function htmlTableToGrid(table) {
    var grid = [];
    directRows(table).forEach(function (row, r) {
      if (!grid[r]) grid[r] = [];
      var c = 0;
      directCells(row).forEach(function (cell) {
        while (grid[r][c] !== undefined) c += 1;
        var text = cellText(cell);
        var colspan = spanCount(cell, 'colspan');
        var rowspan = spanCount(cell, 'rowspan');
        var rr, cc;
        for (rr = 0; rr < rowspan; rr++) {
          if (!grid[r + rr]) grid[r + rr] = [];
          for (cc = 0; cc < colspan; cc++) {
            grid[r + rr][c + cc] = (rr === 0 && cc === 0) ? text : '';
          }
        }
        c += colspan;
      });
    });
    return normalizeGrid(grid);
  }

  function spanCount(cell, attr) {
    var n = parseInt(cell.getAttribute(attr) || '1', 10);
    if (!n || n < 1) return 1;
    if (n > 1000) return 1000;
    return n;
  }

  function cellText(cell) {
    var clone = cell.cloneNode(true);
    each(clone.querySelectorAll('script, style, img, svg, object, embed, canvas, picture'), function (el) {
      el.remove();
    });
    each(clone.querySelectorAll('br'), function (br) {
      br.replaceWith('\n');
    });
    each(clone.querySelectorAll('p, div, li, h1, h2, h3, h4, h5, h6'), function (el) {
      el.prepend(clone.ownerDocument.createTextNode('\n'));
    });
    return cleanCellText(clone.textContent || '');
  }

  function cleanCellText(text) {
    return String(text)
      .replace(/\u00a0/g, ' ')
      .replace(/[ \t]+\n/g, '\n')
      .replace(/\n[ \t]+/g, '\n')
      .replace(/\n{3,}/g, '\n\n')
      .trim();
  }

  function parseTsv(text) {
    if (text == null) return null;
    var source = String(text).replace(/\r\n/g, '\n').replace(/\r/g, '\n');
    if (source.indexOf('\t') === -1) return null;

    var rows = [];
    var row = [];
    var field = '';
    var inQuotes = false;
    var i, ch;

    for (i = 0; i < source.length; i++) {
      ch = source.charAt(i);
      if (inQuotes) {
        if (ch === '"') {
          if (source.charAt(i + 1) === '"') {
            field += '"';
            i += 1;
          } else {
            inQuotes = false;
          }
        } else {
          field += ch;
        }
        continue;
      }
      if (ch === '"' && field === '') {
        inQuotes = true;
        continue;
      }
      if (ch === '\t') {
        row.push(field);
        field = '';
        continue;
      }
      if (ch === '\n') {
        row.push(field);
        rows.push(row);
        row = [];
        field = '';
        continue;
      }
      field += ch;
    }

    if (inQuotes) return naiveTsv(source);

    if (field.length || row.length) {
      row.push(field);
      rows.push(row);
    }
    return rows;
  }

  function naiveTsv(source) {
    var lines = source.split('\n');
    if (lines.length && lines[lines.length - 1] === '') lines.pop();
    return lines.map(function (line) { return line.split('\t'); });
  }

  function plainIsRectangularTable(plain) {
    var grid = parseTsv(plain);
    if (!grid || !isRectangular(grid)) return false;
    if (!grid[0] || grid[0].length < 2) return false;
    return isUsefulGrid(grid);
  }

  function isRectangular(grid) {
    if (!grid || !grid.length || !grid[0]) return false;
    var width = grid[0].length;
    for (var i = 1; i < grid.length; i++) {
      if (!grid[i] || grid[i].length !== width) return false;
    }
    return width > 0;
  }

  function normalizeGrid(grid) {
    var width = 0;
    var r, c;
    for (r = 0; r < grid.length; r++) {
      if (!grid[r]) grid[r] = [];
      if (grid[r].length > width) width = grid[r].length;
    }
    for (r = 0; r < grid.length; r++) {
      for (c = 0; c < width; c++) {
        if (grid[r][c] == null) grid[r][c] = '';
      }
    }
    return grid;
  }

  function trimGrid(grid) {
    if (!grid) return [];
    var rows = grid.map(function (row) { return row.slice(); });
    while (rows.length && rows[rows.length - 1].every(function (cell) {
      return String(cell == null ? '' : cell).trim() === '';
    })) {
      rows.pop();
    }
    if (!rows.length) return [];
    var width = rows.reduce(function (max, row) { return Math.max(max, row.length); }, 0);
    while (width > 0 && rows.every(function (row) {
      return String(row[width - 1] == null ? '' : row[width - 1]).trim() === '';
    })) {
      width -= 1;
    }
    return rows.map(function (row) {
      var next = row.slice(0, width);
      while (next.length < width) next.push('');
      return next;
    });
  }

  function isUsefulGrid(grid) {
    if (!grid || !grid.length || !grid[0] || !grid[0].length) return false;
    if (!(grid.length > 1 || grid[0].length > 1)) return false;
    for (var r = 0; r < grid.length; r++) {
      for (var c = 0; c < grid[r].length; c++) {
        if (String(grid[r][c] == null ? '' : grid[r][c]).trim() !== '') return true;
      }
    }
    return false;
  }

  function gridToMarkdown(grid) {
    var header = grid[0].map(escapeCell);
    var lines = [markdownRow(header), markdownRow(header.map(function () { return '---'; }))];
    for (var r = 1; r < grid.length; r++) {
      lines.push(markdownRow(grid[r].map(escapeCell)));
    }
    return lines.join('\n');
  }

  function markdownRow(cells) {
    return '| ' + cells.join(' | ') + ' |';
  }

  function escapeCell(value) {
    var text = String(value == null ? '' : value).replace(/\t/g, ' ').trim();
    text = text.replace(/\\/g, '\\\\').replace(/\|/g, '\\|');
    text = text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
    text = text.replace(/ {2,}/g, function (spaces) {
      return ' ' + new Array(spaces.length).join('&nbsp;');
    });
    text = text.replace(/\r\n|[\n\r\u2028\u2029]/g, '<br>');
    return text;
  }

  function ensureBlock(before, after, markdown) {
    var prefix = '';
    var suffix = '';
    if (before.length) {
      if (before.slice(-2) === '\n\n') prefix = '';
      else if (before.slice(-1) === '\n') prefix = '\n';
      else prefix = '\n\n';
    }
    if (after.length) {
      if (after.slice(0, 2) === '\n\n') suffix = '';
      else if (after.slice(0, 1) === '\n') suffix = '\n';
      else suffix = '\n\n';
    } else {
      suffix = '\n';
    }
    return prefix + markdown + suffix;
  }

  function insideFence(value, index) {
    var before = String(value || '').slice(0, index || 0);
    var re = /^[ ]{0,3}(```+|~~~+)/gm;
    var fence = null;
    var match;
    while ((match = re.exec(before))) {
      var marker = match[1].charAt(0);
      var len = match[1].length;
      if (!fence) fence = { marker: marker, len: len };
      else if (fence.marker === marker && len >= fence.len) fence = null;
    }
    return !!fence;
  }

  function insertText(textarea, text, start, end) {
    if (typeof textarea.setRangeText === 'function') {
      textarea.setRangeText(text, start, end, 'end');
    } else {
      var value = textarea.value;
      textarea.value = value.slice(0, start) + text + value.slice(end);
      var cursor = start + text.length;
      if (typeof textarea.setSelectionRange === 'function') {
        textarea.setSelectionRange(cursor, cursor);
      }
    }
    try {
      textarea.dispatchEvent(new InputEvent('input', { bubbles: true, inputType: 'insertFromPaste', data: text }));
    } catch (e) {
      textarea.dispatchEvent(new Event('input', { bubbles: true }));
    }
  }

  return {
    install: install,
    convertClipboard: convertClipboard,
    parseTsv: parseTsv,
    gridToMarkdown: gridToMarkdown,
    trimGrid: trimGrid,
    isSpreadsheetHtml: isSpreadsheetHtml,
    ensureBlock: ensureBlock,
    insideFence: insideFence,
    escapeCell: escapeCell,
    shouldIgnorePaste: shouldIgnorePaste
  };
});
