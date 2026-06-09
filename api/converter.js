const multer = require('multer');

const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 10 * 1024 * 1024 }
});

function escapeXml(str) {
  return str
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function txtToXml(txtBuffer) {
  const text = txtBuffer.toString('binary');
  const rawLines = text.split('\n');
  if (rawLines[rawLines.length - 1].trim() === '') rawLines.pop();

  const paras = rawLines.map(line => {
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

module.exports = (req, res) => {
  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Método não permitido' });
  }

  upload.single('file')(req, res, (err) => {
    if (err) return res.status(400).json({ error: 'Erro no upload: ' + err.message });
    if (!req.file) return res.status(400).json({ error: 'Nenhum arquivo enviado' });

    const baseName = (req.file.originalname || 'pedidos').replace(/\.txt$/i, '');

    try {
      const xmlContent = txtToXml(req.file.buffer);
      res.setHeader('Content-Type', 'application/xml; charset=utf-8');
      res.setHeader('Content-Disposition', `attachment; filename="${baseName}.xml"`);
      res.send(xmlContent);
    } catch (e) {
      res.status(500).json({ error: 'Falha na conversão: ' + e.message });
    }
  });
};
