const express = require('express');
const fs = require('fs');
const Docxtpl = require('docxtpl');
const PizZip = require('pizzip');

const app = express();
app.use(express.json());

app.post('/generar', (req, res) => {
  const data = req.body;

  const template = fs.readFileSync('plantilla_cv.docx', 'binary');
  const zip = new PizZip(template);
  const doc = new Docxtpl(zip);

  doc.setData(data);
  try {
    doc.render();
  } catch (error) {
    return res.status(500).send('Error al generar el documento');
  }

  const buf = doc.getZip().generate({ type: 'nodebuffer' });

  res.set({
    'Content-Type':
      'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    'Content-Disposition': 'attachment; filename=curriculum.docx',
  });

  res.send(buf);
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`Servidor escuchando en el puerto ${PORT}`);
});