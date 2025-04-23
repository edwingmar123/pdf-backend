const express = require('express');
const fs = require('fs');
const path = require('path');
const Docxtpl = require('docxtpl');
const PizZip = require('pizzip');
const bodyParser = require('body-parser');

const app = express();
const PORT = process.env.PORT || 3000;

app.use(bodyParser.json({ limit: '10mb' }));

app.post('/generar-cv', async (req, res) => {
  try {
    const datos = req.body;

    const templatePath = path.join(__dirname, 'plantilla_cv.docx');
    const content = fs.readFileSync(templatePath, 'binary');
    const zip = new PizZip(content);
    const doc = new Docxtpl(zip);

    doc.setData(datos);
    doc.render();

    const buf = doc.getZip().generate({ type: 'nodebuffer' });

    res.set({
      'Content-Type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      'Content-Disposition': 'attachment; filename=curriculum.docx',
    });
    res.send(buf);
  } catch (error) {
    console.error(error);
    res.status(500).send({ error: 'Error generando el CV' });
  }
});

app.get('/', (req, res) => {
  res.send('API de generación de CV está activa');
});

app.listen(PORT, () => {
  console.log(Servidor escuchando en el puerto ${PORT});
});