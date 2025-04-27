const express = require("express");
const fs = require("fs");
const { Document, Packer, Paragraph, TextRun, ImageRun } = require("docx");
const axios = require("axios");
const cors = require("cors");

const app = express();
const PORT = process.env.PORT || 3000;

app.use(cors());
app.use(express.json());

app.post("/generar-cv", async (req, res) => {
  try {
    const data = req.body;

    let imageBuffer = null;
    if (data.foto_pixar && data.foto_pixar !== "No disponible") {
      try {
        const response = await axios.get(data.foto_pixar, { responseType: "arraybuffer" });
        imageBuffer = response.data;
      } catch (error) {
        console.error("Error descargando imagen:", error.message);
        imageBuffer = null;
      }
    }

    const children = [];

    if (imageBuffer) {
      children.push(
        new Paragraph({
          children: [
            new ImageRun({
              data: imageBuffer,
              transformation: {
                width: 150,
                height: 150,
              },
            }),
          ],
          alignment: "center",
        })
      );
    }

    children.push(
      new Paragraph({
        children: [
          new TextRun({ text: `\nPuesto: ${data.puesto || "No especificado"}`, bold: true }),
          new TextRun({ text: `\nDirección: ${data.direccion || "No especificado"}` }),
          new TextRun({ text: `\nTeléfono: ${data.telefono || "No especificado"}` }),
          new TextRun({ text: `\nWebsite: ${data.website || "No especificado"}` }),
          new TextRun({ text: `\nMensajería: ${data.mensajeria || "No especificado"}` }),
          new TextRun({ text: `\nEmail: ${data.email || "No especificado"}` }),
          new TextRun({ text: `\nGénero: ${data.genero || "No especificado"}` }),
          new TextRun({ text: `\nNacionalidad: ${data.nacionalidad || "No especificado"}` }),
        ],
      })
    );

    const doc = new Document({
      sections: [{ properties: {}, children }],
    });

    const buffer = await Packer.toBuffer(doc);

    res.setHeader("Content-Disposition", "attachment; filename=CV_Generado.docx");
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document");

    res.send(buffer);
  } catch (error) {
    console.error(error);
    res.status(500).json({ message: "Error generando CV", error: error.message });
  }
});

app.listen(PORT, () => {
  console.log(`Servidor escuchando en puerto ${PORT}`);
});
