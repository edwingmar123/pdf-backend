const express = require("express");
const fs = require("fs");
const {
  Document,
  Packer,
  Paragraph,
  TextRun,
  ImageRun,
  HeadingLevel,
  AlignmentType,
  BorderStyle,
} = require("docx");
const axios = require("axios");
const cors = require("cors");

const app = express();
const PORT = process.env.PORT || 3000;

app.use(cors());
app.use(express.json());

app.post("/generar-itinerario", async (req, res) => {
  try {
    const ciudades = req.body;
    const docSections = [];

    // Título principal
    docSections.push(new Paragraph({
      children: [
        new TextRun({
          text: "🗺️ Itinerario de Viaje",
          bold: true,
          color: "2E86C1",
          size: 48,
        })
      ],
      heading: HeadingLevel.HEADING_1,
      alignment: AlignmentType.CENTER,
      spacing: { after: 400 },
    }));

    for (const { ciudad, imagen_url, recomendaciones } of ciudades) {
      // Separador decorativo
      docSections.push(new Paragraph({
        children: [new TextRun({ text: " ".repeat(50) })],
        border: {
          bottom: { style: BorderStyle.SINGLE, size: 6, color: "2E86C1" },
        },
        spacing: { after: 200 },
      }));

      // Ciudad (título)
      docSections.push(new Paragraph({
        children: [
          new TextRun({
            text: `📍 ${ciudad}`,
            bold: true,
            size: 32,
            color: "1A5276",
          })
        ],
        heading: HeadingLevel.HEADING_2,
        spacing: { after: 200 },
      }));

      // Imagen de ciudad
      if (imagen_url && imagen_url.startsWith("http")) {
        try {
          const imageResp = await axios.get(imagen_url, {
            responseType: "arraybuffer",
            timeout: 10000,
            headers: { 'User-Agent': 'Mozilla/5.0' },
          });

          docSections.push(new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [
              new ImageRun({
                data: imageResp.data,
                transformation: { width: 500, height: 300 },
              }),
            ],
            spacing: { after: 200 },
          }));
        } catch (error) {
          console.warn(`⚠ No se pudo insertar imagen de ${ciudad}`);
        }
      }

      // Subtítulo de recomendaciones
      docSections.push(new Paragraph({
        children: [
          new TextRun({
            text: "⭐ Recomendaciones:",
            bold: true,
            underline: {},
            color: "2874A6",
            size: 28,
          }),
        ],
        spacing: { after: 150 },
      }));

      // Lista de recomendaciones
      (recomendaciones || []).forEach((reco, i) => {
        docSections.push(new Paragraph({
          children: [
            new TextRun({
              text: `${i + 1}. ${reco}`,
              size: 24,
              color: "212F3C",
            })
          ],
          spacing: { after: 100 },
        }));
      });

      docSections.push(new Paragraph({ text: "", spacing: { after: 300 } }));
    }

    const doc = new Document({
      sections: [
        {
          properties: {},
          children: docSections,
        },
      ],
    });

    const buffer = await Packer.toBuffer(doc);

    res.setHeader("Content-Disposition", "attachment; filename=itinerario.docx");
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
    res.send(buffer);
    console.log("✔ Documento enviado correctamente");
  } catch (error) {
    console.error("❌ Error generando el itinerario:", error);
    res.status(500).json({ message: "Error al generar el itinerario", error: error.message });
  }
});

app.listen(PORT, '0.0.0.0', () => {

  console.log(`🟢 Servidor escuchando en el puerto ${PORT}`);
});
