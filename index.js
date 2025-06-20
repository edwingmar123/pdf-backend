const express = require("express");
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
          text: "🌍 Itinerario de Viaje",
          bold: true,
          color: "1F618D",
          size: 60,
        })
      ],
      heading: HeadingLevel.TITLE,
      alignment: AlignmentType.CENTER,
      spacing: { after: 600 },
    }));

    for (const { ciudad, imagen_url, recomendaciones } of ciudades) {
      // Slide separator
      docSections.push(new Paragraph({
        children: [new TextRun({ text: "".padEnd(70, " ") })],
        border: {
          bottom: { style: BorderStyle.DOUBLE, size: 12, color: "5DADE2" },
        },
        spacing: { after: 300 },
      }));

      // Ciudad - Estilo título
      docSections.push(new Paragraph({
        children: [
          new TextRun({
            text: `📌 ${ciudad}`,
            bold: true,
            size: 48,
            color: "154360",
          }),
        ],
        alignment: AlignmentType.CENTER,
        spacing: { after: 300 },
      }));

      // Imagen de la ciudad
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
            spacing: { after: 300 },
          }));
        } catch (error) {
          console.warn(`⚠ No se pudo insertar imagen para ${ciudad}`);
        }
      }

      // Subtítulo
      docSections.push(new Paragraph({
        children: [
          new TextRun({
            text: "✨ Recomendaciones:",
            bold: true,
            color: "1A5276",
            size: 30,
            underline: {},
          })
        ],
        spacing: { after: 200 },
      }));

      // Lista de recomendaciones
      (recomendaciones || []).forEach((reco, i) => {
        docSections.push(new Paragraph({
          children: [
            new TextRun({
              text: `👉 ${reco}`,
              size: 26,
              color: "283747",
            })
          ],
          spacing: { after: 150 },
        }));
      });

      // Espacio final
      docSections.push(new Paragraph({ text: "", spacing: { after: 500 } }));
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
    console.log("✅ Documento enviado con diseño atractivo");
  } catch (error) {
    console.error("❌ Error generando el itinerario:", error);
    res.status(500).json({ message: "Error al generar el itinerario", error: error.message });
  }
});

app.listen(PORT, () => {
  console.log(`🟢 Servidor escuchando en el puerto ${PORT}`);
});
