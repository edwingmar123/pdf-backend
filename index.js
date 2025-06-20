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
  ShadingType,
  Table,
  TableRow,
  TableCell,
  WidthType
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

    // Portada con efecto de gradiente
    docSections.push(new Paragraph({
      children: [
        new TextRun({
          text: "✈️ ITINERARIO DE VIAJE",
          bold: true,
          color: "FFFFFF",
          size: 48,
          font: "Arial",
        })
      ],
      heading: HeadingLevel.HEADING_1,
      alignment: AlignmentType.CENTER,
      shading: {
        type: ShadingType.GRADIENT,
        color: "1A5276",
        fill: "2E86C1",
        angle: 180
      },
      spacing: { before: 600, after: 300 },
      border: {
        bottom: { style: BorderStyle.DOUBLE, size: 12, color: "F1C40F" }
      }
    }));

    docSections.push(new Paragraph({
      text: "La guía definitiva para tu aventura",
      alignment: AlignmentType.CENTER,
      color: "FFFFFF",
      size: 28,
      shading: { type: ShadingType.SOLID, color: "1A5276" },
      spacing: { after: 600 },
    }));

    // Sección para cada ciudad
    for (const { ciudad, imagen_url, recomendaciones } of ciudades) {
      // Cabecera con efecto de cinta
      docSections.push(new Paragraph({
        children: [
          new TextRun({
            text: `📍 ${ciudad.toUpperCase()}`,
            bold: true,
            size: 32,
            color: "FFFFFF",
            font: "Arial",
          })
        ],
        heading: HeadingLevel.HEADING_2,
        shading: { 
          type: ShadingType.GRADIENT,
          color: "E74C3C",
          fill: "C0392B",
          angle: 90
        },
        spacing: { before: 400, after: 200 },
        border: { 
          top: { style: BorderStyle.SINGLE, size: 4, color: "F1C40F" },
          bottom: { style: BorderStyle.SINGLE, size: 4, color: "F1C40F" }
        }
      }));

      // Imagen con marco decorativo
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
                border: {
                  top: { style: BorderStyle.SINGLE, size: 8, color: "F1C40F" },
                  bottom: { style: BorderStyle.SINGLE, size: 8, color: "F1C40F" },
                  left: { style: BorderStyle.SINGLE, size: 8, color: "F1C40F" },
                  right: { style: BorderStyle.SINGLE, size: 8, color: "F1C40F" },
                },
              }),
            ],
            spacing: { after: 200 },
          }));
        } catch (error) {
          console.warn(`⚠ No se pudo insertar imagen de ${ciudad}`);
          // Placeholder elegante en lugar de imagen
          docSections.push(new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [
              new TextRun({
                text: `[Imagen de ${ciudad}]`,
                color: "95A5A6",
                italic: true,
                size: 24,
              })
            ],
            shading: { type: ShadingType.SOLID, color: "F9E79F" },
            spacing: { after: 200 },
          }));
        }
      }

      // Recomendaciones en formato tarjeta
      docSections.push(new Paragraph({
        text: "🌟 EXPERIENCIAS IMPERDIBLES",
        bold: true,
        color: "1A5276",
        size: 26,
        shading: { type: ShadingType.SOLID, color: "D6EAF8" },
        spacing: { before: 100, after: 150 },
      }));

      // Lista de recomendaciones con viñetas mejoradas
      (recomendaciones || []).forEach((reco, i) => {
        docSections.push(new Paragraph({
          children: [
            new TextRun({
              text: "★ ",
              color: "F1C40F",
              size: 28,
            }),
            new TextRun({
              text: reco,
              size: 24,
              color: "212F3C",
            })
          ],
          spacing: { after: 120 },
        }));
      });

      // Separador decorativo
      docSections.push(new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [
          new TextRun({
            text: "❖❖❖",
            color: "3498DB",
            size: 28,
          })
        ],
        spacing: { before: 300, after: 300 },
      }));
    }

    // Sección final
    docSections.push(new Paragraph({
      text: "¡Que tengas un viaje inolvidable!",
      alignment: AlignmentType.CENTER,
      size: 28,
      color: "27AE60",
      bold: true,
      spacing: { before: 400 },
    }));

    docSections.push(new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [
        new TextRun({
          text: "El equipo de TravelExperts",
          color: "7F8C8D",
          italic: true,
          size: 22,
        })
      ],
      spacing: { after: 400 },
    }));

    const doc = new Document({
      sections: [
        {
          properties: {
            page: {
              margin: {
                top: 700,
                bottom: 700,
                right: 700,
                left: 700,
              }
            }
          },
          children: docSections,
        },
      ],
    });

    const buffer = await Packer.toBuffer(doc);

    res.setHeader("Content-Disposition", "attachment; filename=Itinerario.docx");
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