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
  WidthType,
  ExternalHyperlink,
  TabStopPosition,
  TabStopType,
  LineRuleType
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
          text: "✈️ ITINERARIO DE JAPÓN",
          bold: true,
          color: "FFFFFF",
          size: 48,
          font: "Arial",
        })
      ],
      heading: HeadingLevel.HEADING_1,
      alignment: AlignmentType.CENTER,
      spacing: { before: 1200, after: 400 },
      shading: {
        type: ShadingType.GRADIENT,
        color: "1A5276",
        fill: "2E86C1",
        angle: 180
      },
      border: {
        bottom: { style: BorderStyle.DOUBLE, size: 12, color: "F1C40F" }
      }
    }));

    docSections.push(new Paragraph({
      text: "La guía definitiva para explorar la tierra del sol naciente",
      alignment: AlignmentType.CENTER,
      color: "FFFFFF",
      size: 24,
      shading: { type: ShadingType.SOLID, color: "1A5276" },
      spacing: { after: 800 },
    }));

    // Índice interactivo
    const indexItems = ciudades.map((ciudad, i) => ({
      text: `${i + 1}. ${ciudad.ciudad}`,
      link: `#${ciudad.ciudad.replace(/\s+/g, '_')}`
    }));

    docSections.push(new Paragraph({
      text: "ÍNDICE DE DESTINOS",
      bold: true,
      color: "1A5276",
      size: 28,
      border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: "F1C40F" } },
      spacing: { after: 300 },
    }));

    indexItems.forEach(item => {
      docSections.push(new Paragraph({
        children: [
          new ExternalHyperlink({
            children: [new TextRun({
              text: item.text,
              color: "2874A6",
              underline: {},
            })],
            anchor: item.link,
          })
        ],
        spacing: { after: 150 },
      }));
    });

    docSections.push(new Paragraph({ text: "", spacing: { after: 600 } }));

    // Sección para cada ciudad
    for (const { ciudad, imagen_url, recomendaciones } of ciudades) {
      // Ancla para el índice
      docSections.push(new Paragraph({
        children: [new TextRun({ text: "", id: ciudad.replace(/\s+/g, '_') })]
      }));

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
          docSections.push(new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [
              new TextRun({
                text: `[Imagen de ${ciudad}]`,
                color: "95a5a6",
                italic: true,
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

      // Tabla con efecto de tarjetas
      const tableRows = [];
      const chunkSize = 2;
      
      for (let i = 0; i < recomendaciones.length; i += chunkSize) {
        const chunk = recomendaciones.slice(i, i + chunkSize);
        const rowCells = [];
        
        for (const reco of chunk) {
          rowCells.push(new TableCell({
            width: { size: 50, type: WidthType.PERCENTAGE },
            children: [
              new Paragraph({
                children: [
                  new TextRun({
                    text: "★ ",
                    color: "F1C40F",
                    size: 24,
                  }),
                  new TextRun({
                    text: reco,
                    size: 22,
                    color: "212F3C",
                  })
                ],
                spacing: { after: 100 },
              })
            ],
            shading: { type: ShadingType.SOLID, color: "FFFFFF" },
            margins: { top: 100, bottom: 100, left: 100, right: 100 },
            borders: {
              top: { style: BorderStyle.SINGLE, size: 2, color: "AED6F1" },
              bottom: { style: BorderStyle.SINGLE, size: 2, color: "AED6F1" },
              left: { style: BorderStyle.SINGLE, size: 2, color: "AED6F1" },
              right: { style: BorderStyle.SINGLE, size: 2, color: "AED6F1" },
            }
          }));
        }
        
        // Rellenar celdas vacías si es necesario
        while (rowCells.length < chunkSize) {
          rowCells.push(new TableCell({
            children: [new Paragraph("")],
            shading: { type: ShadingType.SOLID, color: "FFFFFF" },
          }));
        }
        
        tableRows.push(new TableRow({
          children: rowCells
        }));
      }

      docSections.push(new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        rows: tableRows,
        margins: { top: 100, bottom: 100 }
      }));

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

    // Sección final con efecto de firma
    docSections.push(new Paragraph({
      text: "Que tengas un viaje inolvidable",
      alignment: AlignmentType.CENTER,
      size: 28,
      color: "27AE60",
      spacing: { before: 400 },
    }));

    docSections.push(new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [
        new TextRun({
          text: "El equipo de JapanTravelExperts",
          color: "7F8C8D",
          italic: true,
        })
      ],
      spacing: { after: 200 },
    }));

    docSections.push(new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [
        new ImageRun({
          data: fs.readFileSync("path/to/signature.png"), // Reemplaza con tu ruta
          transformation: { width: 200, height: 80 },
        })
      ],
      spacing: { after: 400 },
    }));

    const doc = new Document({
      styles: {
        paragraphStyles: [{
          id: "normal",
          name: "Normal",
          run: { 
            font: "Calibri",
            size: 24,
            color: "2C3E50",
          },
          paragraph: {
            spacing: { line: 360, lineRule: LineRuleType.AUTO },
          }
        }]
      },
      sections: [{
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
      }],
    });

    const buffer = await Packer.toBuffer(doc);

    res.setHeader("Content-Disposition", "attachment; filename=Itinerario_Japon.docx");
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
    res.send(buffer);
    console.log("✔ Documento enviado correctamente");
  } catch (error) {
    console.error("❌ Error generando el itinerario:", error);
    res.status(500).json({ message: "Error al generar el itinerario", error: error.message });
  }
});

app.listen(PORT, () => {

  console.log(`🟢 Servidor escuchando en el puerto ${PORT}`);
});