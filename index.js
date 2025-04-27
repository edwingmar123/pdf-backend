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
} = require("docx");
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
    if (data.foto_pixar && data.foto_pixar.startsWith("http")) {
      try {
        console.log(`Intentando descargar imagen desde: ${data.foto_pixar}`);
        const response = await axios.get(data.foto_pixar, {
          responseType: "arraybuffer",
          timeout: 10000,
          headers: {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
          }
        });

        if (response.status === 200 && response.data) {
            imageBuffer = response.data;
            console.log("Imagen descargada exitosamente.");
        } else {
            console.error(`Error descargando imagen: Status ${response.status}, Respuesta recibida pero sin datos válidos.`);
            imageBuffer = null;
        }
      } catch (error) {
        console.error("------------------------------------------");
        console.error("¡FALLO AL DESCARGAR LA IMAGEN!");
        console.error("URL que falló:", data.foto_pixar);

        if (error.response) {
          console.error("Status Code:", error.response.status);
          console.error("Headers de Respuesta:", error.response.headers);
          let responseData = error.response.data;
           try {
             if (responseData instanceof ArrayBuffer) {
               responseData = Buffer.from(responseData).toString('utf-8');
             }
           } catch (e) { /* Ignorar si no se puede convertir */ }
          console.error("Datos de Respuesta:", responseData);

        } else if (error.request) {
          console.error("No se recibió respuesta del servidor.");
        } else {
          console.error("Error en configuración de Axios:", error.message);
        }
        console.error("Código de Error (si existe):", error.code);
        console.error("------------------------------------------");

        imageBuffer = null;
      }
    } else {
        if (data.foto_pixar) {
            console.log(`Imagen omitida: La URL '${data.foto_pixar}' no comienza con 'http'.`);
        } else {
            console.log("Imagen omitida: No se proporcionó URL en foto_pixar.");
        }
    }

    const docSections = [];

    if (imageBuffer) {
      docSections.push(
        new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [
            new ImageRun({
              data: imageBuffer,
              transformation: { width: 150, height: 150 },
            }),
          ],
        })
      );
    }

    docSections.push(
      new Paragraph({
        text: data.nombre || "Nombre No especificado",
        heading: HeadingLevel.HEADING_1,
        alignment: AlignmentType.CENTER,
        spacing: { after: 200 },
      }),
      new Paragraph({
        text: data.puesto || "Puesto No especificado",
        alignment: AlignmentType.CENTER,
        spacing: { after: 400 },
      })
    );

    docSections.push(
      new Paragraph({
        text: "📄 Datos Personales",
        heading: HeadingLevel.HEADING_2,
      }),
      crearLineaInfo("Email", data.email),
      crearLineaInfo("Teléfono", data.telefono),
      crearLineaInfo("Dirección", data.direccion),
      crearLineaInfo("Website", data.website),
      crearLineaInfo("Mensajería", data.mensajeria),
      crearLineaInfo("Género", data.genero),
      crearLineaInfo("Fecha de Nacimiento", data.fechaNacimiento),
      crearLineaInfo("Nacionalidad", data.nacionalidad)
    );

    docSections.push(
      new Paragraph({
        text: "📝 Declaración Personal",
        heading: HeadingLevel.HEADING_2,
      }),
      new Paragraph({
        text: data.declaracionPersonal || "No especificado",
        spacing: { after: 300 },
      })
    );

    docSections.push(
      new Paragraph({
        text: "🛠️ Habilidades",
        heading: HeadingLevel.HEADING_2,
      }),
      new Paragraph({
        text: data.habilidades || "No especificado",
        spacing: { after: 300 },
      })
    );

    const experiencias = Array.isArray(data.experiencias) ? data.experiencias : [];
    const educaciones = Array.isArray(data.educaciones) ? data.educaciones : [];
    const idiomas = Array.isArray(data.idiomas) ? data.idiomas : [];

    docSections.push(
      new Paragraph({
        text: "💼 Experiencia Laboral",
        heading: HeadingLevel.HEADING_2,
      })
    );
    experiencias.forEach((exp) => {
      docSections.push(
        new Paragraph({
          text: `• ${exp.puesto || "Puesto no especificado"} en ${
            exp.empleador || "Empleador no especificado"
          } (${exp.fecha || "Fecha no especificada"})`,
        }),
        new Paragraph({
          text: `  ➔ Sector: ${exp.sector || "No especificado"}`,
        }),
        new Paragraph({
          text: `  ➔ Responsabilidades: ${
            exp.responsabilidades || "No especificado"
          }`,
          spacing: { after: 200 },
        })
      );
    });

    docSections.push(
      new Paragraph({
        text: "🎓 Formación Académica",
        heading: HeadingLevel.HEADING_2,
      })
    );
    educaciones.forEach((edu) => {
      docSections.push(
        new Paragraph({
          text: `• ${edu.titulo || "Título no especificado"} (${
            edu.fecha || "Fecha no especificada"
          })`,
        }),
        new Paragraph({
          text: `  ➔ Institución: ${edu.institucion || "No especificado"}`,
        }),
        new Paragraph({ text: `  ➔ Nivel: ${edu.nivel || "No especificado"}` }),
        new Paragraph({
          text: `  ➔ Materias: ${edu.materias || "No especificado"}`,
        }),
        new Paragraph({
          text: `  ➔ Logros: ${edu.logros || "No especificado"}`,
          spacing: { after: 200 },
        })
      );
    });

    docSections.push(
      new Paragraph({ text: "🌍 Idiomas", heading: HeadingLevel.HEADING_2 })
    );
    idiomas.forEach((id) => {
      docSections.push(
        new Paragraph({ text: `• ${id.idioma || "Idioma no especificado"}` }),
        new Paragraph({
          text: `  ➔ Comprensión: ${id.comprension || "No especificado"}`,
        }),
        new Paragraph({
          text: `  ➔ Hablado: ${id.hablado || "No especificado"}`,
        }),
        new Paragraph({
          text: `  ➔ Escrito: ${id.escrito || "No especificado"}`,
        }),
        new Paragraph({
          text: `  ➔ Certificado: ${id.certificado || "No especificado"}`,
          spacing: { after: 200 },
        })
      );
    });

    const doc = new Document({
      sections: [{ properties: {}, children: docSections }],
    });

    const buffer = await Packer.toBuffer(doc);

    res.setHeader(
      "Content-Disposition",
      "attachment; filename=CV_Completo_Elegante.docx"
    );
    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    );

    res.send(buffer);
  } catch (error) {
    console.error("Error generando el CV:", error);
    res.status(500).json({
      message: "Error generando el CV",
      detalle: error.message,
    });
  }
});

app.listen(PORT, () => {
  console.log(`Servidor escuchando en el puerto ${PORT}`);
});

function crearLineaInfo(label, value) {
  return new Paragraph({
    spacing: { after: 100 },
    children: [
      new TextRun({ text: label + ": ", bold: true }),
      new TextRun({ text: value || "No especificado" }),
    ],
  });
}