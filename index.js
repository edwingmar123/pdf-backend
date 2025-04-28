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
    let imageUrlAttempted = null;

    if (data.foto_pixar && data.foto_pixar.startsWith("http")) {
      imageUrlAttempted = data.foto_pixar;
      console.log(`Intentando descargar imagen desde: ${imageUrlAttempted}`);
      try {
        const response = await axios.get(imageUrlAttempted, {
          responseType: "arraybuffer",
          timeout: 15000,
          headers: {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
          },
          validateStatus: function (status) {
            return status < 500;
          }
        });

        if (response.status === 200 && response.data) {
            imageBuffer = response.data;
            console.log(`Imagen descargada exitosamente (Status: ${response.status}) desde ${imageUrlAttempted}`);
        } else {
            console.error(`------------------------------------------`);
            console.error(`¡FALLO AL DESCARGAR O PROCESAR IMAGEN! (Respuesta recibida pero no OK)`);
            console.error(`URL: ${imageUrlAttempted}`);
            console.error(`Status Code recibido: ${response.status}`);
            let responseData = response.data;
            try {
              if (responseData instanceof ArrayBuffer) {
                responseData = Buffer.from(responseData).toString('utf-8');
              }
            } catch (e) { /* Ignorar */ }
            console.error("Datos de Respuesta (si aplica):", responseData);
            console.error(`------------------------------------------`);
            imageBuffer = null;
        }
      } catch (error) {
        console.error("------------------------------------------");
        console.error("¡FALLO CRÍTICO AL INTENTAR DESCARGAR LA IMAGEN! (Error en Axios/Red)");
        console.error("URL que falló:", imageUrlAttempted);

        if (error.response) {
          console.error("Status Code:", error.response.status);
          console.error("Headers de Respuesta:", error.response.headers);
          let responseData = error.response.data;
           try {
             if (responseData instanceof ArrayBuffer) {
               responseData = Buffer.from(responseData).toString('utf-8');
             }
           } catch (e) { /* Ignorar */ }
          console.error("Datos de Respuesta:", responseData);
        } else if (error.request) {
          console.error("No se recibió respuesta del servidor.");
        } else {
          console.error("Error en configuración de Axios:", error.message);
        }
        console.error("Código de Error (si existe):", error.code);
        console.error("Mensaje de Error:", error.message);
        console.error("------------------------------------------");
        imageBuffer = null;
      }
    } else {
        if (data.foto_pixar) {
            console.log(`Imagen omitida: La URL proporcionada no comienza con 'http'. URL: ${data.foto_pixar}`);
        } else {
            console.log("Imagen omitida: No se proporcionó URL en foto_pixar.");
        }
    }

    const docSections = [];

    if (imageBuffer) {
      try {
        console.log(`Intentando añadir imagen (tamaño buffer: ${imageBuffer.byteLength}) al DOCX.`);
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
        console.log("Imagen añadida a las secciones del DOCX.");
      } catch(imageError) {
         console.error("------------------------------------------");
         console.error("¡ERROR AL PROCESAR/AÑADIR LA IMAGEN AL DOCX CON LA LIBRERÍA 'docx'!");
         if (imageUrlAttempted) {
             console.error("Imagen descargada desde:", imageUrlAttempted);
         }
         console.error("Error de librería 'docx':", imageError.message);
         console.error("------------------------------------------");
      }
    } else {
      console.log("No se añadió imagen al DOCX porque no se pudo obtener o procesar.");
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

    console.log("Preparando para crear el objeto Document de docx...");
    const doc = new Document({
        sections: [{ properties: {}, children: docSections }],
    });
    console.log("Objeto Document creado exitosamente.");

    console.log("Generando buffer del documento DOCX...");
    const buffer = await Packer.toBuffer(doc);
    console.log("Buffer DOCX generado.");

    res.setHeader(
      "Content-Disposition",
      "attachment; filename=CV_Completo_Elegante.docx"
    );
    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    );

    res.send(buffer);
    console.log("Respuesta DOCX enviada.");

  } catch (error) {
    console.error("------------------------------------------");
    console.error("Error FATAL generando el CV (Bloque Catch Principal):");
    console.error(error);
    console.error("Mensaje:", error.message);
    console.error("------------------------------------------");

    if (!res.headersSent) {
      res.status(500).json({
        message: "Error fatal generando el CV",
        detalle: error.message,
      });
    }
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