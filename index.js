const express = require("express");
const fs = require("fs");
const { Document, Packer, Paragraph, TextRun, ImageRun, HeadingLevel, AlignmentType } = require("docx");
const axios = require("axios");
const cors = require("cors");

const app = express();
const PORT = process.env.PORT || 3000;

app.use(cors());
app.use(express.json());

app.post("/generar-cv", async (req, res) => {
  try {
    const data = req.body;

    // Descargar imagen si existe
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

    const docSections = [];

    // Foto y Nombre
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

    // Nombre y Puesto
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

    // Datos Personales
    docSections.push(
      new Paragraph({ text: "📄 Datos Personales", heading: HeadingLevel.HEADING_2 }),
      crearLineaInfo("Email", data.email),
      crearLineaInfo("Teléfono", data.telefono),
      crearLineaInfo("Dirección", data.direccion),
      crearLineaInfo("Website", data.website),
      crearLineaInfo("Mensajería", data.mensajeria),
      crearLineaInfo("Género", data.genero),
      crearLineaInfo("Fecha de Nacimiento", data.fechaNacimiento),
      crearLineaInfo("Nacionalidad", data.nacionalidad),
    );

    // Declaración Personal
    docSections.push(
      new Paragraph({ text: "📝 Declaración Personal", heading: HeadingLevel.HEADING_2 }),
      new Paragraph({ text: data.declaracionPersonal || "No especificado", spacing: { after: 300 } })
    );

    // Habilidades
    docSections.push(
      new Paragraph({ text: "🛠️ Habilidades", heading: HeadingLevel.HEADING_2 }),
      new Paragraph({ text: data.habilidades || "No especificado", spacing: { after: 300 } })
    );

    // Experiencia Laboral
    docSections.push(
      new Paragraph({ text: "💼 Experiencia Laboral", heading: HeadingLevel.HEADING_2 })
    );
    (data.experiencias || []).forEach(exp => {
      docSections.push(
        new Paragraph({ text: `• ${exp.puesto || "Puesto no especificado"} en ${exp.empleador || "Empleador no especificado"} (${exp.fecha || "Fecha no especificada"})` }),
        new Paragraph({ text: `  ➔ Sector: ${exp.sector || "No especificado"}` }),
        new Paragraph({ text: `  ➔ Responsabilidades: ${exp.responsabilidades || "No especificado"}`, spacing: { after: 200 } })
      );
    });

    // Formación Académica
    docSections.push(
      new Paragraph({ text: "🎓 Formación Académica", heading: HeadingLevel.HEADING_2 })
    );
    (data.educaciones || []).forEach(edu => {
      docSections.push(
        new Paragraph({ text: `• ${edu.titulo || "Título no especificado"} (${edu.fecha || "Fecha no especificada"})` }),
        new Paragraph({ text: `  ➔ Institución: ${edu.institucion || "No especificado"}` }),
        new Paragraph({ text: `  ➔ Nivel: ${edu.nivel || "No especificado"}` }),
        new Paragraph({ text: `  ➔ Materias: ${edu.materias || "No especificado"}` }),
        new Paragraph({ text: `  ➔ Logros: ${edu.logros || "No especificado"}`, spacing: { after: 200 } })
      );
    });

    // Idiomas
    docSections.push(
      new Paragraph({ text: "🌍 Idiomas", heading: HeadingLevel.HEADING_2 })
    );
    (data.idiomas || []).forEach(id => {
      docSections.push(
        new Paragraph({ text: `• ${id.idioma || "Idioma no especificado"}` }),
        new Paragraph({ text: `  ➔ Comprensión: ${id.comprension || "No especificado"}` }),
        new Paragraph({ text: `  ➔ Hablado: ${id.hablado || "No especificado"}` }),
        new Paragraph({ text: `  ➔ Escrito: ${id.escrito || "No especificado"}` }),
        new Paragraph({ text: `  ➔ Certificado: ${id.certificado || "No especificado"}`, spacing: { after: 200 } })
      );
    });

    // Crear documento
    const doc = new Document({
      sections: [{ properties: {}, children: docSections }],
    });

    const buffer = await Packer.toBuffer(doc);

    res.setHeader("Content-Disposition", "attachment; filename=CV_Completo_Elegante.docx");
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document");

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

// Función de ayuda
function crearLineaInfo(label, value) {
  return new Paragraph({
    spacing: { after: 100 },
    children: [
      new TextRun({ text: label + ": ", bold: true }),
      new TextRun({ text: value || "No especificado" })
    ],
  });
}
