const express = require("express");
const fs = require("fs");
const path = require("path");
const Docxtemplater = require("docxtemplater");
const PizZip = require("pizzip");

const app = express();
app.use(express.json());

// Middleware para validar datos básicos
app.use("/generar-cv", (req, res, next) => {
  if (!req.body || Object.keys(req.body).length === 0) {
    return res.status(400).json({ error: "Datos del CV no proporcionados" });
  }
  next();
});

app.post("/generar-cv", (req, res) => {
  try {
    const data = req.body;
    console.log("Datos recibidos:", JSON.stringify(data, null, 2));

    // 1. Validar y completar datos faltantes
    const processedData = {
      nombre: data.nombre || "No especificado",
      email: data.email || "No especificado",
      telefono: data.telefono || "No especificado",
      // ... completar con todos los campos necesarios
      experiencias: Array.isArray(data.experiencias) ? data.experiencias : [],
      educaciones: Array.isArray(data.educaciones) ? data.educaciones : [],
      idiomas: Array.isArray(data.idiomas) ? data.idiomas : [],
    };

    // 2. Ruta del template con verificación
    const templatePath = path.join(
      __dirname,
      "plantilla_cv_final_sin_errores.docx"
    );

    // Verificar si el archivo existe antes de intentar leerlo
    if (!fs.existsSync(templatePath)) {
      throw new Error(
        `El archivo de plantilla no existe en la ruta: ${templatePath}`
      );
    }

    // 3. Procesar el documento
    const content = fs.readFileSync(templatePath, "binary");
    const zip = new PizZip(content);
    const doc = new Docxtemplater(zip, {
      paragraphLoop: true,
      linebreaks: true,
    });

    // Configurar manejo de errores de docxtemplater
    doc.setData(processedData);

    try {
      doc.render();
    } catch (error) {
      console.error("Error al renderizar el documento:", error);
      throw new Error("Error al procesar la plantilla del CV");
    }

    const buffer = doc.getZip().generate({ type: "nodebuffer" });

    // 4. Configurar respuesta
    res.set({
      "Content-Type":
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      "Content-Disposition": "attachment; filename=CV_Generado.docx",
    });

    res.send(buffer);
  } catch (error) {
    console.error("Error generando CV:", error);
    res.status(500).json({
      error: "Error al generar el CV",
      detalle: error.message,
      rutaPlantilla: path.join(
        __dirname,
        "plantilla_cv_final_sin_errores.docx"
      ),
    });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`Servidor escuchando en el puerto ${PORT}`);
  console.log(`Ruta actual: ${__dirname}`);
});
