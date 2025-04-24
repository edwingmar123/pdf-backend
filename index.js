const express = require("express");
const fs = require("fs");
const path = require("path");
const Docxtemplater = require("docxtemplater");
const PizZip = require("pizzip");

const app = express();
app.use(express.json());

// Configuración de middleware para validar datos
app.use(express.urlencoded({ extended: true }));
app.use("/generar-cv", (req, res, next) => {
  if (!req.body || Object.keys(req.body).length === 0) {
    return res.status(400).json({ error: "Datos del CV no proporcionados" });
  }
  next();
});

// Función para limpiar y formatear los datos
const limpiarDatos = (data) => {
  const result = {};

  Object.keys(data).forEach((key) => {
    if (typeof data[key] === "string" && data[key].startsWith("=")) {
      result[key] = data[key].substring(1);
    } else if (typeof data[key] === "object" && data[key] !== null) {
      result[key] = JSON.parse(JSON.stringify(data[key]));
    } else {
      result[key] = data[key] || "No especificado";
    }
  });

  return result;
};

app.post("/generar-cv", async (req, res) => {
  try {
    // 1. Limpiar y formatear los datos de entrada
    const rawData = req.body;
    console.log("Datos recibidos:", rawData);

    const data = limpiarDatos(rawData);
    console.log("Datos procesados:", data);

    // 2. Cargar la plantilla
    const templatePath = path.join(
      __dirname,
      "plantilla_cv_final_sin_errores.docx"
    );

    if (!fs.existsSync(templatePath)) {
      throw new Error(`Archivo de plantilla no encontrado en: ${templatePath}. 
        Directorio actual: ${__dirname}. 
        Archivos disponibles: ${fs.readdirSync(__dirname).join(", ")}`);
    }

    const content = fs.readFileSync(templatePath, "binary");
    const zip = new PizZip(content);

    // 3. Configurar docxtemplater (versión actualizada)
    const doc = new Docxtemplater(zip, {
      paragraphLoop: true,
      linebreaks: true,
    });

    // Reemplazar el método obsoleto setData
    doc.compile();
    doc
      .resolveData(data)
      .then(() => {
        doc.render();

        const buffer = doc.getZip().generate({ type: "nodebuffer" });

        res.set({
          "Content-Type":
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
          "Content-Disposition": "attachment; filename=CV_Generado.docx",
        });

        res.send(buffer);
      })
      .catch((error) => {
        console.error("Error al procesar plantilla:", error);
        res.status(500).json({
          error: "Error al procesar la plantilla",
          detalle: error.message,
        });
      });
  } catch (error) {
    console.error("Error general:", error);
    res.status(500).json({
      error: "Error al generar el CV",
      detalle: error.message,
      rutaPlantilla: path.join(
        __dirname,
        "plantilla_cv_final_sin_errores.docx"
      ),
      directorioActual: __dirname,
      archivosDisponibles: fs.readdirSync(__dirname),
    });
  }
});

const PORT = process.env.PORT || 10000;
app.listen(PORT, () => {
  console.log(`Servidor escuchando en el puerto ${PORT}`);
  console.log(`Ruta actual: ${__dirname}`);
  console.log(`Archivos disponibles: ${fs.readdirSync(__dirname).join(", ")}`);
});
