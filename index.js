const express = require("express");
const fs = require("fs");
const path = require("path");
const Docxtemplater = require("docxtemplater");
const PizZip = require("pizzip");
const ImageModule = require("docxtemplater-image-module-free");
const axios = require("axios");
const cors = require("cors");

const app = express();
const PORT = process.env.PORT || 3000;

app.use(cors());
app.use(express.json());

app.post("/generar-cv", async (req, res) => {
  try {
    const data = req.body;

    const templatePath = path.join(__dirname, "plantilla_cv_final_simple_con_nombre.docx");

    if (!fs.existsSync(templatePath)) {
      throw new Error("No se encontró la plantilla DOCX en el servidor.");
    }

    const content = fs.readFileSync(templatePath, "binary");
    const zip = new PizZip(content);

    // Configuración para cargar imágenes de forma segura
    const imageOpts = {
      centered: false,
      getImage: async function (tagValue) {
        try {
          const response = await axios.get(tagValue, {
            responseType: "arraybuffer",
          });
          return response.data;
        } catch (error) {
          console.error("Error cargando imagen:", error.message);
          // Si no se puede cargar la imagen, devolver un buffer vacío
          return Buffer.from("");
        }
      },
      getSize: function () {
        return [150, 150];
      },
    };

    const imageModule = new ImageModule(imageOpts);

    const doc = new Docxtemplater(zip, {
      paragraphLoop: true,
      linebreaks: true,
      modules: [imageModule],
    });

    // Preparar datos con valores por defecto
    const templateData = {
      nombre: data.nombre || "No especificado",
      FOTO_PIXAR: data.foto_pixar || "",
      direccion: data.direccion || "No especificado",
      telefono: data.telefono || "No especificado",
      website: data.website || "No especificado",
      mensajeria: data.mensajeria || "No especificado",
      email: data.email || "No especificado",
      genero: data.genero || "No especificado",
      nacionalidad: data.nacionalidad || "No especificado",
      puesto: data.puesto || "No especificado",
    };

    // Renderizar el documento
    await doc.renderAsync(templateData);

    const buffer = doc.getZip().generate({ type: "nodebuffer" });

    res.set({
      "Content-Type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      "Content-Disposition": "attachment; filename=CV_Generado.docx",
    });

    res.send(buffer);
  } catch (error) {
    console.error("Error al generar el CV:", error);

    // Manejo de error específico si es de plantilla
    if (error.properties && error.properties.errors) {
      return res.status(400).json({
        error: "Error de plantilla",
        detalles: error.properties.errors.map(err => err.explanation),
      });
    }

    res.status(500).json({
      error: "Error al generar el CV",
      detalle: error.message,
    });
  }
});

app.listen(PORT, () => {
  console.log(`Servidor escuchando en el puerto ${PORT}`);
});
