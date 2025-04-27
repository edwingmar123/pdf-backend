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

    const templatePath = path.join(
      __dirname,
      "plantilla_cv_super_elegante.docx"
    );

    if (!fs.existsSync(templatePath)) {
      throw new Error("No se encontró la plantilla DOCX en el servidor.");
    }

    const content = fs.readFileSync(templatePath, "binary");
    const zip = new PizZip(content);

    const imageOpts = {
      centered: false,
      getImage: async function (tagValue) {
        const response = await axios.get(tagValue, {
          responseType: "arraybuffer",
        });
        return response.data;
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

    // Ahora solo enviamos lo que quieres
    const templateData = {
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

    await doc.renderAsync(templateData);

    const buffer = doc.getZip().generate({ type: "nodebuffer" });

    res.set({
      "Content-Type":
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      "Content-Disposition": "attachment; filename=CV_Generado.docx",
    });

    res.send(buffer);
  } catch (error) {
    console.error("Error al generar el CV:", error);
    res.status(500).json({
      error: "Error al generar el CV",
      detalle: error.message,
    });
  }
});

app.listen(PORT, () => {
  console.log(`Servidor escuchando en el puerto ${PORT}`);
});
