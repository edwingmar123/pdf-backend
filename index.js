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
      "plantilla_cv_final_con_imagen.docx"
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

    
    const experiencias = (data.experiencias || []).map((exp) => ({
      fecha: exp.fecha,
      puestoExp: exp.puesto,
      empleador: exp.empleador,
      responsabilidades: exp.responsabilidades,
      sector: exp.sector,
    }));

    const educaciones = (data.educaciones || []).map((edu) => ({
      fecha: edu.fecha,
      titulo: edu.titulo,
      institucion: edu.institucion,
      nivel: edu.nivel,
      materias: edu.materias,
      logros: edu.logros,
    }));

    const idiomas = (data.idiomas || []).map((idio) => ({
      idioma: idio.idioma,
      nivelIdioma: `${idio.comprension}/${idio.hablado}/${idio.escrito}`,
      certificado: idio.certificado,
    }));

    const templateData = {
      nombre: data.nombre,
      direccion: data.direccion,
      telefono: data.telefono,
      website: data.website,
      mensajeria: data.mensajeria,
      genero: data.genero,
      fechaNacimiento: data.fechaNacimiento,
      nacionalidad: data.nacionalidad,
      puesto: data.puesto,
      declaracionPersonal: data.declaracionPersonal,
      habilidades: data.habilidades,
      digitales: `Basado en: ${data.habilidades}`,
      experiencias,
      educaciones,
      idiomas,
      image: data.foto_pixar || "", 
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
