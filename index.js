const express = require("express");
const fs = require("fs");
const path = require("path");
const Docxtemplater = require("docxtemplater");
const PizZip = require("pizzip");
const cors = require("cors");

const app = express();
const PORT = process.env.PORT || 3000;

app.use(cors());
app.use(express.json());

app.post("/generar-cv", async (req, res) => {
  try {
    const rawData = req.body;
    console.log("Datos recibidos:", rawData);

    const data = limpiarDatos(rawData);
    console.log("Datos procesados:", data);

    const templatePath = path.join(__dirname, "plantilla_cv_final_sin_errores.docx");

    if (!fs.existsSync(templatePath)) {
      throw new Error(`Archivo de plantilla no encontrado en: ${templatePath}. 
        Directorio actual: ${__dirname}. 
        Archivos disponibles: ${fs.readdirSync(__dirname).join(", ")}`);
    }

    const content = fs.readFileSync(templatePath, "binary");
    const zip = new PizZip(content);
    const doc = new Docxtemplater(zip, {
      paragraphLoop: true,
      linebreaks: true,
    });

    const templateData = {
      "First name(s) Surname(s)": data.nombre,
      "Replace with house number, street name": data.direccion,
      "[Phone Number]": data.telefono,
      "[State personal website(s)]": data.website,
      "Replace with type of IM service": data.mensajeria,
      "Sex Enter sex": data.genero,
      "Date of birth dd/mm/yyyy": data.fechaNacimiento,
      "Nationality Enter nationality/-ies": data.nacionalidad,
      "Replace with job applied for": data.puesto,
      "personal statement": data.declaracionPersonal,

      experiencias: Array.isArray(data.experiencias) ? data.experiencias.map(exp => ({
        "dates (from - to)": exp.fecha,
        "Replace with occupation or position held": exp.puesto,
        "Replace with employer's name and locality": exp.empleador,
        "Replace with main activities and responsibilities": exp.responsabilidades,
        "Business or sector": exp.sector
      })) : [],

      educaciones: Array.isArray(data.educaciones) ? data.educaciones.map(edu => ({
        "dates (from - to)": edu.fecha,
        "Replace with qualification awarded": edu.titulo,
        "Replace with education or training organisation's name": edu.institucion,
        "EQF (or other) level": edu.nivel,
        "principal subjects covered": edu.materias,
        "Replace with expected achievements": edu.logros
      })) : [],

      idiomas: Array.isArray(data.idiomas) ? data.idiomas.map(idio => ({
        "Replace with language": idio.idioma,
        "Enter level": idio.comprension,
        "Replace with name of language certificate": idio.certificado
      })) : [],

      "Communication skills": data.habilidades || "",
      "Digital skills": "Basado en: " + (data.habilidades || ""),
      foto: data.foto || data.secure_url || ""
    };

    doc.compile();
    await doc.resolveData(templateData);

    doc.render();
    const buffer = doc.getZip().generate({ type: "nodebuffer" });

    res.set({
      "Content-Type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      "Content-Disposition": "attachment; filename=CV_Generado.docx",
    });

    res.send(buffer);
  } catch (error) {
    console.error("Error general:", error);
    res.status(500).json({
      error: "Error al generar el CV",
      detalle: error.message,
      rutaPlantilla: path.join(__dirname, "plantilla_cv_final_sin_errores.docx"),
      directorioActual: __dirname,
      archivosDisponibles: fs.readdirSync(__dirname),
    });
  }
});

app.listen(PORT, () => {
  console.log(`Servidor escuchando en el puerto ${PORT}`);
});

// Función de limpieza (puedes ajustarla según tus necesidades)
function limpiarDatos(datos) {
  return datos || {};
}
