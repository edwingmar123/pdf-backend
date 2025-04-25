app.post("/generar-cv", async (req, res) => {
  try {
    // 1. Limpiar y formatear los datos de entrada
    const rawData = req.body;
    console.log("Datos recibidos:", rawData);

    const data = limpiarDatos(rawData);
    console.log("Datos procesados:", data);

    // 2. Cargar la plantilla
    const templatePath = path.join(__dirname, "plantilla_cv_final_sin_errores.docx");

    if (!fs.existsSync(templatePath)) {
      throw new Error(`Archivo de plantilla no encontrado en: ${templatePath}. 
        Directorio actual: ${__dirname}. 
        Archivos disponibles: ${fs.readdirSync(__dirname).join(", ")}`);
    }

    const content = fs.readFileSync(templatePath, "binary");
    const zip = new PizZip(content);

    // 3. Configurar docxtemplater
    const doc = new Docxtemplater(zip, {
      paragraphLoop: true,
      linebreaks: true,
    });

    // Preparar datos específicos para la plantilla
    const templateData = {
      // Datos personales
      "First name(s) Surname(s)": data.nombre,
      "Replace with house number, street name": data.direccion,
      "[Phone Number]": data.telefono,
      "[State personal website(s)]": data.website,
      "Replace with type of IM service": data.mensajeria,
      "Sex Enter sex": data.genero,
      "Date of birth dd/mm/yyyy": data.fechaNacimiento,
      "Nationality Enter nationality/-ies": data.nacionalidad,
      
      // Puesto y perfil
      "Replace with job applied for": data.puesto,
      "personal statement": data.declaracionPersonal,
      
      // Experiencia laboral
      experiencias: data.experiencias.map(exp => ({
        "dates (from - to)": exp.fecha,
        "Replace with occupation or position held": exp.puesto,
        "Replace with employer's name and locality": exp.empleador,
        "Replace with main activities and responsibilities": exp.responsabilidades,
        "Business or sector": exp.sector
      })),
      
      // Educación
      educaciones: data.educaciones.map(edu => ({
        "dates (from - to)": edu.fecha,
        "Replace with qualification awarded": edu.titulo,
        "Replace with education or training organisation's name": edu.institucion,
        "EQF (or other) level": edu.nivel,
        "principal subjects covered": edu.materias,
        "Replace with expected achievements": edu.logros
      })),
      
      // Idiomas
      idiomas: data.idiomas.map(idio => ({
        "Replace with language": idio.idioma,
        "Enter level": idio.comprension,
        "Replace with name of language certificate": idio.certificado
      })),
      
      // Habilidades
      "Communication skills": data.habilidades,
      "Digital skills": "Basado en: " + data.habilidades,
      
      // Foto (si existe)
      foto: data.foto || data.secure_url
    };

    // Procesar la plantilla
    doc.compile();
    doc.resolveData(templateData)
      .then(() => {
        doc.render();

        const buffer = doc.getZip().generate({ type: "nodebuffer" });

        res.set({
          "Content-Type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
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
      rutaPlantilla: path.join(__dirname, "plantilla_cv_final_sin_errores.docx"),
      directorioActual: __dirname,
      archivosDisponibles: fs.readdirSync(__dirname),
    });
  }
});