const express = require("express");
const axios = require("axios");
const cors = require("cors");
const AdmZip = require('adm-zip');
const { XMLBuilder } = require('fast-xml-parser');

const app = express();
const PORT = process.env.PORT || 3000;

app.use(cors());
app.use(express.json());

// Generador simplificado de PPTX
class SimplePPTXGenerator {
  constructor() {
    this.files = {
      '[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
        <Default Extension="xml" ContentType="application/xml"/>
        <Default Extension="jpeg" ContentType="image/jpeg"/>
        <Default Extension="png" ContentType="image/png"/>
        <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
      </Types>`,
      'ppt/presentation.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      <p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
        <p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>
        <p:sldIdLst><!-- SLIDES WILL BE INSERTED HERE --></p:sldIdLst>
      </p:presentation>`
    };
    this.slideCount = 0;
    this.imageCounter = 1;
    this.zip = new AdmZip();
  }

  async addSlide(title, content, imageData) {
    this.slideCount++;
    const slideId = 256 + this.slideCount;
    
    // Crear archivo de slide
    const slideXml = this.buildSlideXML(title, content, imageData);
    this.files[`ppt/slides/slide${this.slideCount}.xml`] = slideXml;
    
    // Actualizar presentation.xml
    const presXml = this.files['ppt/presentation.xml'];
    this.files['ppt/presentation.xml'] = presXml.replace(
      '<!-- SLIDES WILL BE INSERTED HERE -->',
      `<p:sldId id="${slideId}" r:id="rId${this.slideCount + 1}"/>` + 
      '<!-- SLIDES WILL BE INSERTED HERE -->'
    );
    
    // Agregar imagen si existe
    if (imageData) {
      const ext = imageData.includes('/jpeg') ? 'jpeg' : 'png';
      this.files[`ppt/media/image${this.imageCounter}.${ext}`] = imageData;
      this.imageCounter++;
    }
  }

  buildSlideXML(title, content, imageData) {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
      <p:cSld>
        <p:spTree>
          <p:nvGrpSpPr>
            <p:cNvPr id="1" name=""/>
            <p:cNvGrpSpPr/>
            <p:nvPr/>
          </p:nvGrpSpPr>
          <p:grpSpPr>
            <a:xfrm>
              <a:off x="0" y="0"/>
              <a:ext cx="9144000" cy="6858000"/>
              <a:chOff x="0" y="0"/>
              <a:chExt cx="9144000" cy="6858000"/>
            </a:xfrm>
          </p:grpSpPr>
          
          <!-- Título -->
          <p:sp>
            <p:nvSpPr>
              <p:cNvPr id="2" name="Title"/>
              <p:cNvSpPr/>
              <p:nvPr/>
            </p:nvSpPr>
            <p:spPr>
              <a:xfrm>
                <a:off x="914400" y="457200"/>
                <a:ext cx="7315200" cy="457200"/>
              </a:xfrm>
            </p:spPr>
            <p:txBody>
              <a:bodyPr/>
              <a:lstStyle/>
              <a:p>
                <a:r>
                  <a:rPr lang="en-US" sz="4400" b="1"/>
                  <a:t>${title}</a:t>
                </a:r>
              </a:p>
            </p:txBody>
          </p:sp>
          
          <!-- Contenido -->
          <p:sp>
            <p:nvSpPr>
              <p:cNvPr id="3" name="Content"/>
              <p:cNvSpPr/>
              <p:nvPr/>
            </p:nvSpPr>
            <p:spPr>
              <a:xfrm>
                <a:off x="914400" y="1371600"/>
                <a:ext cx="7315200" cy="4572000"/>
              </a:xfrm>
            </p:spPr>
            <p:txBody>
              <a:bodyPr/>
              <a:lstStyle/>
              ${content.split('\n').map(line => `
                <a:p>
                  <a:r>
                    <a:rPr lang="en-US" sz="3200"/>
                    <a:t>${line}</a:t>
                  </a:r>
                </a:p>
              `).join('')}
            </p:txBody>
          </p:sp>
          
          <!-- Imagen (si existe) -->
          ${imageData ? `
          <p:pic>
            <p:nvPicPr>
              <p:cNvPr id="4" name="Picture"/>
              <p:cNvPicPr/>
              <p:nvPr/>
            </p:nvPicPr>
            <p:blipFill>
              <a:blip r:embed="rIdImage${this.imageCounter}"/>
              <a:stretch>
                <a:fillRect/>
              </a:stretch>
            </p:blipFill>
            <p:spPr>
              <a:xfrm>
                <a:off x="3200400" y="914400"/>
                <a:ext cx="1828800" cy="1371600"/>
              </a:xfrm>
              <a:prstGeom prst="rect">
                <a:avLst/>
              </a:prstGeom>
            </p:spPr>
          </p:pic>
          ` : ''}
        </p:spTree>
      </p:cSld>
    </p:sld>`;
  }

  async generate() {
    // Agregar todos los archivos al ZIP
    for (const [path, content] of Object.entries(this.files)) {
      if (typeof content === 'string') {
        this.zip.addFile(path, Buffer.from(content));
      } else {
        this.zip.addFile(path, content);
      }
    }
    
    return this.zip.toBuffer();
  }
}

app.post("/generar-presentacion", async (req, res) => {
  try {
    const ciudades = req.body;
    const pptx = new SimplePPTXGenerator();

    // Portada
    await pptx.addSlide(
      "✈️ ITINERARIO DE VIAJE", 
      "La guía definitiva para tu aventura\nDestinos seleccionados personalmente para ti"
    );

    // Slides por ciudad
    for (const { ciudad, imagen_url, recomendaciones } of ciudades) {
      let imageData = null;
      
      if (imagen_url && imagen_url.startsWith("http")) {
        try {
          const response = await axios.get(imagen_url, {
            responseType: 'arraybuffer',
            timeout: 10000
          });
          imageData = response.data;
        } catch (error) {
          console.warn(`⚠ No se pudo cargar imagen para ${ciudad}`);
        }
      }
      
      const content = `🌟 ${ciudad.toUpperCase()}\n\n` +
        recomendaciones.map((r, i) => `★ ${r}`).join('\n');
      
      await pptx.addSlide(ciudad, content, imageData);
    }

    // Slide final
    await pptx.addSlide(
      "¡Que tengas un viaje inolvidable!", 
      "El equipo de TravelExperts"
    );

    // Generar y enviar
    const buffer = await pptx.generate();
    
    res.setHeader("Content-Disposition", "attachment; filename=Itinerario_Presentacion.pptx");
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.presentationml.presentation");
    res.send(buffer);
    
  } catch (error) {
    console.error("❌ Error generando presentación:", error);
    res.status(500).json({ error: error.message });
  }
});

app.listen(PORT, '0.0.0.0', () => {
  console.log(`🟢 Servidor en puerto ${PORT}`);
});