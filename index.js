const express = require("express");
const axios = require("axios");
const cors = require("cors");
const AdmZip = require('adm-zip');
const { v4: uuidv4 } = require('uuid');

const app = express();
const PORT = process.env.PORT || 3000;

app.use(cors());
app.use(express.json());

class PPTXGenerator {
  constructor() {
    this.files = {};
    this.slideCount = 0;
    this.imageCounter = 1;
    this.zip = new AdmZip();
    this.initialize();
  }

  initialize() {
    // Archivos mínimos necesarios para un PPTX válido
    this.files['[Content_Types].xml'] = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="jpeg" ContentType="image/jpeg"/>
  <Default Extension="png" ContentType="image/png"/>
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
</Types>`;

    this.files['_rels/.rels'] = `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
</Relationships>`;

    this.files['ppt/_rels/presentation.xml.rels'] = `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/presProps" Target="presProps.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/viewProps" Target="viewProps.xml"/>
  <!-- SLIDE RELATIONSHIPS -->
</Relationships>`;

    this.files['ppt/presentation.xml'] = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:sldMasterIdLst>
    <p:sldMasterId id="2147483648" r:id="rId1"/>
  </p:sldMasterIdLst>
  <p:sldIdLst><!-- SLIDES WILL BE INSERTED HERE --></p:sldIdLst>
  <p:sldSz cx="9144000" cy="6858000"/>
  <p:notesSz cx="6858000" cy="9144000"/>
</p:presentation>`;

    this.files['ppt/slideMasters/slideMaster1.xml'] = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr>
        <p:cNvPr id="1" name=""/>
        <p:cNvGrpSpPr/>
        <p:nvPr/>
      </p:nvGrpSpPr>
      <p:grpSpPr>
        <a:xfrm xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <a:off x="0" y="0"/>
          <a:ext cx="0" cy="0"/>
          <a:chOff x="0" y="0"/>
          <a:chExt cx="0" cy="0"/>
        </a:xfrm>
      </p:grpSpPr>
    </p:spTree>
  </p:cSld>
</p:sldMaster>`;

    this.files['ppt/presProps.xml'] = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentationPr xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>`;

    this.files['ppt/viewProps.xml'] = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:viewPr xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>`;
  }

  async addSlide(title, content, imageData) {
    this.slideCount++;
    const slideId = 256 + this.slideCount;
    const slideUuid = uuidv4();
    
    // Crear archivo de slide
    const slideXml = this.buildSlideXML(title, content, imageData);
    this.files[`ppt/slides/slide${this.slideCount}.xml`] = slideXml;
    
    // Crear archivo .rels para el slide
    const relsXml = `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
  ${imageData ? `<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image${this.imageCounter}.${imageData.type}"/>` : ''}
</Relationships>`;
    this.files[`ppt/slides/_rels/slide${this.slideCount}.xml.rels`] = relsXml;
    
    // Actualizar presentation.xml
    const presXml = this.files['ppt/presentation.xml'];
    this.files['ppt/presentation.xml'] = presXml.replace(
      '<!-- SLIDES WILL BE INSERTED HERE -->',
      `<p:sldId id="${slideId}" r:id="${slideUuid}"/>` + 
      '<!-- SLIDES WILL BE INSERTED HERE -->'
    );
    
    // Actualizar presentation.xml.rels
    const presRels = this.files['ppt/_rels/presentation.xml.rels'];
    this.files['ppt/_rels/presentation.xml.rels'] = presRels.replace(
      '<!-- SLIDE RELATIONSHIPS -->',
      `<Relationship Id="${slideUuid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${this.slideCount}.xml"/>` +
      '<!-- SLIDE RELATIONSHIPS -->'
    );
    
    // Agregar imagen si existe
    if (imageData) {
      this.files[`ppt/media/image${this.imageCounter}.${imageData.type}`] = imageData.buffer;
      this.imageCounter++;
    }
  }

  buildSlideXML(title, content, imageData) {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" 
       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" 
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
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
        <p:spPr/>
        <p:txBody>
          <a:bodyPr/>
          <a:lstStyle/>
          <a:p>
            <a:r>
              <a:rPr lang="en-US" sz="4400" b="1"/>
              <a:t>${this.escapeXml(title)}</a:t>
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
        <p:spPr/>
        <p:txBody>
          <a:bodyPr/>
          <a:lstStyle/>
          ${content.split('\n').map(line => `
            <a:p>
              <a:r>
                <a:rPr lang="en-US" sz="3200"/>
                <a:t>${this.escapeXml(line)}</a:t>
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
          <a:blip r:embed="rId2"/>
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

  escapeXml(unsafe) {
    return unsafe.replace(/[<>&'"]/g, c => {
      switch (c) {
        case '<': return '&lt;';
        case '>': return '&gt;';
        case '&': return '&amp;';
        case '\'': return '&apos;';
        case '"': return '&quot;';
        default: return c;
      }
    });
  }

  async generate() {
    // Agregar todos los archivos al ZIP
    for (const [path, content] of Object.entries(this.files)) {
      if (Buffer.isBuffer(content)) {
        this.zip.addFile(path, content);
      } else {
        this.zip.addFile(path, Buffer.from(content, 'utf8'));
      }
    }
    
    return this.zip.toBuffer();
  }
}

app.post("/generar-presentacion", async (req, res) => {
  try {
    const ciudades = req.body;
    const pptx = new PPTXGenerator();

    // Portada
    await pptx.addSlide(
      "✈️ ITINERARIO DE VIAJE", 
      "La guía definitiva para tu aventura\nDestinos seleccionados personalmente para ti",
      null
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
          
          // Determinar tipo de imagen
          let type = 'jpg';
          if (imagen_url.toLowerCase().endsWith('.png')) type = 'png';
          if (imagen_url.toLowerCase().endsWith('.gif')) type = 'gif';
          
          imageData = {
            buffer: Buffer.from(response.data),
            type: type
          };
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
      "El equipo de TravelExperts",
      null
    );

    // Generar y enviar
    const buffer = await pptx.generate();
    
    res.setHeader("Content-Disposition", "attachment; filename=Itinerario_Presentacion.pptx");
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.presentationml.presentation");
    res.send(buffer);
    console.log("✔ Presentación generada y enviada");
    
  } catch (error) {
    console.error("❌ Error generando presentación:", error);
    res.status(500).json({ error: error.message });
  }
});

// Endpoint de prueba
app.get('/', (req, res) => {
  res.send('Servidor de generación de presentaciones activo');
});

app.listen(PORT, '0.0.0.0', () => {
  console.log(`🟢 Servidor en puerto ${PORT}`);
});