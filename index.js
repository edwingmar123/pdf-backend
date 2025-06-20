const express = require("express");
const axios = require("axios");
const cors = require("cors");
const PptxGenJS = require("pptxgenjs");

const app = express();
const PORT = process.env.PORT || 3000;

app.use(cors());
app.use(express.json());

// Paleta de colores profesional
const COLOR_PALETTE = {
  primary: "#2E86C1",
  secondary: "#F1C40F",
  accent: "#E74C3C",
  dark: "#1A5276",
  light: "#D6EAF8",
  background: "#FFFFFF",
  text: "#212F3C"
};

app.post("/generar-presentacion", async (req, res) => {
  try {
    const ciudades = req.body;
    const pptx = new PptxGenJS();

    // Configuración global estilo Canva
    pptx.title = "Itinerario de Viaje";
    pptx.author = "TravelExperts";
    pptx.company = "JapanTravel";
    pptx.layout = "LAYOUT_WIDE"; // 16:9
    pptx.theme = {
      bodyFontFace: "Montserrat",
      headingFontFace: "Montserrat",
      bodyFontSize: 18,
      titleFontSize: 36,
      titleColor: COLOR_PALETTE.primary,
      textColor: COLOR_PALETTE.text,
      background: COLOR_PALETTE.background
    };

    // ===== PORTADA =====
    const cover = pptx.addSlide();
    cover.background = { color: COLOR_PALETTE.primary };
    cover.addText("✈️ ITINERARIO DE VIAJE", {
      x: 0.5,
      y: 1.5,
      w: "90%",
      h: 1.5,
      fontSize: 48,
      bold: true,
      color: COLOR_PALETTE.secondary,
      align: "center",
      fontFace: "Montserrat",
      shadow: { type: "outer", opacity: 0.5, blur: 3, offset: 2 },
      effect: "growShrink"
    });
    
    cover.addText("La guía definitiva para tu aventura", {
      x: 0.5,
      y: 3.5,
      w: "90%",
      fontSize: 28,
      color: "FFFFFF",
      align: "center",
      fontFace: "Montserrat",
      effect: "fade"
    });
    
    cover.addText("Destinos seleccionados personalmente para ti", {
      x: 0.5,
      y: 6,
      w: "90%",
      fontSize: 22,
      color: COLOR_PALETTE.light,
      align: "center",
      effect: "fade",
      effectDelay: 1000
    });

    // ===== SLIDE DE ÍNDICE =====
    const indexSlide = pptx.addSlide();
    indexSlide.background = { color: COLOR_PALETTE.background };
    
    // Título con efecto
    indexSlide.addText("DESTINOS", {
      x: 0.5,
      y: 0.5,
      w: "90%",
      h: 1,
      fontSize: 36,
      bold: true,
      color: COLOR_PALETTE.primary,
      align: "center",
      effect: "fly",
      effectDir: "b"
    });
    
    // Línea decorativa
    indexSlide.addShape(pptx.ShapeType.line, {
      x: 1,
      y: 1.6,
      w: 8,
      h: 0,
      line: { color: COLOR_PALETTE.secondary, width: 3 },
      effect: "wipe"
    });
    
    // Lista de ciudades con efecto escalonado
    ciudades.forEach((ciudad, i) => {
      const yPos = 2 + (i * 0.8);
      
      // Bullet decorativo
      indexSlide.addShape(pptx.ShapeType.roundRect, {
        x: 1.5,
        y: yPos,
        w: 0.4,
        h: 0.4,
        fill: { color: COLOR_PALETTE.accent },
        line: { color: COLOR_PALETTE.dark, width: 1 },
        effect: "appear"
      });
      
      // Nombre de ciudad
      indexSlide.addText(ciudad.ciudad, {
        x: 2.2,
        y: yPos - 0.1,
        w: 6,
        fontSize: 24,
        bold: true,
        color: COLOR_PALETTE.text,
        effect: "appear",
        effectDelay: 200 * (i + 1)
      });
      
      // Línea decorativa
      indexSlide.addShape(pptx.ShapeType.line, {
        x: 2,
        y: yPos + 0.5,
        w: 6,
        h: 0,
        line: { color: COLOR_PALETTE.light, width: 1, dashType: "dash" },
        effect: "appear",
        effectDelay: 200 * (i + 1)
      });
    });

    // ===== SLIDES POR CIUDAD =====
    for (const [index, { ciudad, imagen_url, recomendaciones }] of ciudades.entries()) {
      const slide = pptx.addSlide();
      
      // Fondo con gradiente moderno
      slide.background = { 
        color: { type: "gradient", 
        stops: [
          { color: COLOR_PALETTE.primary, position: 0 },
          { color: COLOR_PALETTE.light, position: 100 }
        ],
        angle: 45 }
      };
      
      // Marco decorativo
      slide.addShape(pptx.ShapeType.rect, {
        x: 0.5,
        y: 0.5,
        w: 9,
        h: 6.5,
        fill: { color: "FFFFFF", transparency: 90 },
        line: { color: COLOR_PALETTE.secondary, width: 4 },
        shadow: { type: "outer", blur: 12, offset: 4, opacity: 0.3 },
        effect: "zoom"
      });
      
      // Título con efecto
      slide.addText(ciudad, {
        x: 1,
        y: 0.7,
        w: 8,
        h: 0.8,
        fontSize: 36,
        bold: true,
        color: COLOR_PALETTE.dark,
        align: "center",
        fontFace: "Montserrat",
        effect: "float",
        effectDir: "up"
      });
      
      // Imagen con efectos modernos
      if (imagen_url && imagen_url.startsWith("http")) {
        try {
          const imageResp = await axios.get(imagen_url, {
            responseType: "arraybuffer",
            timeout: 10000,
          });
          
          slide.addImage({
            data: imageResp.data,
            x: 1.5,
            y: 1.8,
            w: 4.5,
            h: 3,
            sizing: { type: "cover" },
            rounding: true,
            border: { pt: 2, color: "FFFFFF" },
            shadow: { type: "outer", blur: 8, opacity: 0.5 },
            hyperlink: { url: imagen_url },
            effect: "zoom",
            effectStart: "with-previous"
          });
        } catch (error) {
          console.warn(`⚠ No se pudo insertar imagen de ${ciudad}`);
          slide.addText(`[Imagen de ${ciudad}]`, {
            x: 1.5,
            y: 2.5,
            w: 4.5,
            h: 1,
            fontSize: 18,
            italic: true,
            color: COLOR_PALETTE.accent,
            align: "center",
            fill: { color: COLOR_PALETTE.light }
          });
        }
      }
      
      // Título recomendaciones
      slide.addText("🌟 EXPERIENCIAS IMPERDIBLES", {
        x: 6.5,
        y: 1.8,
        w: 2.5,
        fontSize: 20,
        bold: true,
        color: COLOR_PALETTE.primary,
        effect: "fly",
        effectDir: "r"
      });
      
      // Lista de recomendaciones con efectos escalonados
      (recomendaciones || []).forEach((reco, i) => {
        // Ícono decorativo
        slide.addShape(pptx.ShapeType.star5, {
          x: 6.2,
          y: 2.4 + (i * 0.7),
          w: 0.4,
          h: 0.4,
          fill: { color: COLOR_PALETTE.secondary },
          effect: "appear",
          effectDelay: 200 * i
        });
        
        // Texto de recomendación
        slide.addText(reco, {
          x: 6.7,
          y: 2.35 + (i * 0.7),
          w: 3,
          fontSize: 18,
          color: COLOR_PALETTE.text,
          bullet: true,
          effect: "appear",
          effectDelay: 200 * i
        });
      });
      
      // Pie de página con progreso
      slide.addText(`Destino ${index + 1} de ${ciudades.length}`, {
        x: 0.5,
        y: 6.8,
        w: "90%",
        fontSize: 14,
        color: COLOR_PALETTE.dark,
        align: "center",
        effect: "fade"
      });
      
      // Transición entre slides
      slide.transition = { 
        type: "push", 
        direction: "left", 
        duration: 500 
      };
    }

    // ===== SLIDE FINAL =====
    const closingSlide = pptx.addSlide();
    closingSlide.background = { color: COLOR_PALETTE.dark };
    
    // Mensaje principal
    closingSlide.addText("¡Que tengas un viaje inolvidable!", {
      x: 0.5,
      y: 2,
      w: "90%",
      h: 1.5,
      fontSize: 42,
      bold: true,
      color: COLOR_PALETTE.secondary,
      align: "center",
      fontFace: "Montserrat",
      effect: "growShrink"
    });
    
    // Firma
    closingSlide.addText("El equipo de TravelExperts", {
      x: 0.5,
      y: 4,
      w: "90%",
      fontSize: 28,
      color: COLOR_PALETTE.light,
      align: "center",
      italic: true,
      effect: "fade",
      effectDelay: 1000
    });
    
    // Elementos decorativos
    closingSlide.addShape(pptx.ShapeType.heart, {
      x: 4.5,
      y: 3.5,
      w: 1,
      h: 1,
      fill: { color: COLOR_PALETTE.accent },
      rotation: 45,
      effect: "spin"
    });
    
    closingSlide.addShape(pptx.ShapeType.heart, {
      x: 4.5,
      y: 3.5,
      w: 1,
      h: 1,
      fill: { color: "transparent" },
      line: { color: COLOR_PALETTE.secondary, width: 2 },
      rotation: -45,
      effect: "spin",
      effectDelay: 500
    });

    // Generar el PPTX en memoria
    const buffer = await pptx.write({ outputType: "nodebuffer" });

    // Enviar al cliente
    res.setHeader("Content-Disposition", "attachment; filename=Itinerario_Presentacion.pptx");
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.presentationml.presentation");
    res.send(buffer);
    console.log("✔ Presentación enviada correctamente");
    
  } catch (error) {
    console.error("❌ Error generando la presentación:", error);
    res.status(500).json({ message: "Error al generar la presentación", error: error.message });
  }
});

app.listen(PORT, '0.0.0.0', () => {
  console.log(`🟢 Servidor escuchando en el puerto ${PORT}`);
});
