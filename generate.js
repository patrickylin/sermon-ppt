const PptxGenJS = require("pptxgenjs");
const fs = require("fs");

const data = JSON.parse(fs.readFileSync("sermon-data.json", "utf-8"));

const theme = {
  navy: "1E2761",
  white: "FFFFFF",
  lightBlue: "CADCFC",
  gray: "363636"
};

let pres = new PptxGenJS();

/* ---------- TEMPLATES ---------- */

function addTitleSlide() {
  const slide = pres.addSlide();
  slide.background = { color: theme.navy };

  slide.addText(data.title, {
    x: 0.5, y: 1.8, w: 9,
    fontSize: 44, bold: true, color: theme.white, align: "center"
  });

  slide.addText(data.subtitle, {
    x: 0.5, y: 3.2, w: 9,
    fontSize: 24, color: theme.lightBlue, align: "center"
  });

  slide.addText(`${data.preacher} | ${data.date}`, {
    x: 0.5, y: 4.5, w: 9,
    fontSize: 16, color: theme.lightBlue, align: "center"
  });
}

function addScriptureSlides() {
  data.scripture.forEach(text => {
    const slide = pres.addSlide();

    slide.addText(data.subtitle, {
      x: 0.5, y: 0.4, w: 9,
      fontSize: 28, bold: true, align: "center"
    });

    slide.addText(text, {
      x: 0.8, y: 1.1, w: 8.4,
      fontSize: 18, lineSpacing: 28
    });
  });
}

function addOutlineSlide() {
  const slide = pres.addSlide();

  slide.addText("信息大綱", {
    x: 0.5, y: 0.5, w: 9,
    fontSize: 36, bold: true, align: "center"
  });

  slide.addText(
    data.outline.map(item => ({
      text: item,
      options: { bullet: true, breakLine: true }
    })),
    { x: 1.5, y: 1.8, w: 7, fontSize: 22 }
  );
}

function addSectionSlides() {
  data.points.forEach(p => {
    const slide = pres.addSlide();

    slide.addShape(pres.shapes.RECTANGLE, {
      x: 0, y: 0, w: 10, h: 1.2,
      fill: { color: theme.navy }
    });

    slide.addText(p.title, {
      x: 0.5, y: 0.3, w: 9,
      fontSize: 32, bold: true, color: theme.white, align: "center"
    });

    slide.addText(p.verse, {
      x: 0.5, y: 1.6, w: 9,
      fontSize: 24, align: "center"
    });

    if (p.keyTruth) {
      const s = pres.addSlide();

      s.background = { color: "F8F9FA" };

      s.addText(p.keyTruth, {
        x: 0.5, y: 2, w: 9,
        fontSize: 44, bold: true, align: "center"
      });

      if (p.ref) {
        s.addText(`— ${p.ref}`, {
          x: 0.5, y: 4, w: 9,
          fontSize: 18, align: "center", italic: true
        });
      }
    }
  });
}

function addMaterialsSlide() {
  const slide = pres.addSlide();

  slide.addText("兩種材料", {
    x: 0.5, y: 0.5, w: 9,
    fontSize: 32, bold: true, align: "center"
  });

  const L = data.materials.left;
  const R = data.materials.right;

  slide.addText(L.title, { x: 1, y: 1.5, fontSize: 22, bold: true });
  slide.addText(L.items.join("\n"), { x: 1, y: 2.2 });

  slide.addText(R.title, { x: 5.5, y: 1.5, fontSize: 22, bold: true });
  slide.addText(R.items.join("\n"), { x: 5.5, y: 2.2 });
}

function addConclusionSlide() {
  const slide = pres.addSlide();
  slide.background = { color: theme.navy };

  slide.addText(data.title, {
    x: 0.5, y: 1.2, w: 9,
    fontSize: 40, bold: true, color: theme.white, align: "center"
  });

  slide.addText(
    data.conclusion.map(i => ({
      text: i,
      options: { bullet: true, breakLine: true }
    })),
    { x: 2, y: 2.8, w: 6, fontSize: 20, color: theme.lightBlue }
  );
}

function addHymnSlide() {
  const slide = pres.addSlide();

  slide.addText("回應詩歌", {
    x: 0.5, y: 1.5, w: 9,
    fontSize: 32, bold: true, align: "center"
  });

  slide.addText(data.hymn.title, {
    x: 0.5, y: 2.5, w: 9,
    fontSize: 48, bold: true, align: "center"
  });

  slide.addText(data.hymn.subtitle, {
    x: 0.5, y: 3.5, w: 9,
    fontSize: 20, align: "center", italic: true
  });
}

/* ---------- BUILD ---------- */

addTitleSlide();
addScriptureSlides();
addOutlineSlide();
addSectionSlides();
addMaterialsSlide();
addConclusionSlide();
addHymnSlide();

/* ---------- EXPORT ---------- */

pres.writeFile({ fileName: `sermon-${data.date}.pptx` });

console.log("✅ Presentation generated!");
