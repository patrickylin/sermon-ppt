const PptxGenJS = require("pptxgenjs");
const fs = require("fs");

// Accept JSON filename as command-line argument
// Usage: node generate.js sermon-data-matt12_1-14.json
const inputFile = process.argv[2] || "sermon-data.json";
const data = JSON.parse(fs.readFileSync(inputFile, "utf-8"));

const theme = {
  navy: "1E2761",
  white: "FFFFFF",
  lightBlue: "CADCFC",
  gray: "363636"
};

const layout = {
  headerHeight: 1.2,
  contentTop: 1.5,
  contentLeft: 1,
  contentWidth: 8
};

let pres = new PptxGenJS();

/* ---------- TEMPLATES ---------- */

function addTitleSlide() {
  const slide = pres.addSlide();
  slide.background = { color: theme.navy };

  slide.addText(data.title, {
    x: 0.5, y: layout.contentTop + 0.3, w: 9,
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
      x: 0.5, y: layout.headerHeight, w: 9,
      fontSize: 28, bold: true, align: "center"
    });

    slide.addText(text, {
      x: 0.8, y: layout.contentLeft + layout.contentTop, w: 8.4,
      fontSize: 18, lineSpacing: 28
    });
  });
}

// NEW: renders data.intro before the outline slide
function addIntroSlide() {
  if (!data.intro) return;
  const intro = data.intro;
  const slide = pres.addSlide();

  addHeader(slide, intro.title);

  slide.addText(
    intro.tensions.map(t => ({ text: t, options: { bullet: true, breakLine: true } })),
    {
      x: 1, y: layout.contentTop + 0.5, w: 8,
      fontSize: 20, lineSpacing: 34
    }
  );

  slide.addText(`${intro.pivot}\n\n${intro.question}`, {
    x: 1, y: 3.6, w: 8,
    fontSize: 18, italic: true, align: "center", color: theme.gray
  });
}

// CHANGED: now reads data.outline.items instead of data.outline
function addOutlineSlide() {
  const slide = pres.addSlide();

  slide.addText("信息大綱", {
    x: 0.5, y: layout.headerHeight, w: 9,
    fontSize: 36, bold: true, align: "center"
  });

  slide.addText(
    data.outline.items.map(item => ({
      text: item,
      options: { bullet: true, breakLine: true }
    })),
    { x: 1.5, y: layout.contentLeft + layout.contentTop, w: 7, fontSize: 22 }
  );
}

function addHeader(slide, title) {
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 10, h: layout.headerHeight,
    fill: { color: theme.navy }
  });

  slide.addText(title, {
    x: 0.5, y: 0.8, w: 9,
    fontSize: 32, bold: true, color: theme.white, align: "center"
  });
}

function addSectionTitleSlide(title, verse) {
  const slide = pres.addSlide();

  addHeader(slide, title);

  slide.addText(verse, {
    x: 0.5, y: layout.contentTop, w: 9,
    fontSize: 24, align: "center"
  });
}

// NEW: renders p.subPoints as a single slide
// title comes from p.subPointsTitle in JSON
function addSubPointsSlide(title, subPoints) {
  const slide = pres.addSlide();

  addHeader(slide, title);

  slide.addText(
    subPoints.flatMap(sp => ([
      {
        text: `${sp.label}（${sp.ref}）`,
        options: { bold: true, breakLine: true }
      },
      {
        text: sp.truth,
        options: { fontSize: 18, color: theme.gray, breakLine: true }
      },
      { text: " ", options: { breakLine: true } }
    ])),
    {
      x: 1.5, y: layout.headerHeight + layout.contentTop + 0.5, w: 7,
      fontSize: 20, lineSpacing: 30
    }
  );
}

function addGreekWordsSlide(title, words) {
  const slide = pres.addSlide();

  slide.addText(title, {
    x: 0.5, y: layout.headerHeight, w: 9,
    fontSize: 32, bold: true, align: "center"
  });

  slide.addText(
    words.flatMap(w => ([
      {
        text: w.translit
          ? `${w.term} (${w.translit})`
          : w.term,
        options: { bold: true, breakLine: true }
      },
      {
        text: w.meaning,
        options: { fontSize: 18, breakLine: true }
      },
      { text: " ", options: { breakLine: true } }
    ])),
    {
      x: 2, y: layout.headerHeight + layout.contentTop + 0.3, w: 6,
      fontSize: 20
    }
  );
}

function addTeachingSlide(title, points) {
  const slide = pres.addSlide();

  slide.addText(title, {
    x: 0.5, y: layout.headerHeight, w: 9,
    fontSize: 32, bold: true, align: "center"
  });

  slide.addText(
    points.map(p => ({
      text: p,
      options: { breakLine: true }
    })),
    {
      x: 1.5, y: layout.headerHeight + layout.contentTop + 0.3, w: 7,
      fontSize: 22, lineSpacing: 36
    }
  );
}

function addKeyTruthSlide(text, ref = null) {
  const slide = pres.addSlide();

  slide.background = { color: "F8F9FA" };

  slide.addText(text, {
    x: 0.5, y: layout.contentTop + 0.3, w: 9,
    fontSize: 44, bold: true, align: "center"
  });

  if (ref) {
    slide.addText(`— ${ref}`, {
      x: 0.5, y: 4, w: 9,
      fontSize: 18, align: "center", italic: true
    });
  }
}

function addMaterialsSlide(materials) {
  const slide = pres.addSlide();

  slide.addText(materials.title, {
    x: 0.5, y: 0.5, w: 9,
    fontSize: 32, bold: true, align: "center"
  });

  const L = materials.left;
  const R = materials.right;

  slide.addText(L.title, {
    x: 0.8, y: layout.contentTop, w: 4,
    fontSize: 22, bold: true, align: "center"
  });

  slide.addText(
    L.items.map(i => ({ text: i, options: { breakLine: true } })),
    {
      x: 1.2, y: layout.contentTop + 1.2, w: 3,
      fontSize: 18
    }
  );

  slide.addText(R.title, {
    x: 5.2, y: layout.contentTop, w: 4,
    fontSize: 22, bold: true, align: "center"
  });

  slide.addText(
    R.items.map(i => ({ text: i, options: { breakLine: true } })),
    {
      x: 5.6, y: layout.contentTop + 1.2, w: 3,
      fontSize: 18
    }
  );
}

function addEmphasisScriptureSlide(title, lines) {
  const slide = pres.addSlide();

  slide.addText(title, {
    x: 0.5, y: layout.headerHeight, w: 9,
    fontSize: 32, bold: true, align: "center"
  });

  slide.addText(
    lines.map(l => ({
      text: l.text,
      options: {
        breakLine: true,
        bold: l.bold || false,
        color: l.color || "000000"
      }
    })),
    {
      x: 1, y: layout.headerHeight + layout.contentTop + 0.5, w: 8,
      fontSize: 20, lineSpacing: 34, align: "center"
    }
  );
}

// CHANGED: handles new conclusion object with twoWorlds + invitation + closing
// produces two slides instead of one
function addConclusionSlide() {
  const conc = data.conclusion;

  // Slide 1: Two-worlds comparison
  const slide1 = pres.addSlide();
  slide1.background = { color: theme.navy };

  const L = conc.twoWorlds.left;
  const R = conc.twoWorlds.right;

  slide1.addText(L.label, {
    x: 0.5, y: 0.6, w: 4.2,
    fontSize: 22, bold: true, color: theme.white, align: "center"
  });
  slide1.addText(
    L.items.map(i => ({ text: i, options: { bullet: true, breakLine: true } })),
    {
      x: 0.5, y: 1.4, w: 4.2,
      fontSize: 18, color: theme.lightBlue, lineSpacing: 30
    }
  );

  slide1.addText(R.label, {
    x: 5.2, y: 0.6, w: 4.2,
    fontSize: 22, bold: true, color: theme.white, align: "center"
  });
  slide1.addText(
    R.items.map(i => ({ text: i, options: { bullet: true, breakLine: true } })),
    {
      x: 5.2, y: 1.4, w: 4.2,
      fontSize: 18, color: theme.lightBlue, lineSpacing: 30
    }
  );

  // Slide 2: Invitation + Closing
  const slide2 = pres.addSlide();
  slide2.background = { color: theme.navy };

  slide2.addText(conc.invitation, {
    x: 1, y: 1.5, w: 8,
    fontSize: 24, color: theme.lightBlue, align: "center", italic: true
  });

  slide2.addText(conc.closing, {
    x: 1, y: 3, w: 8,
    fontSize: 32, bold: true, color: theme.white, align: "center"
  });
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
addIntroSlide();      // NEW: intro before outline
addOutlineSlide();    // now reads data.outline.items

data.points.forEach(p => {
  addSectionTitleSlide(p.title, p.verse);

  if (p.subPoints) {                                      // NEW
    addSubPointsSlide(p.subPointsTitle || "要點", p.subPoints);
  }

  if (p.greekWords) {
    addGreekWordsSlide("關鍵字詞", p.greekWords);
  }

  if (p.keyTruth) {
    addKeyTruthSlide(p.keyTruth, p.ref);
  }

  if (p.teaching) {
    addTeachingSlide(p.teaching.title, p.teaching.points);
  }

  if (p.materials) {
    addMaterialsSlide(p.materials);
  }

  if (p.emphasis) {
    addEmphasisScriptureSlide(p.emphasis.title, p.emphasis.lines);
  }
});

addConclusionSlide();
addHymnSlide();

/* ---------- EXPORT ---------- */

pres.writeFile({ fileName: `sermon-${data.date}.pptx` });

console.log(`✅ Presentation generated from ${inputFile}!`);