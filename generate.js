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
  headerHeight: 1.2, // height of the blue rectangle shape
  headerText: 0.625, // text position in the rectangle shape
  contentTitle: 1,   // content page — title position
  contentBullet: 3.125,  // content page — content position
  contentTop: 1.5,
  contentLeft: 1,
  contentWidth: 8
};

let pres = new PptxGenJS();

/* ---------- TEMPLATES ---------- */

function addTitleSlide() {
  const slide = pres.addSlide();
  slide.background = { color: theme.navy };

  slide.addText(data.title ?? "", {
    x: 0.5, y: layout.contentTop + 0.3, w: 9,
    fontSize: 44, bold: true, color: theme.white, align: "center"
  });

  if (data.subtitle) {
    slide.addText(data.subtitle, {
      x: 0.5, y: 3.2, w: 9,
      fontSize: 24, color: theme.lightBlue, align: "center"
    });
  }

  slide.addText(`${data.preacher ?? ""} | ${data.date ?? ""}`, {
    x: 0.5, y: 4.5, w: 9,
    fontSize: 16, color: theme.lightBlue, align: "center"
  });
}

function addScriptureSlides() {
  // LAYER 1: guard against missing scripture array
  (data.scripture ?? []).forEach(text => {
    const slide = pres.addSlide();

    slide.addText(data.subtitle ?? "", {
      x: 0.5, y: layout.contentTitle, w: 9,
      fontSize: 28, bold: true, align: "center"
    });

    slide.addText(text, {
      x: 0.8, y: layout.contentBullet, w: 8.4,
      fontSize: 18, lineSpacing: 28
    });
  });
}

function addIntroSlide() {
  if (!data.intro) return;
  const intro = data.intro;
  const slide = pres.addSlide();

  addHeader(slide, intro.title ?? data.labels?.intro ?? "Intro");

  // LAYER 1: guard against missing tensions array
  if (intro.tensions?.length) {
    slide.addText(
      intro.tensions.map(t => ({ text: t, options: { bullet: true, breakLine: true } })),
      {
        x: 1, y: layout.headerHeight + layout.contentTop, w: 8,
        fontSize: 20, lineSpacing: 34
      }
    );
  }

  // LAYER 1: only render footer if at least one part is present
  const footerParts = [intro.pivot, intro.question].filter(Boolean);
  if (footerParts.length) {
    slide.addText(footerParts.join("\n\n"), {
      x: 1, y: layout.headerHeight + layout.contentTop + 2, w: 8,
      fontSize: 18, italic: true, align: "center", color: theme.gray
    });
  }
}

function addOutlineSlide() {
  // LAYER 1: backward compat — accepts both old array format and new object format
  const outlineItems = Array.isArray(data.outline)
    ? data.outline
    : (data.outline?.items ?? []);

  if (!outlineItems.length) return;

  const slide = pres.addSlide();

  slide.addText(data.labels?.outline ?? "Outline", {
    x: 0.5, y: layout.contentTitle, w: 9,
    fontSize: 36, bold: true, align: "center"
  });

  slide.addText(
    outlineItems.map(item => ({
      text: item,
      options: { bullet: true, breakLine: true }
    })),
    { x: 1.5, y: layout.contentBullet, w: 7, fontSize: 22 }
  );
}

function addHeader(slide, title) {
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 10, h: layout.headerHeight,
    fill: { color: theme.navy }
  });

  slide.addText(title, {
    x: 0.5, y: layout.headerText, w: 9,
    fontSize: 32, bold: true, color: theme.white, align: "center"
  });
}

// LAYER 2: accepts optional tagline; adjusts verse position accordingly
function addSectionTitleSlide(title, verse, tagline) {
  const slide = pres.addSlide();

  addHeader(slide, title);

  // Without a tagline, drop the verse lower so it feels centered in the slide.
  // With a tagline, verse and tagline sit together in the upper content area.
  const verseY = tagline ? 2.2 : 3.0;

  if (verse) {
    slide.addText(verse, {
      x: 0.5, y: verseY, w: 9,
      fontSize: 24, align: "center", color: theme.gray
    });
  }

  if (tagline) {
    slide.addText(tagline, {
      x: 1, y: 3.2, w: 8,
      fontSize: 20, italic: true, align: "center", color: theme.gray
    });
  }
}

// title comes from p.subPointsTitle in JSON
function addSubPointsSlide(title, subPoints) {
  // LAYER 1: guard against empty array
  if (!subPoints?.length) return;

  const slide = pres.addSlide();

  addHeader(slide, title);

  slide.addText(
    subPoints.flatMap(sp => ([
      {
        text: `${sp.label ?? ""}（${sp.ref ?? ""}）`,
        options: { bold: true, breakLine: true }
      },
      {
        text: sp.truth ?? "",
        options: { fontSize: 18, color: theme.gray, breakLine: true }
      },
      { text: " ", options: { breakLine: true } }
    ])),
    {
      x: 1.5, y: layout.headerHeight + layout.contentTop + 1, w: 7,
      fontSize: 20, lineSpacing: 30
    }
  );
}

function addGreekWordsSlide(title, words) {
  // LAYER 1: guard against empty array
  if (!words?.length) return;

  const slide = pres.addSlide();

  slide.addText(title, {
    x: 0.5, y: layout.contentTitle, w: 9,
    fontSize: 32, bold: true, align: "center"
  });

  slide.addText(
    words.flatMap(w => ([
      {
        text: w.translit ? `${w.term} (${w.translit})` : (w.term ?? ""),
        options: { bold: true, breakLine: true }
      },
      {
        text: w.meaning ?? "",
        options: { fontSize: 18, breakLine: true }
      },
      { text: " ", options: { breakLine: true } }
    ])),
    {
      x: 2, y: layout.contentBullet, w: 6,
      fontSize: 20
    }
  );
}

function addTeachingSlide(title, points) {
  // LAYER 1: guard against missing points array
  if (!points?.length) return;

  const slide = pres.addSlide();

  slide.addText(title ?? "", {
    x: 0.5, y: layout.contentTitle, w: 9,
    fontSize: 32, bold: true, align: "center"
  });

  slide.addText(
    points.map(p => ({
      text: p,
      options: { breakLine: true }
    })),
    {
      x: 1.5, y: layout.contentBullet, w: 7,
      fontSize: 22, lineSpacing: 36
    }
  );
}

function addKeyTruthSlide(text, ref = null) {
  // LAYER 1: guard against missing text
  if (!text) return;

  const slide = pres.addSlide();

  slide.background = { color: "F8F9FA" };

  slide.addText(text, {
    x: 0.5, y: layout.contentTop + 0.3, w: 9,
    fontSize: 40, bold: true, align: "center"
  });

  if (ref) {
    slide.addText(`— ${ref}`, {
      x: 0.5, y: 4, w: 9,
      fontSize: 18, align: "center", italic: true
    });
  }
}

function addMaterialsSlide(materials) {
  // LAYER 1: guard against missing or incomplete materials
  if (!materials?.left || !materials?.right) return;

  const slide = pres.addSlide();

  slide.addText(materials.title ?? "", {
    x: 0.5, y: 0.5, w: 9,
    fontSize: 32, bold: true, align: "center"
  });

  const L = materials.left;
  const R = materials.right;

  slide.addText(L.title ?? "", {
    x: 0.8, y: layout.contentTop, w: 4,
    fontSize: 22, bold: true, align: "center"
  });

  slide.addText(
    (L.items ?? []).map(i => ({ text: i, options: { breakLine: true } })),
    {
      x: 1.2, y: layout.contentTop + 1.2, w: 3,   // user adjustment: 1.2
      fontSize: 18
    }
  );

  slide.addText(R.title ?? "", {
    x: 5.2, y: layout.contentTop, w: 4,
    fontSize: 22, bold: true, align: "center"
  });

  slide.addText(
    (R.items ?? []).map(i => ({ text: i, options: { breakLine: true } })),
    {
      x: 5.6, y: layout.contentTop + 1.2, w: 3,   // user adjustment: 1.2
      fontSize: 18
    }
  );
}

function addEmphasisScriptureSlide(title, lines) {
  // LAYER 1: guard against missing lines array
  if (!lines?.length) return;

  const slide = pres.addSlide();

  slide.addText(title ?? "", {
    x: 0.5, y: layout.headerHeight, w: 9,
    fontSize: 32, bold: true, align: "center"
  });

  slide.addText(
    lines.map(l => ({
      text: l.text ?? "",
      options: {
        breakLine: true,
        bold: l.bold ?? false,
        color: l.color ?? "000000"
      }
    })),
    {
      x: 1, y: layout.headerHeight + layout.contentTop + 0.5, w: 8,
      fontSize: 20, lineSpacing: 34, align: "center"
    }
  );
}

function addConclusionSlide() {
  const conc = data.conclusion;
  if (!conc) return;

  // LAYER 1: backward compat — accepts both old array format and new object format
  if (Array.isArray(conc)) {
    const slide = pres.addSlide();
    slide.background = { color: theme.navy };

    slide.addText(data.title ?? "", {
      x: 0.5, y: 1.2, w: 9,
      fontSize: 40, bold: true, color: theme.white, align: "center"
    });

    slide.addText(
      conc.map(i => ({ text: i, options: { bullet: true, breakLine: true } })),
      { x: 2, y: 2.8, w: 6, fontSize: 20, color: theme.lightBlue }
    );
    return;
  }

  // New object format: two-worlds comparison + invitation/closing

  // Slide 1: Two-worlds comparison
  if (conc.twoWorlds?.left && conc.twoWorlds?.right) {
    const slide1 = pres.addSlide();
    slide1.background = { color: theme.navy };

    const L = conc.twoWorlds.left;
    const R = conc.twoWorlds.right;

    slide1.addText(L.label ?? "", {
      x: 0.5, y: 0.6, w: 4.2,
      fontSize: 22, bold: true, color: theme.white, align: "center"
    });
    slide1.addText(
      (L.items ?? []).map(i => ({ text: i, options: { bullet: true, breakLine: true } })),
      { x: 0.5, y: 1.4, w: 4.2, fontSize: 18, color: theme.lightBlue, lineSpacing: 30 }
    );

    slide1.addText(R.label ?? "", {
      x: 5.2, y: 0.6, w: 4.2,
      fontSize: 22, bold: true, color: theme.white, align: "center"
    });
    slide1.addText(
      (R.items ?? []).map(i => ({ text: i, options: { bullet: true, breakLine: true } })),
      { x: 5.2, y: 1.4, w: 4.2, fontSize: 18, color: theme.lightBlue, lineSpacing: 30 }
    );
  }

  // Slide 2: Invitation + Closing
  if (conc.invitation || conc.closing) {
    const slide2 = pres.addSlide();
    slide2.background = { color: theme.navy };

    if (conc.invitation) {
      slide2.addText(conc.invitation, {
        x: 1, y: 1.5, w: 8,
        fontSize: 24, color: theme.lightBlue, align: "center", italic: true
      });
    }

    if (conc.closing) {
      slide2.addText(conc.closing, {
        x: 1, y: 3, w: 8,
        fontSize: 32, bold: true, color: theme.white, align: "center"
      });
    }
  }
}

function addHymnSlide() {
  // LAYER 1: guard against missing hymn block
  if (!data.hymn) return;

  const slide = pres.addSlide();

  slide.addText(data.labels?.hymn ?? "Hymn", {
    x: 0.5, y: 1.5, w: 9,
    fontSize: 32, bold: true, align: "center"
  });

  slide.addText(data.hymn.title ?? "", {
    x: 0.5, y: 2.5, w: 9,
    fontSize: 48, bold: true, align: "center"
  });

  if (data.hymn.subtitle) {
    slide.addText(data.hymn.subtitle, {
      x: 0.5, y: 3.5, w: 9,
      fontSize: 20, align: "center", italic: true
    });
  }
}

/* ---------- BUILD ---------- */

addTitleSlide();
addScriptureSlides();
addIntroSlide();      // intro before outline
addOutlineSlide();    // reads data.outline.items (with array fallback)

data.points.forEach(p => {
  addSectionTitleSlide(p.title, p.verse, p.tagline);  // LAYER 2: tagline

  if (p.subPoints) {
    addSubPointsSlide(p.subPointsTitle ?? data.labels?.subPoints ?? "Points", p.subPoints);
  }

  if (p.greekWords) {
    addGreekWordsSlide(data.labels?.greekWords ?? "Key Terms", p.greekWords);
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

pres.writeFile({ fileName: `sermon-${data.date ?? "output"}.pptx` });

console.log(`✅ Presentation generated from ${inputFile}!`);