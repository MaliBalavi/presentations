// Pesmokaz — generator for the PFE presentation deck.
// Run: node deck.js   (then rename/validate with the pptx skill)
const path = require("path");
const pptxgen = require("pptxgenjs");
const { applyTheme } = require(
  "/home/andrej/.config/opencode/skills/pptx/scripts/apply_theme.js",
);

const A = path.join(__dirname, "assets");
const LOGO = path.join(A, "pfe-logo.png");
const SPEC = path.join(A, "spectrogram.png");
const WAVE = path.join(A, "waveform.png");

const F = "Calibri";
const LANG = "sr-Latn-RS";

const C = {
  ink: "0E1118",
  ink2: "151B25",
  ink3: "222C3B",
  paper: "F6F7F9",
  white: "FFFFFF",
  teal: "009999",
  tealDark: "007373",
  tealLight: "5FD6D6",
  tealSoft: "DDF1F1",
  violet: "7C5CFC",
  violetSoft: "EAE4FF",
  amber: "FFB020",
  amberDark: "A86A00",
  amberSoft: "FFF1D6",
  text: "1A2230",
  muted: "5B6675",
  mutedOnDark: "9AA5B5",
  line: "E2E7EE",
};

const THEME = {
  name: "Pesmokaz",
  headFontFace: F,
  bodyFontFace: F,
  colors: {
    dk1: C.ink,
    lt1: C.paper,
    dk2: C.text,
    lt2: C.line,
    accent1: C.teal,
    accent2: C.violet,
    accent3: C.amber,
    accent4: C.muted,
    accent5: C.tealDark,
    accent6: C.tealSoft,
    hlink: C.teal,
    folHlink: C.tealDark,
  },
};

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";
pres.title = "Pesmokaz — Razvijanje i poređenje metoda prepoznavanja pesama";
pres.subject = "PFE u Galeriji nauke i tehnike SANU";
pres.author = "Lazar Aleksić, Andrej Vasiljević";
pres.company = "Naučno-inženjerski centar PFE";
pres.theme = { headFontFace: F, bodyFontFace: F };

// ---------------------------------------------------------------- helpers
const S = pres.ShapeType;

let currentSection = "Uvod";

function setSection(title) {
  currentSection = title;
  pres.addSection({ title });
}

function addText(slide, value, opts) {
  slide.addText(value, {
    fontFace: F,
    lang: LANG,
    margin: 0,
    valign: "top",
    ...opts,
  });
}

function addCard(slide, x, y, w, h, opts = {}) {
  slide.addShape(S.roundRect, {
    x,
    y,
    w,
    h,
    rectRadius: 0.06,
    fill: { color: opts.fill || C.white },
    line: { color: opts.line || C.line, width: opts.lineWidth ?? 0.75 },
    objectName: opts.name || "card",
  });
}

function addArrow(slide, x, y, w, h, color = C.muted) {
  slide.addShape(S.line, {
    x,
    y,
    w,
    h,
    line: { color, width: 1.25, endArrowType: "triangle" },
    objectName: "arrow",
  });
}

function addPill(slide, x, y, w, h, label, opts = {}) {
  slide.addShape(S.roundRect, {
    x,
    y,
    w,
    h,
    rectRadius: h / 2,
    fill: { color: opts.fill || C.white },
    line: { color: opts.line || C.line, width: 0.75 },
  });
  addText(slide, label, {
    x,
    y,
    w,
    h,
    align: "center",
    valign: "mid",
    fontSize: opts.fontSize || 13,
    bold: true,
    color: opts.color || C.text,
  });
}

function addStepNumber(slide, x, y, d, label, color = C.teal) {
  slide.addShape(S.ellipse, {
    x,
    y,
    w: d,
    h: d,
    fill: { color: C.white },
    line: { color, width: 1.5 },
  });
  addText(slide, label, {
    x,
    y,
    w: d,
    h: d,
    align: "center",
    valign: "mid",
    fontSize: d > 0.4 ? 14 : 12,
    bold: true,
    color,
  });
}

function addKicker(slide, label, dark) {
  addText(slide, label.toUpperCase(), {
    x: 0.8,
    y: 0.46,
    w: 9,
    h: 0.28,
    fontSize: 11,
    bold: true,
    charSpacing: 2,
    color: dark ? C.tealLight : C.teal,
  });
}

function addTitle(slide, label, dark, size = 34) {
  addText(slide, label, {
    x: 0.8,
    y: 0.74,
    w: 11.7,
    h: 0.95,
    fontSize: size,
    bold: true,
    color: dark ? C.white : C.text,
    valign: "top",
  });
}

function contentSlide(kicker, title, opts = {}) {
  const dark = !!opts.dark;
  const s = pres.addSlide({ masterName: dark ? "dark" : "light", sectionTitle: currentSection });
  activeSlide = s;
  addKicker(s, kicker, dark);
  addTitle(s, title, dark, opts.size);
  return s;
}

let activeSlide = null;

function addNotes(notes) {
  if (!activeSlide) throw new Error("addNotes called before a slide was created");
  activeSlide.addNotes(notes);
}

// ---------------------------------------------------------------- masters
pres.defineSlideMaster({
  title: "light",
  background: { color: C.paper },
  slideNumber: {
    x: 12.35,
    y: 6.98,
    w: 0.5,
    h: 0.26,
    color: C.muted,
    fontSize: 9,
    fontFace: F,
    align: "r",
  },
  objects: [
    { image: { path: LOGO, x: 0.62, y: 6.95, w: 0.27, h: 0.27 } },
    {
      text: {
        text: "Pesmokaz  ·  PFE u Galeriji nauke i tehnike SANU",
        options: {
          x: 0.98,
          y: 6.98,
          w: 8.4,
          h: 0.26,
          fontSize: 9,
          color: C.muted,
          fontFace: F,
          lang: LANG,
          margin: 0,
          valign: "mid",
        },
      },
    },
  ],
});

pres.defineSlideMaster({
  title: "dark",
  background: { color: C.ink },
  slideNumber: {
    x: 12.35,
    y: 6.98,
    w: 0.5,
    h: 0.26,
    color: C.mutedOnDark,
    fontSize: 9,
    fontFace: F,
    align: "r",
  },
  objects: [
    { image: { path: LOGO, x: 0.62, y: 6.95, w: 0.27, h: 0.27 } },
    {
      text: {
        text: "Pesmokaz  ·  PFE u Galeriji nauke i tehnike SANU",
        options: {
          x: 0.98,
          y: 6.98,
          w: 8.4,
          h: 0.26,
          fontSize: 9,
          color: C.mutedOnDark,
          fontFace: F,
          lang: LANG,
          margin: 0,
          valign: "mid",
        },
      },
    },
  ],
});

pres.defineSlideMaster({
  title: "cover",
  background: { color: C.ink },
  objects: [
    { image: { path: SPEC, x: 6.4, y: 4.75, w: 6.6, h: 2.15, altText: "Stilizovani mel spektrogram" } },
  ],
});

pres.addSection({ title: "Uvod" });
// ---------------------------------------------------------------- 1 cover
{
  const s = pres.addSlide({ masterName: "cover", sectionTitle: "Uvod" });
  activeSlide = s;

  s.addImage({ path: LOGO, x: 0.8, y: 0.62, w: 0.55, h: 0.55, altText: "PFE logo" });
  addText(s, "Naučno-inženjerski centar PFE", {
    x: 1.5,
    y: 0.72,
    w: 5.5,
    h: 0.3,
    fontSize: 11,
    color: C.mutedOnDark,
  });
  addText(s, "PFE u Galeriji nauke i tehnike SANU  ·  24–25. novembar 2026.", {
    x: 6.2,
    y: 0.72,
    w: 6.3,
    h: 0.3,
    fontSize: 11,
    color: C.mutedOnDark,
    align: "r",
  });

  addText(s, "Pesmokaz", {
    x: 0.8,
    y: 1.95,
    w: 8.4,
    h: 1.3,
    fontSize: 72,
    bold: true,
    color: C.white,
  });
  addText(s, "Razvijanje i poređenje metoda prepoznavanja pesama", {
    x: 0.82,
    y: 3.2,
    w: 8.2,
    h: 0.8,
    fontSize: 22,
    color: C.tealLight,
  });

  addText(s, "Lazar Aleksić  ·  Andrej Vasiljević", {
    x: 0.82,
    y: 4.55,
    w: 8,
    h: 0.4,
    fontSize: 17,
    bold: true,
    color: C.white,
  });
  addText(s, "Mentori: Andrej Bantulić  ·  Dimitrije Pešić", {
    x: 0.82,
    y: 4.98,
    w: 8,
    h: 0.35,
    fontSize: 13,
    color: C.mutedOnDark,
  });
  addText(s, "Gimnazija „Bora Stanković“, Vranje  ·  Prva kragujevačka gimnazija", {
    x: 0.82,
    y: 5.35,
    w: 9,
    h: 0.35,
    fontSize: 11,
    color: C.mutedOnDark,
  });

  addNotes(
    "Dobar dan. Ja sam Andrej Vasiljević, a ovo je Pesmokaz — projekat koji smo Lazar i ja radili uz mentore sa ETF-a. " +
      "Za deset minuta pokazaćemo kako smo napravili jedan sistem koji prepoznaje pesmu iz kratkog snimka i, važnije, kako smo pošteno poredili dve potpuno različite metode na istom skupu podataka. (20 s)",
  );
}

// ---------------------------------------------------------------- 2 problem
{
  const s = contentSlide("Problem", "Pet sekundi snimka, cela baza pesama");

  addText(s, "5 s", {
    x: 0.8,
    y: 1.9,
    w: 4.8,
    h: 1.45,
    fontSize: 96,
    bold: true,
    color: C.teal,
  });
  addText(s, "tipičan upit sa telefona", {
    x: 0.85,
    y: 3.28,
    w: 4.8,
    h: 0.35,
    fontSize: 14,
    color: C.muted,
  });
  s.addImage({ path: WAVE, x: 0.8, y: 3.95, w: 5.1, h: 0.91, altText: "Talasan oblik audio signala" });
  addText(s, "Snimak retko stiže čist — i to je ceo problem.", {
    x: 0.82,
    y: 5.15,
    w: 5.2,
    h: 0.6,
    fontSize: 14,
    italic: true,
    color: C.muted,
  });

  const items = [
    ["01", "Šum, kompresija, reverb", "Signal sa telefona menja se pre nego što stigne do modela."],
    ["02", "Ista pesma, druga verzija", "Kover ili snimak uživo menja frekvencijsku sliku."],
    ["03", "Baza koja raste", "Pretraga mora da ostane tačna i brza i sa milionima zapisa."],
  ];
  items.forEach(([num, head, body], i) => {
    const y = 1.9 + i * 1.32;
    addCard(s, 6.3, y, 6.2, 1.12);
    addText(s, num, { x: 6.6, y: y + 0.16, w: 0.5, h: 0.3, fontSize: 11, bold: true, color: C.teal });
    addText(s, head, { x: 7.15, y: y + 0.14, w: 5.1, h: 0.32, fontSize: 16, bold: true, color: C.text });
    addText(s, body, { x: 7.15, y: y + 0.5, w: 5.1, h: 0.5, fontSize: 14, color: C.muted });
  });

  addText(s, "Cilj: pronaći tačnu pesmu — i izmeriti koliko je to teško.", {
    x: 0.8,
    y: 6.12,
    w: 11.6,
    h: 0.4,
    fontSize: 17,
    bold: true,
    color: C.text,
  });

  addNotes(
    "Krenimo od svakodnevnog primera: snimite pet sekundi pesme u kafiću i očekujete da telefon pogodi tačnu pesmu. " +
      "Tri stvari to otežavaju: signal je prljav, ista pesma postoji u više verzija, a baza je ogromna. " +
      "Naš cilj nije bio da napravimo još jedan Šazam, nego da izmerimo dve ideje pod istim uslovima. (45 s)",
  );
}

// ---------------------------------------------------------------- 3 two families
{
  const s = contentSlide("Dve porodice metoda", "Heširanje i duboko učenje rešavaju isti problem drugačije");

  // left card
  addCard(s, 0.8, 1.95, 5.7, 3.55, { fill: C.white });
  addText(s, "Determinističko heširanje", { x: 1.1, y: 2.2, w: 5.1, h: 0.4, fontSize: 20, bold: true, color: C.tealDark });
  addText(s, "Wang, 2003", { x: 1.1, y: 2.6, w: 5.1, h: 0.3, fontSize: 11, color: C.muted });
  addText(
    s,
    [
      { text: "Pikovi spektrograma → parovi → heševi", options: { bullet: true, breakLine: true } },
      { text: "Otporno na šum: najjači pikovi prežive", options: { bullet: true, breakLine: true } },
      { text: "Brzo, ali ne razume strukturu pesme", options: { bullet: true } },
    ],
    { x: 1.1, y: 3.02, w: 5.1, h: 1.5, fontSize: 14, color: C.text, paraSpaceAfter: 8 },
  );
  // mini hash diagram
  const dots = [
    [3.0, 4.75],
    [3.55, 4.5],
    [3.95, 4.95],
    [4.55, 4.62],
    [5.05, 4.88],
  ];
  dots.forEach(([x, y]) => {
    s.addShape(S.ellipse, { x, y, w: 0.09, h: 0.09, fill: { color: C.teal }, line: { color: C.teal, width: 0 } });
  });
  [[0, 1], [1, 2], [2, 3], [3, 4]].forEach(([a, b]) => {
    s.addShape(S.line, {
      x: dots[a][0] + 0.045,
      y: dots[a][1] + 0.045,
      w: dots[b][0] - dots[a][0],
      h: dots[b][1] - dots[a][1],
      line: { color: C.teal, width: 0.75, transparency: 35 },
    });
  });
  addText(s, "pikovi + vremenski pomeraj", { x: 3.0, y: 5.1, w: 3, h: 0.25, fontSize: 10, color: C.muted });

  // right card
  addCard(s, 6.85, 1.95, 5.7, 3.55, { fill: C.white });
  addText(s, "Duboko učenje", { x: 7.15, y: 2.2, w: 5.1, h: 0.4, fontSize: 20, bold: true, color: C.violet });
  addText(s, "Schroff, 2015  ·  Chang, 2021", { x: 7.15, y: 2.6, w: 5.1, h: 0.3, fontSize: 11, color: C.muted });
  addText(
    s,
    [
      { text: "Triplet loss uči rastojanje u latentnom prostoru", options: { bullet: true, breakLine: true } },
      { text: "Može da prepozna kover i izvođenje uživo", options: { bullet: true, breakLine: true } },
      { text: "Zavisi od podataka i kvaliteta treniranja", options: { bullet: true } },
    ],
    { x: 7.15, y: 3.02, w: 5.1, h: 1.5, fontSize: 14, color: C.text, paraSpaceAfter: 8 },
  );
  // mini embedding vector
  for (let i = 0; i < 10; i++) {
    const hot = i === 3 || i === 6 || i === 8;
    s.addShape(S.roundRect, {
      x: 7.2 + i * 0.34,
      y: 4.62,
      w: 0.24,
      h: 0.42,
      rectRadius: 0.03,
      fill: { color: hot ? C.amber : C.violet, transparency: hot ? 0 : 55 },
      line: { color: C.violet, width: 0 },
    });
  }
  addText(s, "embedding: ista pesma blizu, različite daleko", { x: 7.15, y: 5.1, w: 5, h: 0.25, fontSize: 10, color: C.muted });

  addText(s, "Retko se porede na istom skupu i istom protokolu.", {
    x: 0.8,
    y: 5.95,
    w: 11.6,
    h: 0.5,
    fontSize: 19,
    bold: true,
    color: C.text,
  });

  addNotes(
    "Dve porodice metoda. Sa leve strane je klasično heširanje: iz spektrograma se uzimaju najjači pikovi i prave parovi. " +
      "Sa desne strane je duboko učenje: mreža uči da istu pesmu smesti blizu, a različite daleko. " +
      "Prva je otporna na šum, druga obećava koveri i snimke uživo. Ono što retko vidimo je pošteno poređenje — to je praznina koju smo popunili. (55 s)",
  );
}

// ---------------------------------------------------------------- 4 question
{
  const s = contentSlide("Istraživačko pitanje", "Isto pitanje, isti protokol", { dark: true });

  addText(s, "Koja metoda je tačnija i brža\nna istom skupu podataka?", {
    x: 0.8,
    y: 1.95,
    w: 10.6,
    h: 1.5,
    fontSize: 34,
    bold: true,
    color: C.white,
  });
  addText(s, "Hipoteza: matematika je tačnija na šumu; CNN je brži i možda bolji na strukturnim mutacijama.", {
    x: 0.82,
    y: 3.5,
    w: 10.8,
    h: 0.5,
    fontSize: 15,
    color: C.mutedOnDark,
  });

  const contrib = [
    ["01", "Kompletan lanac obrade", "od audio ulaza do rang liste, sa testovima"],
    ["02", "Zajednički interfejs", "dve metode iza istog Python API-ja"],
    ["03", "Protokol evaluacije", "776 pesama, pet nivoa šuma, iste metrike"],
  ];
  contrib.forEach(([n, head, body], i) => {
    const x = 0.8 + i * 4.0;
    addText(s, n, { x, y: 4.55, w: 1.2, h: 0.5, fontSize: 26, bold: true, color: C.teal });
    addText(s, head, { x, y: 5.1, w: 3.5, h: 0.35, fontSize: 15, bold: true, color: C.white });
    addText(s, body, { x, y: 5.45, w: 3.5, h: 0.6, fontSize: 13, color: C.mutedOnDark });
  });

  addNotes(
    "Naše pitanje je jednostavno: ako obe metode pustimo na isti skup i isti protokol, koja je tačnija i koja je brža? " +
      "Hipoteza je bila da je matematika tačnija pod šumom, a da je CNN brži i da bi mogao biti bolji na mutacijama. " +
      "Da bismo to izmerili, napravili smo tri stvari: ceo lanac, zajednički interfejs i protokol koji se ne menja između metoda. (40 s)",
  );
}

setSection("Metoda");

// ---------------------------------------------------------------- 5 pipeline
{
  const s = contentSlide("Sistem", "Lanac obrade: prvo isečak, pa normalizacija");

  const nodes = [
    ["Ulaz", "22050 Hz, mono"],
    ["Isečci", "5 s, preklapanje 2,5 s"],
    ["Normalizacija", "−23 LUFS, po isečku"],
    ["Mel spektrogram", "128 filtera · opciono U-Net"],
  ];
  const wNode = 2.55;
  const gap = 0.45;
  nodes.forEach(([head, sub], i) => {
    const x = 0.8 + i * (wNode + gap);
    addCard(s, x, 1.95, wNode, 1.05);
    addText(s, head, { x: x + 0.18, y: 2.06, w: wNode - 0.36, h: 0.32, fontSize: 14, bold: true, color: C.text });
    addText(s, sub, { x: x + 0.18, y: 2.4, w: wNode - 0.36, h: 0.5, fontSize: 11, color: C.muted });
    if (i < nodes.length - 1) addArrow(s, x + wNode + 0.06, 2.47, gap - 0.12, 0);
  });

  addText(s, "iz istog mel spektrograma — dve grane", {
    x: 0.8,
    y: 3.12,
    w: 11.6,
    h: 0.28,
    fontSize: 11,
    color: C.muted,
    align: "center",
  });

  addCard(s, 0.8, 3.45, 5.6, 1.0, { fill: C.tealSoft, line: C.tealSoft });
  addText(s, "Matematička grana", { x: 1.05, y: 3.57, w: 5.1, h: 0.3, fontSize: 15, bold: true, color: C.tealDark });
  addText(s, "pikovi → heševi → vremensko poravnanje", { x: 1.05, y: 3.9, w: 5.1, h: 0.35, fontSize: 13, color: C.text });

  addCard(s, 6.95, 3.45, 5.6, 1.0, { fill: C.violetSoft, line: C.violetSoft });
  addText(s, "CNN grana", { x: 7.2, y: 3.57, w: 5.1, h: 0.3, fontSize: 15, bold: true, color: C.violet });
  addText(s, "mel → ResNet-18 → 128-d embedding", { x: 7.2, y: 3.9, w: 5.1, h: 0.35, fontSize: 13, color: C.text });

  addArrow(s, 3.6, 4.5, 0, 0.3, C.tealDark);
  addArrow(s, 9.75, 4.5, 0, 0.3, C.violet);

  addCard(s, 0.8, 4.85, 11.75, 0.9, { fill: C.white });
  addText(s, "Pretraga u bazi  →  rang lista pesama", {
    x: 1.05,
    y: 4.97,
    w: 6.0,
    h: 0.3,
    fontSize: 15,
    bold: true,
    color: C.text,
  });
  addText(s, "heš baza (dict, pickle)  ·  vektorska baza (FAISS, L2)", {
    x: 1.05,
    y: 5.32,
    w: 6.2,
    h: 0.28,
    fontSize: 11,
    color: C.muted,
  });
  addText(s, "najbliža pesma + ocena", {
    x: 8.6,
    y: 4.97,
    w: 3.7,
    h: 0.6,
    fontSize: 13,
    bold: true,
    color: C.teal,
    align: "r",
    valign: "mid",
  });

  addCard(s, 0.8, 6.02, 11.75, 0.72, { fill: C.amberSoft, line: C.amberSoft });
  addText(s, "Prvo isečak, pa normalizacija — inače isti segment dobija različitu glasnoću kod svakog snimka.", {
    x: 1.05,
    y: 6.12,
    w: 11.2,
    h: 0.5,
    fontSize: 14,
    bold: true,
    color: C.amberDark,
    valign: "mid",
  });

  addNotes(
    "Ovo je ceo sistem u jednoj slici. Audio se deli na isečke od pet sekundi sa preklapanjem od 2,5, svaki isečak se posebno normalizuje na −23 LUFS, " +
      "pa se računa mel spektrogram sa 128 filtera. Iz njega idu dve grane: heševi i embedding. " +
      "Jedan detalj je važniji nego što izgleda: prvo isečak, pa normalizacija. Da radimo obrnuto, isti segment bi dobio različitu glasnoću kod svakog snimka, i poređenje bi pucalo. (75 s)",
  );
}

// ---------------------------------------------------------------- 6 dataset
{
  const s = contentSlide("Podaci", "776 pesama, 16 žanrova, pet nivoa šuma");

  const stats = [
    ["776", "upotrebljivih pesama"],
    ["16", "žanrova, ravnomerno"],
    ["~71.000", "isečaka od 5 s"],
    ["7.281", "zašumljenih upita"],
  ];
  stats.forEach(([num, label], i) => {
    const x = 0.8 + i * 2.98;
    addCard(s, x, 1.95, 2.7, 1.45);
    addText(s, num, { x: x + 0.2, y: 2.08, w: 2.3, h: 0.75, fontSize: 40, bold: true, color: C.teal });
    addText(s, label, { x: x + 0.2, y: 2.88, w: 2.35, h: 0.35, fontSize: 12, color: C.muted });
  });

  addText(s, "Nivoi šuma u evaluaciji", { x: 0.8, y: 3.68, w: 3, h: 0.28, fontSize: 11, color: C.muted });
  ["−5 dB", "0 dB", "5 dB", "10 dB", "20 dB"].forEach((label, i) => {
    addPill(s, 0.8 + i * 1.62, 4.0, 1.45, 0.46, label, { fontSize: 13 });
  });
  addText(s, "isti niz slučajnih brojeva za obe metode", {
    x: 9.0,
    y: 4.08,
    w: 3.5,
    h: 0.3,
    fontSize: 11,
    color: C.muted,
    align: "r",
  });

  addText(
    s,
    [
      { text: "1000 → 776: 224 korumpirane pesme uklonjene MD5 proverom", options: { bullet: true, breakLine: true } },
      { text: "Podela po pesmama, ne po isečcima — nema curenja između treninga i testa", options: { bullet: true, breakLine: true } },
      { text: "χ² = 24,21 (df = 15): raspodela po žanrovima je ravnomerna", options: { bullet: true } },
    ],
    { x: 0.8, y: 4.82, w: 11.4, h: 1.5, fontSize: 15, color: C.text, paraSpaceAfter: 9 },
  );

  addText(s, "FMA, 5,5 GB, CC licence  ·  šum: pink + govor u pozadini (kada postoji)  ·  bez reverba", {
    x: 0.8,
    y: 6.35,
    w: 11.6,
    h: 0.3,
    fontSize: 10,
    color: C.muted,
  });

  addNotes(
    "Skup je 776 pesama iz FMA kolekcije, stratifikovano po 16 žanrova. Krenuli smo od 1000, ali 224 fajla su bila korumpirana — to smo uhvatili MD5 proverom. " +
      "Podela je po pesmama, ne po isečcima, da model ne bi video deo iste pesme u treningu i testu. " +
      "Evaluacija ide na pet nivoa šuma, od −5 do 20 dB, sa istim slučajnim nizom za obe metode, ukupno 7.281 upit. (55 s)",
  );
}

// ---------------------------------------------------------------- 7 hashing
{
  const s = contentSlide("Matematička metoda", "Heširanje: pikovi, parovi, poravnanje");

  const steps = [
    ["1", "Pikovi", "Maximum filter 16×16 na mel spektrogramu; zadržavamo pikove iznad praga 0,25."],
    ["2", "Heševi", "Za svaki pik 15 parova sa susedima: heš je (f₁, f₂, Δt), uz ograničen pomeraj."],
    ["3", "Poravnanje", "Ocena pesme = najviše poklapanja na jednom vremenskom pomaku."],
  ];
  steps.forEach(([n, head, body], i) => {
    const y = 1.95 + i * 1.18;
    addStepNumber(s, 0.82, y + 0.05, 0.44, n);
    addText(s, head, { x: 1.42, y, w: 2.2, h: 0.32, fontSize: 16, bold: true, color: C.text });
    addText(s, body, { x: 3.35, y, w: 3.95, h: 0.95, fontSize: 13, color: C.muted });
  });

  // evolution panel
  addCard(s, 7.55, 1.95, 5.0, 3.5, { fill: C.white });
  addText(s, "Top-1 tačnost kroz iteracije", { x: 7.85, y: 2.12, w: 4.4, h: 0.3, fontSize: 12, color: C.muted });
  const bars = [
    ["Originalni parametri", 25, C.muted],
    ["Tuning heševa i praga", 70, C.teal],
    ["Vremensko poravnanje", 97.56, C.amberDark],
  ];
  bars.forEach(([label, value, color], i) => {
    const y = 2.62 + i * 0.88;
    addText(s, label, { x: 7.85, y: y - 0.26, w: 4.3, h: 0.26, fontSize: 10, color: C.muted });
    const maxW = 3.5;
    s.addShape(S.roundRect, {
      x: 7.85,
      y,
      w: (value / 100) * maxW,
      h: 0.26,
      rectRadius: 0.04,
      fill: { color },
      line: { color, width: 0 },
    });
    addText(s, `${String(value).replace(".", ",")} %`, {
      x: 7.85 + (value / 100) * maxW + 0.12,
      y: y - 0.03,
      w: 1.3,
      h: 0.3,
      fontSize: 14,
      bold: true,
      color: C.text,
    });
  });

  addText(s, "Poravnanje je najveći pojedinačni skok: 70 % → 97,56 %.", {
    x: 0.8,
    y: 5.9,
    w: 11.6,
    h: 0.5,
    fontSize: 20,
    bold: true,
    color: C.tealDark,
  });
  addText(s, "Bez poravnanja heševi pogađaju, ali ne znaju koji su u istom trenutku pesme.", {
    x: 0.82,
    y: 6.42,
    w: 11.4,
    h: 0.3,
    fontSize: 12,
    color: C.muted,
  });

  addNotes(
    "Matematička metoda radi na mel spektrogramu. Prvo izvučemo lokalne maksimume — filter 16 puta 16, prag 0,25. " +
      "Za svaki pik napravimo 15 heševa sa susedima: frekvencija pika, frekvencija cilja i vremenski pomeraj. " +
      "Ključni korak je pretraga: ne brojimo samo poklapanja, nego tražimo pomak na kome se najviše poklapaju. " +
      "To je podiglo tačnost sa oko 70 na 97,56 odsto — najveći pojedinačni doprinos u celom radu. (70 s)",
  );
}

// ---------------------------------------------------------------- 8 CNN
{
  const s = contentSlide("CNN metoda", "Embedding umesto heša");

  const steps = [
    ["Mel", "128 × ~215 frejmova"],
    ["ResNet-18", "1 kanal, modifikovan"],
    ["128-d vektor", "L2 normalizovan"],
    ["FAISS", "najbliži susedi, L2"],
  ];
  const wNode = 2.55;
  const gap = 0.45;
  steps.forEach(([head, sub], i) => {
    const x = 0.8 + i * (wNode + gap);
    addCard(s, x, 1.95, wNode, 1.1, { fill: i === 2 ? C.violetSoft : C.white, line: i === 2 ? C.violetSoft : C.line });
    addText(s, head, { x: x + 0.18, y: 2.07, w: wNode - 0.36, h: 0.32, fontSize: 15, bold: true, color: i === 2 ? C.violet : C.text });
    addText(s, sub, { x: x + 0.18, y: 2.42, w: wNode - 0.36, h: 0.5, fontSize: 11, color: C.muted });
    if (i < steps.length - 1) addArrow(s, x + wNode + 0.06, 2.47, gap - 0.12, 0);
  });

  addText(s, "treniranje: triplet loss + hard triplet mining  ·  inference: ONNX, INT8 kvantizacija", {
    x: 0.8,
    y: 3.18,
    w: 11.6,
    h: 0.3,
    fontSize: 11,
    color: C.muted,
  });

  // embedding vector + neighbours
  addText(s, "embedding (128 dimenzija)", { x: 0.82, y: 3.75, w: 4.5, h: 0.28, fontSize: 11, color: C.muted });
  for (let i = 0; i < 14; i++) {
    const hot = [2, 7, 11].includes(i);
    s.addShape(S.roundRect, {
      x: 0.82 + i * 0.31,
      y: 4.08,
      w: 0.22,
      h: 0.4,
      rectRadius: 0.03,
      fill: { color: hot ? C.amber : C.violet, transparency: hot ? 0 : 45 },
      line: { color: C.violet, width: 0 },
    });
  }

  addCard(s, 5.55, 3.75, 7.0, 1.9, { fill: C.white });
  addText(s, "Najbliži susedi u bazi", { x: 5.8, y: 3.88, w: 3.5, h: 0.3, fontSize: 12, color: C.muted });
  [
    ["Pesma A", "0,12", C.teal],
    ["Pesma B", "0,31", C.muted],
    ["Pesma C", "0,58", C.muted],
  ].forEach(([name, dist, color], i) => {
    const y = 4.2 + i * 0.46;
    s.addShape(S.ellipse, { x: 5.82, y: y + 0.08, w: 0.14, h: 0.14, fill: { color }, line: { color, width: 0 } });
    addText(s, name, { x: 6.08, y, w: 2.2, h: 0.3, fontSize: 13, bold: i === 0, color: C.text });
    addText(s, `L2 = ${dist}`, { x: 10.4, y, w: 1.8, h: 0.3, fontSize: 12, color: C.muted, align: "r" });
  });

  addText(s, "Isti podaci, ista arhitektura — razlika je šta se poredi: heševi ili vektori.", {
    x: 0.8,
    y: 5.95,
    w: 11.6,
    h: 0.4,
    fontSize: 16,
    color: C.text,
  });

  addNotes(
    "CNN grana koristi ResNet-18 prilagođen jednokanalnom mel spektrogramu. Izlaz je vektor od 128 dimenzija, normalizovan na jediničnu dužinu. " +
      "Trenira se triplet loss-om sa teškim tripletima, a za inferencu se izvozi u ONNX i kvantizuje na INT8. " +
      "Pretraga je najbliži sused po L2 rastojanju preko FAISS-a. Bitno je da su podaci isti kao u matematičkoj grani — razlika je samo u reprezentaciji. (55 s)",
  );
}

setSection("Rezultati");

// ---------------------------------------------------------------- 9 accuracy
{
  const s = contentSlide("Rezultati", "Tačnost: heširanje vodi na svakoj metriki");

  s.addChart(
    pres.ChartType.bar,
    [
      {
        name: "Matematička metoda",
        labels: ["Top-1", "Top-3", "Top-5", "Top-10"],
        values: [97.56, 98.68, 98.78, 99.04],
      },
      {
        name: "CNN embedding",
        labels: ["Top-1", "Top-3", "Top-5", "Top-10"],
        values: [81.39, 87.75, 89.75, 91.97],
      },
    ],
    {
      x: 0.8,
      y: 1.95,
      w: 7.55,
      h: 4.05,
      barDir: "col",
      chartColors: [C.teal, C.violet],
      chartColorsOpacity: 100,
      showValue: true,
      dataLabelPosition: "outEnd",
      dataLabelColor: C.text,
      dataLabelFontSize: 10,
      dataLabelFontFace: "+mn-lt",
      dataLabelFormatCode: '0.00"%"',
      catAxisLabelColor: C.text,
      catAxisLabelFontSize: 12,
      catAxisLabelFontFace: "+mn-lt",
      valAxisLabelColor: C.muted,
      valAxisLabelFontSize: 10,
      valAxisLabelFontFace: "+mn-lt",
      valAxisMinVal: 0,
      valAxisMaxVal: 108,
      valGridLine: { color: C.line, size: 0.75 },
      catGridLine: { style: "none" },
      catAxisLineColor: C.line,
      valAxisLineColor: C.line,
      showLegend: true,
      legendPos: "b",
      legendColor: C.text,
      legendFontSize: 11,
      legendFontFace: "+mn-lt",
      showTitle: false,
    },
  );

  addText(s, "MRR", { x: 8.75, y: 1.98, w: 3.8, h: 0.28, fontSize: 11, color: C.muted });
  addText(s, "0,976", { x: 8.75, y: 2.25, w: 1.8, h: 0.55, fontSize: 32, bold: true, color: C.tealDark });
  addText(s, "0,850", { x: 10.7, y: 2.25, w: 1.8, h: 0.55, fontSize: 32, bold: true, color: C.violet });
  addText(s, "matematika", { x: 8.75, y: 2.82, w: 1.8, h: 0.26, fontSize: 10, color: C.muted });
  addText(s, "CNN", { x: 10.7, y: 2.82, w: 1.8, h: 0.26, fontSize: 10, color: C.muted });

  addCard(s, 8.75, 3.3, 3.8, 1.28, { fill: C.amberSoft, line: C.amberSoft });
  addText(s, "16,17 pp", { x: 8.98, y: 3.4, w: 3.4, h: 0.62, fontSize: 36, bold: true, color: C.amberDark });
  addText(s, "razlika u top-1 tačnosti", { x: 9.0, y: 4.02, w: 3.3, h: 0.3, fontSize: 12, color: C.amberDark });

  addText(s, "7.281 zašumljeni upit  ·  776 pesama  ·  pet nivoa šuma", {
    x: 8.75,
    y: 4.78,
    w: 3.9,
    h: 0.55,
    fontSize: 10,
    color: C.muted,
  });

  addText(s, "Top-k tačnost (%); više je bolje.", {
    x: 0.82,
    y: 6.12,
    w: 6,
    h: 0.28,
    fontSize: 10,
    color: C.muted,
  });

  addNotes(
    "Glavni rezultat: 97,56 prema 81,39 odsto na top-1. Matematika vodi na svakoj metriki, a razlika je 16,17 procentnih poena. " +
      "MRR je 0,976 prema 0,850 — kad promaši prvo mesto, heširanje je i dalje mnogo bliže vrhu. " +
      "Namerno kažem „na svakoj metriki u ovom protokolu“; nemamo intervale pouzdanosti ni test značajnosti, pa ne tvrdimo da je razlika statistički značajna. (65 s)",
  );
}

// ---------------------------------------------------------------- 10 speed
{
  const s = contentSlide("Brzina i treniranje", "20× brža inferenca, trening stao na epohi 26");

  addText(s, "Vreme inferencije", { x: 0.82, y: 1.98, w: 5, h: 0.28, fontSize: 11, color: C.muted });

  const rows = [
    ["CNN", 16, C.violet, "16 ms"],
    ["Matematička", 325, C.teal, "325 ms"],
  ];
  rows.forEach(([label, ms, color, value], i) => {
    const y = 2.5 + i * 1.0;
    addText(s, label, { x: 0.82, y: y + 0.02, w: 2.6, h: 0.35, fontSize: 14, bold: true, color: C.text });
    const maxW = 3.0;
    const w = Math.max(0.18, (ms / 350) * maxW);
    s.addShape(S.roundRect, {
      x: 3.55,
      y,
      w,
      h: 0.4,
      rectRadius: 0.04,
      fill: { color },
      line: { color, width: 0 },
    });
    addText(s, value, { x: 3.55 + w + 0.14, y: y - 0.03, w: 1.05, h: 0.4, fontSize: 15, bold: true, color: C.text });
  });
  addText(s, "20-pesmeni benchmark; heš pretraga raste sa bazom, CNN pretraga preko FAISS-a ne raste linearno.", {
    x: 0.82,
    y: 4.55,
    w: 6.3,
    h: 0.6,
    fontSize: 11,
    color: C.muted,
  });

  addCard(s, 7.9, 1.95, 4.65, 3.25, { fill: C.white });
  addText(s, "Trening CNN-a", { x: 8.15, y: 2.1, w: 4, h: 0.28, fontSize: 11, color: C.muted });
  addText(s, "epoha 26", { x: 8.15, y: 2.38, w: 4.2, h: 0.75, fontSize: 40, bold: true, color: C.violet });
  addText(s, "najbolji top-1: 81,39 %", { x: 8.15, y: 3.15, w: 4.2, h: 0.32, fontSize: 14, bold: true, color: C.text });
  addText(s, "posle epohe 26 tačnost opada — preprilagođavanje.", {
    x: 8.15,
    y: 3.55,
    w: 4.2,
    h: 0.55,
    fontSize: 13,
    color: C.muted,
  });
  addText(s, "Dodatno treniranje nije put; veći model jeste.", {
    x: 8.15,
    y: 4.2,
    w: 4.2,
    h: 0.55,
    fontSize: 13,
    bold: true,
    color: C.amberDark,
  });

  addText(s, "20× brža inferenca ostavlja prostor za veći model.", {
    x: 0.8,
    y: 5.75,
    w: 11.6,
    h: 0.45,
    fontSize: 20,
    bold: true,
    color: C.text,
  });

  addNotes(
    "CNN je otprilike 20 puta brži: 16 milisekundi prema 325. Ali taj benchmark je na 20 pesama i heš pretraga raste sa bazom, pa tu brojku ne treba generalizovati. " +
      "Trening je stao na epohi 26; posle toga tačnost opada, što je klasično preprilagođavanje. Zato dodatne epohe nisu put — veći model je, i za njega ima prostora upravo zato što je inferenca tako brza. (65 s)",
  );
}

// ---------------------------------------------------------------- 11 why
{
  const s = contentSlide("Diskusija", "Zašto heširanje pobeđuje ovde");

  addCard(s, 0.8, 1.95, 5.7, 3.5);
  addText(s, "Testirali smo", { x: 1.1, y: 2.15, w: 5.1, h: 0.35, fontSize: 18, bold: true, color: C.tealDark });
  addText(
    s,
    [
      { text: "zašumljene originale na pet nivoa (−5…20 dB)", options: { bullet: true, breakLine: true } },
      { text: "776 pesama iz 16 žanrova", options: { bullet: true, breakLine: true } },
      { text: "istu bazu i isti protokol za obe metode", options: { bullet: true, breakLine: true } },
      { text: "brzinu inferencije na 20 pesama", options: { bullet: true } },
    ],
    { x: 1.1, y: 2.6, w: 5.1, h: 2.4, fontSize: 14, color: C.text, paraSpaceAfter: 9 },
  );

  addCard(s, 6.85, 1.95, 5.7, 3.5);
  addText(s, "Nismo testirali", { x: 7.15, y: 2.15, w: 5.1, h: 0.35, fontSize: 18, bold: true, color: C.violet });
  addText(
    s,
    [
      { text: "koveri i izvođenja uživo", options: { bullet: true, breakLine: true } },
      { text: "reverb i snimke iz sale", options: { bullet: true, breakLine: true } },
      { text: "razlaganje rezultata po nivou šuma", options: { bullet: true, breakLine: true } },
      { text: "milionske baze", options: { bullet: true } },
    ],
    { x: 7.15, y: 2.6, w: 5.1, h: 2.4, fontSize: 14, color: C.text, paraSpaceAfter: 9 },
  );

  addText(s, "Zato je CNN-ova prednost na koverima hipoteza, a ne rezultat.", {
    x: 0.8,
    y: 5.85,
    w: 11.6,
    h: 0.5,
    fontSize: 20,
    bold: true,
    color: C.text,
  });
  addText(s, "Na zašumljenim originalima heširanje je u svom domenu — poravnanje koristi vremensku strukturu pesme.", {
    x: 0.82,
    y: 6.35,
    w: 11.4,
    h: 0.35,
    fontSize: 12,
    color: C.muted,
  });

  addNotes(
    "Zašto onda heširanje pobeđuje? Zato što smo ga testirali na njegovom terenu: zašumljeni originali, ne koveri. " +
      "Poravnanje koristi vremensku strukturu pesme, a ne samo prisustvo heševa — zato je tako jako. " +
      "CNN-ova prednost na strukturnim mutacijama ostaje neproverena; to je pošteno reći i pre nego što nas pitaju. (50 s)",
  );
}

// ---------------------------------------------------------------- 12 future
{
  const s = contentSlide("Ograničenja i dalji rad", "Šta sledi", { dark: true });

  addText(s, "Danas: lokalne baze (dict + pickle, FAISS), bez reverba, rezultati agregirani po nivou šuma.", {
    x: 0.82,
    y: 1.95,
    w: 11.4,
    h: 0.4,
    fontSize: 13,
    color: C.mutedOnDark,
  });

  const items = [
    ["01", "Test na koverima i uživo", "ključni eksperiment za CNN hipotezu"],
    ["02", "Veći model", "ResNet-18 → MobileNetV3 ili dublja mreža"],
    ["03", "Korisnički interfejs", "snimak → rezultat, bez komandne linije"],
    ["04", "Skalabilna baza", "centralni indeks za milione pesama"],
  ];
  items.forEach(([n, head, body], i) => {
    const x = 0.8 + (i % 2) * 6.0;
    const y = 2.65 + Math.floor(i / 2) * 1.75;
    addCard(s, x, y, 5.65, 1.5, { fill: C.ink2, line: C.ink3 });
    addText(s, n, { x: x + 0.3, y: y + 0.2, w: 0.8, h: 0.4, fontSize: 22, bold: true, color: C.tealLight });
    addText(s, head, { x: x + 1.15, y: y + 0.2, w: 4.3, h: 0.35, fontSize: 16, bold: true, color: C.white });
    addText(s, body, { x: x + 1.15, y: y + 0.63, w: 4.25, h: 0.6, fontSize: 12, color: C.mutedOnDark });
  });

  addNotes(
    "Ograničenja su jasna: nismo testirali koveri i uživo, nema reverba, a baze su lokalne i nisu za milione pesama. " +
      "Zato su i sledeći koraci ovakvi: prvo test mutacija, jer to je jedina stvar koja može da promeni zaključak; onda veći model, pošto imamo prostor u latenciji; pa interfejs i skalabilna baza. (45 s)",
  );
}

setSection("Zaključak");

// ---------------------------------------------------------------- 13 conclusion
{
  const s = pres.addSlide({ masterName: "dark", sectionTitle: "Zaključak" });
  activeSlide = s;
  addKicker(s, "Zaključak", true);
  addTitle(s, "Zaključak", true);

  const stats = [
    ["97,56 %", "heširanje, top-1", C.tealLight],
    ["81,39 %", "CNN, top-1", "A99BFF"],
    ["20×", "brža inferenca", C.amber],
  ];
  stats.forEach(([num, label, color], i) => {
    const x = 0.8 + i * 4.0;
    addText(s, num, { x, y: 1.95, w: 3.6, h: 0.85, fontSize: 48, bold: true, color });
    addText(s, label, { x, y: 2.82, w: 3.6, h: 0.3, fontSize: 13, color: C.mutedOnDark });
  });

  addText(s, "Tačnije: heširanje + vremensko poravnanje.", { x: 0.82, y: 3.55, w: 11.2, h: 0.32, fontSize: 15, color: C.white });
  addText(s, "Brže: CNN inferenca, 16 ms prema 325 ms.", { x: 0.82, y: 3.95, w: 11.2, h: 0.32, fontSize: 15, color: C.white });
  addText(s, "Otvoreno: koveri i izvođenja uživo.", { x: 0.82, y: 4.35, w: 11.2, h: 0.32, fontSize: 15, color: C.white });

  addText(s, "Hvala. Pitanja?", { x: 0.8, y: 5.3, w: 8, h: 0.7, fontSize: 32, bold: true, color: C.amber });

  addNotes(
    "Da zatvorim: heširanje je tačnije, CNN je znatno brži, a prednost CNN-a na koverima je otvoreno pitanje koje tek treba testirati. " +
      "Glavni doprinos nije pobednik, nego postavljen protokol koji omogućava da se ovo poredi ponovljivo. Hvala — spremni smo za pitanja. (20 s)",
  );
}

// ---------------------------------------------------------------- appendix
function appendix(kicker, title) {
  const s = contentSlide(`Rezerva · ${kicker}`, title, { size: 30 });
  addPill(s, 11.35, 0.46, 1.2, 0.3, "Q&A", { fill: C.amberSoft, line: C.amberSoft, color: C.amberDark, fontSize: 10 });
  return s;
}

setSection("Rezervni slajdovi");

// A1 hash details
{
  const s = appendix("Heširanje", "Heš: od pikova do poravnanja");
  addText(s, "heš = (f₁, f₂, Δt)", { x: 0.82, y: 1.9, w: 6, h: 0.6, fontSize: 28, bold: true, color: C.tealDark });
  addText(
    s,
    [
      { text: "f₁ — frekvencija referentnog pika", options: { bullet: true, breakLine: true } },
      { text: "f₂ — frekvencija ciljnog pika", options: { bullet: true, breakLine: true } },
      { text: "Δt — vremenski pomeraj između njih", options: { bullet: true } },
    ],
    { x: 0.85, y: 2.65, w: 5.6, h: 1.6, fontSize: 15, color: C.text, paraSpaceAfter: 8 },
  );
  addCard(s, 0.8, 4.45, 5.7, 1.9);
  addText(s, "Parametri", { x: 1.05, y: 4.58, w: 5, h: 0.3, fontSize: 13, bold: true, color: C.text });
  addText(s, "prozor 16 × 16  ·  prag 0,25\nfanout 15  ·  ograničen Δt i Δf", {
    x: 1.05,
    y: 4.98,
    w: 5.1,
    h: 0.9,
    fontSize: 13,
    color: C.muted,
  });
  addCard(s, 6.85, 1.9, 5.7, 4.45);
  addText(s, "Zašto poravnanje", { x: 7.1, y: 2.05, w: 5, h: 0.3, fontSize: 13, bold: true, color: C.text });
  addText(s, "Dva snimka iste pesme imaju iste heševe, ali na različitim apsolutnim vremenima. Brojanje poklapanja bez pomaka meša slučajne pogotke sa stvarnim. Poravnanje traži pomak na kome se poklapa najviše heševa — i time koristi strukturu pesme, a ne samo prisustvo pikova.", {
    x: 7.1,
    y: 2.45,
    w: 5.15,
    h: 2.6,
    fontSize: 14,
    color: C.text,
  });
  addText(s, "25 % → 70 % → 97,56 %", { x: 7.1, y: 5.35, w: 5.1, h: 0.5, fontSize: 24, bold: true, color: C.amberDark });
  addText(s, "original → tuning → poravnanje", { x: 7.12, y: 5.85, w: 5, h: 0.28, fontSize: 11, color: C.muted });
  addNotes("Rezervni slajd ako pitaju kako tačno izgleda heš i zašto poravnanje menja rezultat.");
}

// A2 metrics
{
  const s = appendix("Metrike", "Šta tačno merimo");
  addText(s, "MRR = (1/N) · Σ 1/rankᵢ", { x: 0.82, y: 1.9, w: 7, h: 0.6, fontSize: 28, bold: true, color: C.tealDark });
  addText(s, "N = broj upita; rankᵢ = pozicija tačne pesme (0 ako je nema u listi).", {
    x: 0.85,
    y: 2.6,
    w: 7.2,
    h: 0.4,
    fontSize: 14,
    color: C.muted,
  });
  addText(
    s,
    [
      { text: "Top-1 / 3 / 5 / 10: da li je tačna pesma u prvih k", options: { bullet: true, breakLine: true } },
      { text: "MRR: koliko visoko je tačna pesma kada nije prva", options: { bullet: true, breakLine: true } },
      { text: "Tie-break: leksikografski po ID-u pesme", options: { bullet: true, breakLine: true } },
      { text: "Isti niz šuma za obe metode, iste isečke", options: { bullet: true } },
    ],
    { x: 0.85, y: 3.2, w: 11.2, h: 2.3, fontSize: 15, color: C.text, paraSpaceAfter: 8 },
  );
  addNotes("Rezervni slajd za definicije metrika i protokol evaluacije.");
}

// A3 genres table
{
  const s = appendix("Skup podataka", "Raspodela po žanrovima");
  const genres = [
    ["Blues", "48", "6,2"], ["Classical", "50", "6,4"], ["Country", "53", "6,8"],
    ["Easy Listening", "19", "2,4"], ["Electronic", "53", "6,8"], ["Experimental", "50", "6,4"],
    ["Folk", "55", "7,1"], ["Hip-Hop", "55", "7,1"], ["Instrumental", "45", "5,8"],
    ["International", "45", "5,8"], ["Jazz", "50", "6,4"], ["Old-Time / Historic", "42", "5,4"],
    ["Pop", "56", "7,2"], ["Rock", "55", "7,1"], ["Soul-RnB", "49", "6,3"], ["Spoken", "51", "6,6"],
  ];
  const rows = [["Žanr", "Pesama", "Udeo"]].concat(genres).concat([["Ukupno", "776", "100"]]);
  s.addTable(
    rows.map((r, i) =>
      r.map((cell, j) => ({
        text: cell,
        options: {
          fontSize: 10,
          fontFace: F,
          lang: LANG,
          color: i === 0 || i === rows.length - 1 ? C.text : C.muted,
          bold: i === 0 || i === rows.length - 1,
          align: j === 0 ? "l" : "r",
          fill: { color: i === 0 ? C.tealSoft : i === rows.length - 1 ? C.amberSoft : C.white },
        },
      })),
    ),
    {
      x: 0.8,
      y: 1.9,
      w: 5.9,
      colW: [3.4, 1.25, 1.25],
      rowH: 0.26,
      border: { type: "solid", pt: 0.5, color: C.line },
      margin: 0.03,
      valign: "mid",
    },
  );
  addText(s, "Zašto ravnomerno", { x: 7.1, y: 2.0, w: 4.9, h: 0.32, fontSize: 14, bold: true, color: C.text });
  addText(s, "Stratifikacija po 16 žanrova sprečava da model nauči jedan žanr i da rezultat bude posledica sastava skupa. Hi-kvadrat test ne odbacuje ravnomernu raspodelu (χ² = 24,21, df = 15, kritična vrednost ≈ 25,0 pri α = 0,05).", {
    x: 7.1,
    y: 2.4,
    w: 4.9,
    h: 2.2,
    fontSize: 13,
    color: C.text,
  });
  addText(s, "1000 → 776", { x: 7.1, y: 4.75, w: 4.6, h: 0.6, fontSize: 30, bold: true, color: C.tealDark });
  addText(s, "224 korumpirane pesme uklonjene MD5 proverom", { x: 7.12, y: 5.38, w: 4.7, h: 0.6, fontSize: 12, color: C.muted });
  addNotes("Rezervni slajd sa punom tabelom žanrova i objašnjenjem stratifikacije.");
}

// A4 hyperparameters
{
  const s = appendix("Hiperparametri", "Sve je fiksirano za ponovljivost");
  const rows = [
    ["Parametar", "Vrednost"],
    ["Frekvencija uzorkovanja", "22050 Hz"],
    ["Prozor STFT-a (n_fft)", "2048"],
    ["Korak prozora (hop)", "512"],
    ["Mel filteri", "128"],
    ["Isečak / preklapanje", "5,0 s / 2,5 s"],
    ["Embedding", "128"],
    ["Fanout", "15"],
    ["Prag lokalnog maksimuma", "0,25"],
    ["Nivoi šuma", "−5, 0, 5, 10, 20 dB"],
  ];
  s.addTable(
    rows.map((r, i) =>
      r.map((cell, j) => ({
        text: cell,
        options: {
          fontSize: 11,
          fontFace: F,
          lang: LANG,
          color: i === 0 ? C.text : C.muted,
          bold: i === 0,
          align: j === 0 ? "l" : "r",
          fill: { color: i === 0 ? C.tealSoft : i % 2 ? C.white : "F0F2F5" },
        },
      })),
    ),
    {
      x: 0.8,
      y: 1.9,
      w: 7.4,
      colW: [4.7, 2.7],
      rowH: 0.4,
      border: { type: "solid", pt: 0.5, color: C.line },
      margin: 0.06,
      valign: "mid",
    },
  );
  addCard(s, 8.6, 1.9, 3.95, 4.55);
  addText(s, "Zašto baš tako", { x: 8.85, y: 2.05, w: 3.5, h: 0.3, fontSize: 13, bold: true, color: C.text });
  addText(s, "22050 Hz je dovoljno za mel do ~11 kHz i duplo manje podataka od 44,1 kHz.\n\n2048 / 512 daje ~93 ms prozor i 75 % preklapanja — ~215 frejmova po isečku.\n\n5 s daje kontekst, a 2,5 s preklapanja pokriva granice.", {
    x: 8.85,
    y: 2.45,
    w: 3.45,
    h: 3.8,
    fontSize: 12,
    color: C.muted,
  });
  addNotes("Rezervni slajd sa hiperparametrima i obrazloženjem izbora.");
}

// A5 what was not done
{
  const s = appendix("Ograničenja", "Šta nije urađeno i zašto");
  addText(
    s,
    [
      { text: "Koveri i izvođenja uživo — nisu bili u skupu; bez toga CNN hipoteza ostaje neproverena.", options: { bullet: true, breakLine: true } },
      { text: "Reverb — nije implementiran; šum je pink i govor, ne impulsni odziv sale.", options: { bullet: true, breakLine: true } },
      { text: "Rezultati po nivou šuma — objavljen je agregat; razlaganje je sledeći korak.", options: { bullet: true, breakLine: true } },
      { text: "Ablacija U-Net denojzera — komponenta postoji u lancu, ali nije izmeren njen doprinos.", options: { bullet: true, breakLine: true } },
      { text: "Skalabilnost — dict + pickle i lokalni FAISS nisu za milione zapisa.", options: { bullet: true } },
    ],
    { x: 0.85, y: 2.0, w: 7.3, h: 3.6, fontSize: 15, color: C.text, paraSpaceAfter: 10 },
  );
  addCard(s, 8.55, 1.95, 4.0, 3.5, { fill: C.amberSoft, line: C.amberSoft });
  addText(s, "Prototip vs. pun skup", { x: 8.8, y: 2.1, w: 3.5, h: 0.3, fontSize: 13, bold: true, color: C.amberDark });
  addText(s, "12 pesama: 98,68 % standardno, 86,03 % pod distorzijom", { x: 8.8, y: 2.5, w: 3.5, h: 0.8, fontSize: 13, color: C.text });
  addText(s, "776 pesama: 81,39 %", { x: 8.8, y: 3.4, w: 3.5, h: 0.5, fontSize: 20, bold: true, color: C.amberDark });
  addText(s, "Nije isti protokol — ne porediti kao „pad modela“, ali smer je jasan.", { x: 8.8, y: 3.95, w: 3.5, h: 0.9, fontSize: 11, color: C.muted });
  addNotes("Rezervni slajd za pošteno nabrajanje ograničenja i kontekst prototipa.");
}

// ---------------------------------------------------------------- write
(async () => {
  const out = path.join(__dirname, "deck.pptx");
  await pres.writeFile({ fileName: out });
  await applyTheme(out, THEME);
  console.log("wrote", out);
})().catch((e) => {
  console.error(e);
  process.exit(1);
});
