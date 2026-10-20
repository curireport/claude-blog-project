// 바로바스 제품별 부품 안내 카드 생성기
// 사용법: node make_deck.js ../products/DB-4000.json
//  → ../output/<모델>_부품안내카드.pptx, ../output/<모델>_나레이션.txt
const fs = require("fs");
const path = require("path");
const pptxgen = require("pptxgenjs");
const React = require("react");
const RDS = require("react-dom/server");
const sharp = require("sharp");
const md = require("react-icons/md");
const { applyTheme } = require(process.env.APPLY_THEME || "./apply_theme.js");

const D = JSON.parse(fs.readFileSync(process.argv[2], "utf8"));
const OUT = path.resolve(__dirname, "../output");
fs.mkdirSync(OUT, { recursive: true });

const HEX = {
  navy: "12324A", navy2: "1C4762", aqua: "1FA3B8", amber: "FFB547",
  red: "E5484D", bg: "F4F8FA", slot: "E6EDF1", line: "B4C4CF", text: "1E2A33", muted: "5D707D", white: "FFFFFF",
};
const THEME = {
  name: "Barobath", headFontFace: "맑은 고딕", bodyFontFace: "맑은 고딕",
  colors: {
    dk1: HEX.text, lt1: HEX.white, dk2: HEX.navy, lt2: HEX.bg,
    accent1: HEX.aqua, accent2: HEX.amber, accent3: HEX.red, accent4: HEX.navy2,
    accent5: HEX.muted, accent6: HEX.slot, hlink: HEX.aqua, folHlink: HEX.muted,
  },
};
const SOFT = "B8CCD8"; // 어두운 배경 위 보조 글자

async function icon(Comp, color, size = 256) {
  const svg = RDS.renderToStaticMarkup(React.createElement(Comp, { color: "#" + color, size: String(size) }));
  return "image/png;base64," + (await sharp(Buffer.from(svg)).png().toBuffer()).toString("base64");
}

(async () => {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9"; // 10 x 5.625 in → PNG 1920x1080
  pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
  pres.title = `${D.model} ${D.productType} 부품 안내`;
  pres.author = "바로바스";
  const C = pres.SchemeColor;

  const I = {
    drop: await icon(md.MdWaterDrop, HEX.aqua),
    camL: await icon(md.MdPhotoCamera, "8FA3B0"),
    camD: await icon(md.MdPhotoCamera, "5F8096"),
    checkA: await icon(md.MdCheckCircle, HEX.aqua),
    checkW: await icon(md.MdCheckCircle, HEX.white),
    alert: await icon(md.MdErrorOutline, HEX.red),
    cart: await icon(md.MdShoppingCart, HEX.navy),
    play: await icon(md.MdPlayCircleFilled, HEX.red),
    info: await icon(md.MdInfoOutline, SOFT),
  };

  // 로고: 영상 워터마크(assets/워터마크_오른쪽위_1920x1080.png)와 같은 자리·크기로 오른쪽 위에 넣음
  // JSON에 "logo": false 를 넣으면 로고 없이 만듦
  const ASSETS = path.resolve(__dirname, "../assets");
  const LOGO = { x: 8.539, y: 0.286, w: 1.152, h: 0.362 };
  const logo = (file) => (D.logo === false ? [] : [{ image: { ...LOGO, path: path.join(ASSETS, file) } }]);
  pres.defineSlideMaster({ title: "LIGHT", background: { color: HEX.bg }, objects: logo("logo_color.png") });
  pres.defineSlideMaster({ title: "DARK", background: { color: HEX.navy }, objects: logo("logo_white.png") });
  const PY = D.logo === false ? 0.5 : 0.95; // 오른쪽 사진 칸 시작 높이 (로고 아래)

  // ---------- 공용 컴포넌트 ----------
  const T = (s, text, o) => s.addText(text, { isTextBox: true, margin: 0, ...o });
  function photoSlot(s, x, y, w, h, label, dark = false) {
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x, y, w, h, rectRadius: 0.12, objectName: "사진 자리",
      fill: { color: dark ? C.accent4 : C.accent6 },
      line: { color: dark ? "5F8096" : HEX.line, width: 1.25, dashType: "dash" },
    });
    const isz = Math.min(0.5, h * 0.28), cy = y + h / 2 - (isz + 0.32) / 2;
    s.addImage({ data: dark ? I.camD : I.camL, x: x + w / 2 - isz / 2, y: cy, w: isz, h: isz });
    T(s, label, { x: x + 0.1, y: cy + isz + 0.04, w: w - 0.2, h: 0.28, align: "center", fontSize: 10, color: dark ? "8FB0C4" : "7D909C" });
  }
  function card(s, x, y, w, h, name) {
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x, y, w, h, rectRadius: 0.1, objectName: name, fill: { color: C.background1 }, line: { type: "none" },
      shadow: { type: "outer", color: "12324A", opacity: 0.1, blur: 8, offset: 2, angle: 90 },
    });
  }
  function numBadge(s, x, y, d, n, fill = C.accent2, color = C.text2) {
    s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: fill }, line: { color: C.background1, width: d < 0.5 ? 1.5 : 0 }, objectName: `번호 ${n}` });
    T(s, String(n), { x, y, w: d, h: d, align: "center", valign: "middle", fontSize: d < 0.5 ? 13 : 18, bold: true, color });
  }
  const notes = [];
  function addNarration(s, label, text) {
    s.addNotes(`[나레이션]\n${text}`);
    notes.push(`■ ${label}\n${text}\n`);
  }
  const fullName = `${D.brand} ${D.model} ${D.productType}`;

  // ===== 1. 호환 모델 =====
  pres.addSection({ title: "전체 영상" });
  let s = pres.addSlide({ masterName: "DARK", sectionTitle: "전체 영상" });
  T(s, `${D.brand} ${D.productType}`, { x: 0.6, y: 0.5, w: 5, h: 0.32, fontSize: 14, bold: true, color: C.accent1 });
  T(s, D.model, { x: 0.6, y: 0.82, w: 5, h: 0.8, fontSize: 44, bold: true, color: C.background1 });
  s.addImage({ data: I.checkW, x: 0.6, y: 1.85, w: 0.26, h: 0.26 });
  T(s, "호환 모델", { x: 0.95, y: 1.8, w: 3, h: 0.36, fontSize: 16, bold: true, color: C.background1, valign: "middle" });
  const models = D.compatible.slice(0, 18);
  const cols = models.length > 10 ? 3 : 2, rowsN = Math.ceil(models.length / cols);
  const cw = (5.0 - (cols - 1) * 0.1) / cols, rowH = Math.min(0.46, (4.85 - 2.3) / rowsN - 0.1);
  models.forEach((m, i) => {
    const x = 0.6 + (i % cols) * (cw + 0.1), y = 2.3 + Math.floor(i / cols) * (rowH + 0.1);
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w: cw, h: rowH, rectRadius: 0.08, fill: { color: C.accent4 }, line: { type: "none" }, objectName: "모델 칸" });
    T(s, m, { x: x + 0.05, y, w: cw - 0.1, h: rowH, align: "center", fontSize: 14, bold: true, color: C.background1, valign: "middle", fit: "shrink" });
  });
  if (D.compatibleNote) {
    s.addImage({ data: I.info, x: 0.6, y: 5.0, w: 0.2, h: 0.2 });
    T(s, D.compatibleNote, { x: 0.88, y: 4.93, w: 4.8, h: 0.4, fontSize: 11, color: SOFT, valign: "middle" });
  }
  photoSlot(s, 5.85, PY, 3.6, (5.1 - PY - 0.15) / 2, `${D.model} 제품 사진`, true);
  photoSlot(s, 5.85, PY + (5.1 - PY - 0.15) / 2 + 0.15, 3.6, (5.1 - PY - 0.15) / 2, "모델명 확인 위치 (라벨) 사진", true);
  addNarration(s, "1. 호환 모델", D.narration.models);

  // ===== 2. 내부 구조 (narration.inside 가 있을 때만) =====
  if (D.narration.inside) {
    s = pres.addSlide({ masterName: "LIGHT", sectionTitle: "전체 영상" });
    T(s, `${D.productType} 내부 구조`, { x: 0.5, y: 0.38, w: 6, h: 0.6, fontSize: 28, bold: true, color: C.text2 });
    T(s, "번호 위치를 사진 속 부품 위로 옮겨 주세요", { x: 0.5, y: 0.98, w: 6, h: 0.3, fontSize: 12, color: C.accent5 });
    photoSlot(s, 0.5, 1.4, 5.7, 3.55, "내부 사진 (덮개를 연 모습)");
    D.parts.forEach((p, i) => numBadge(s, 0.8 + (i % 5) * 1.1, 1.7 + Math.floor(i / 5) * 2.4, 0.42, p.no));
    const lh = Math.min(0.78, 3.55 / D.parts.length);
    D.parts.forEach((p, i) => {
      const y = 1.4 + i * lh;
      card(s, 6.5, y, 3.0, lh - 0.08, "범례");
      numBadge(s, 6.6, y + (lh - 0.08) / 2 - 0.17, 0.34, p.no);
      T(s, p.shortName || p.name, { x: 7.05, y, w: 2.35, h: lh - 0.08, fontSize: 13, bold: true, color: C.text2, valign: "middle", fit: "shrink" });
    });
    addNarration(s, "2. 내부 구조", D.narration.inside);
  }

  // ===== 부품별 역할 + 고장 증상 =====
  function partSlide(p, section) {
    const s = pres.addSlide({ masterName: "LIGHT", sectionTitle: section });
    const LW = 5.6; // 왼쪽 글 영역 폭
    // 글자 수로 줄 수를 어림잡아 카드 높이를 정함 (한글은 글자 폭이 글자 크기와 비슷)
    const nLines = (t, w, fs) => Math.max(1, Math.ceil(t.length / Math.floor(((w * 72) / fs) * 0.92)));
    const lineH = (fs) => (fs / 72) * 1.28;
    numBadge(s, 0.5, 0.4, 0.5, p.no);
    T(s, `부품 ${String(p.no).padStart(2, "0")}`, { x: 1.12, y: 0.4, w: 2, h: 0.5, fontSize: 14, bold: true, color: C.accent1, valign: "middle" });
    T(s, p.name, { x: 0.5, y: 0.98, w: LW, h: 0.6, fontSize: 28, bold: true, color: C.text2, fit: "shrink" });
    // 역할
    const roles = [].concat(p.role || []), rw = LW - 1.1;
    const rLines = roles.reduce((a, r) => a + nLines(r, rw, 13), 0);
    const ry = 1.7, rh = rLines * lineH(13) + 0.24;
    card(s, 0.5, ry, LW, rh, "역할");
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.65, y: ry + 0.12, w: 0.62, h: 0.3, rectRadius: 0.15, fill: { color: C.accent1 }, line: { type: "none" } });
    T(s, "역할", { x: 0.65, y: ry + 0.12, w: 0.62, h: 0.3, align: "center", valign: "middle", fontSize: 12, bold: true, color: C.background1 });
    T(s, roles.map((r, i) => ({ text: r, options: { breakLine: i < roles.length - 1 } })),
      { x: 1.42, y: ry + 0.12, w: rw, h: rh - 0.24, fontSize: 13, color: C.text1, valign: "middle" });
    // 고장 증상
    const hy = ry + rh + 0.2;
    s.addImage({ data: I.alert, x: 0.5, y: hy + 0.05, w: 0.26, h: 0.26 });
    T(s, "고장 증상", { x: 0.85, y: hy, w: 4, h: 0.36, fontSize: 15, bold: true, color: C.accent3, valign: "middle" });
    const sy = p.symptoms.slice(0, 5), top = hy + 0.46, bottom = 5.3, gap = 0.08, sw = LW - 0.7;
    let fs = 14, hs;
    for (; fs > 11; fs--) {
      hs = sy.map((t) => nLines(t, sw, fs) * lineH(fs) + 0.16);
      if (hs.reduce((a, b) => a + b, 0) + gap * (sy.length - 1) <= bottom - top) break;
    }
    hs = sy.map((t) => nLines(t, sw, fs) * lineH(fs) + 0.16);
    let y = top;
    sy.forEach((t, i) => {
      card(s, 0.5, y, LW, hs[i], "증상");
      s.addImage({ data: I.checkA, x: 0.65, y: y + hs[i] / 2 - 0.12, w: 0.24, h: 0.24 });
      T(s, t, { x: 1.02, y, w: sw, h: hs[i], fontSize: fs, color: C.text1, valign: "middle" });
      y += hs[i] + gap;
    });
    photoSlot(s, 6.4, PY, 3.1, 5.25 - PY, `${p.shortName || p.name} 사진`);
    return s;
  }
  D.parts.forEach((p) => addNarration(partSlide(p, "전체 영상"), `${(D.narration.inside ? 2 : 1) + p.no}. ${p.name}`, p.narration));

  // ===== 엔딩 =====
  function ending(section) {
    const s = pres.addSlide({ masterName: "DARK", sectionTitle: section });
    T(s, "수리 완료!", { x: 0.6, y: 0.85, w: 5, h: 0.4, fontSize: 16, bold: true, color: C.accent1 });
    T(s, "필요한 부품은\n바로바스에서 찾으세요", { x: 0.6, y: 1.3, w: 5.2, h: 1.4, fontSize: 32, bold: true, color: C.background1, valign: "top" });
    T(s, `${fullName} 교체 부품`, { x: 0.6, y: 2.8, w: 5.2, h: 0.35, fontSize: 14, color: SOFT });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: 3.45, w: 4.6, h: 0.6, rectRadius: 0.3, fill: { color: C.background1 }, line: { type: "none" } });
    s.addImage({ data: I.cart, x: 0.82, y: 3.6, w: 0.3, h: 0.3 });
    T(s, [{ text: "네이버에서 ", options: { color: C.text1 } }, { text: "'바로바스'", options: { color: C.accent1, bold: true } }, { text: " 검색", options: { color: C.text1 } }],
      { x: 1.22, y: 3.45, w: 3.9, h: 0.6, fontSize: 15, valign: "middle" });
    s.addImage({ data: I.play, x: 0.6, y: 4.33, w: 0.32, h: 0.32 });
    T(s, "구독하고 다른 부품 교체 영상도 확인하세요", { x: 1.02, y: 4.3, w: 4.5, h: 0.38, fontSize: 13, color: SOFT, valign: "middle" });
    photoSlot(s, 5.85, PY, 3.6, 4.95 - PY, "수리 완료된 제품 사진", true);
    return s;
  }
  addNarration(ending("전체 영상"), `${(D.narration.inside ? 3 : 2) + D.parts.length}. 엔딩`, D.narration.ending);

  // ===== 부품별 단독 영상용 오프닝 =====
  pres.addSection({ title: "부품별 단독 영상" });
  notes.push("────────── 부품별 단독 영상용 ──────────\n(오프닝 → 해당 부품 증상 카드 → 교체 촬영 → 엔딩 순서로 사용)\n");
  D.parts.forEach((p) => {
    const s = pres.addSlide({ masterName: "DARK", sectionTitle: "부품별 단독 영상" });
    const nm = p.shortName || p.name;
    T(s, `${D.brand} ${D.model} ${D.productType}`, { x: 0.6, y: 1.0, w: 5, h: 0.35, fontSize: 15, bold: true, color: C.accent1 });
    T(s, [{ text: nm, options: { color: C.accent2 } }, { text: "\n교체 방법", options: { color: C.background1 } }],
      { x: 0.6, y: 1.45, w: 5.1, h: 1.6, fontSize: 38, bold: true, valign: "top", fit: "shrink" });
    T(s, "고장 증상 확인부터 교체까지 한 번에", { x: 0.6, y: 3.2, w: 5, h: 0.35, fontSize: 14, color: SOFT });
    numBadge(s, 0.6, 3.9, 0.5, p.no);
    T(s, `부품 ${String(p.no).padStart(2, "0")} · ${[].concat(p.role)[0].replace(/\.$/, "")}`, { x: 1.22, y: 3.9, w: 4.4, h: 0.5, fontSize: 13, color: SOFT, valign: "middle", fit: "shrink" });
    photoSlot(s, 5.85, PY, 3.6, 4.95 - PY, `${nm} 부품 사진`, true);
    addNarration(s, `[단독] ${nm} 오프닝`,
      `안녕하세요, 바로바스입니다. 이번 영상에서는 ${fullName}의 ${nm} 교체 방법을 알려드릴게요. 먼저 어떤 증상이 있을 때 ${nm} 교체가 필요한지부터 확인해 보겠습니다.`);
  });

  const base = path.join(OUT, `${D.model}_부품안내카드`);
  await pres.writeFile({ fileName: base + ".pptx" });
  await applyTheme(base + ".pptx", THEME);
  fs.writeFileSync(path.join(OUT, `${D.model}_나레이션.txt`), `${fullName} 부품 안내 영상 나레이션\n\n` + notes.join("\n"));
  console.log("생성 완료:", base + ".pptx");
})();
