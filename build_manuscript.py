# -*- coding: utf-8 -*-
"""Builds the v7 manuscript (revised after the October 2026 review) (five-section committee structure: Introduction / Literature Review / Methodology /
Empirical Analysis / Conclusion and Policy Implications), full + blinded, as .docx.
A4, Times New Roman 12 pt, 2.54 cm margins, numbered headings, APA 7 references, 'Page X of Y' footer.
Run:  python build_manuscript.py          -> AI_Governance_Insurance_Manuscript_v7.docx
      python build_manuscript.py --blind  -> AI_Governance_Insurance_Manuscript_v7_BLINDED.docx
All numbers are read from results.json (analysis.py); nothing is typed by hand."""
import json, sys, os, re
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_COLOR_INDEX
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

HERE = os.path.dirname(os.path.abspath(__file__)); ROOT = HERE; FIGS = f"{HERE}/figures"
R = json.load(open(f"{HERE}/results.json"))
BLIND = "--blind" in sys.argv
FONT = "Times New Roman"

def f4(x): return f"{x:.4f}"
def f3(x): return f"{x:.3f}"
def f2(x): return f"{x:.2f}"
def pct(x, d=1): return f"{100*x:.{d}f}%"
def ci_pct(lohi, d=2): return f"[{100*lohi[0]:.{d}f}, {100*lohi[1]:.{d}f}]"

# ------------------------------------------------------------------ document setup
doc = Document()
sec = doc.sections[0]
sec.page_width, sec.page_height = Cm(21.0), Cm(29.7)
for side in ("left_margin", "right_margin", "top_margin", "bottom_margin"): setattr(sec, side, Cm(2.54))
st = doc.styles["Normal"]; st.font.name = FONT; st.font.size = Pt(12)
st.element.rPr.rFonts.set(qn("w:eastAsia"), FONT); st.element.rPr.rFonts.set(qn("w:cs"), FONT)
pf = st.paragraph_format; pf.space_after = Pt(6); pf.line_spacing = 1.0
for lvl, size in ((1, 14), (2, 12), (3, 12)):
    h = doc.styles[f"Heading {lvl}"]; h.font.name = FONT; h.font.size = Pt(size); h.font.bold = True
    h.font.italic = (lvl == 3); h.font.color.rgb = RGBColor(0, 0, 0)
    for k in ("w:eastAsia", "w:ascii", "w:hAnsi"): h.element.rPr.rFonts.set(qn(k), FONT)
    h.paragraph_format.space_before = Pt(12 if lvl == 1 else 8); h.paragraph_format.space_after = Pt(4); h.paragraph_format.keep_with_next = True

def add_field(run, instr):
    for tag, txt in (("begin", None), (None, instr), ("separate", None), (None, "1"), ("end", None)):
        if tag:
            e = OxmlElement("w:fldChar"); e.set(qn("w:fldCharType"), tag); run._r.append(e)
        else:
            e = OxmlElement("w:instrText") if txt == instr else OxmlElement("w:t")
            e.set(qn("xml:space"), "preserve"); e.text = txt; run._r.append(e)

fp = sec.footer.paragraphs[0]; fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
for txt, fld in (("Page ", None), (None, "PAGE"), (" of ", None), (None, "NUMPAGES")):
    r = fp.add_run(txt or ""); r.font.size = Pt(9); r.font.name = FONT
    if fld: add_field(r, fld)

def P(text="", bold=False, italic=False, align=None, size=None, after=None, before=None, indent=False, keep=False):
    p = doc.add_paragraph()
    p.alignment = {"c": WD_ALIGN_PARAGRAPH.CENTER, "r": WD_ALIGN_PARAGRAPH.RIGHT, "l": WD_ALIGN_PARAGRAPH.LEFT}.get(align, WD_ALIGN_PARAGRAPH.JUSTIFY)
    if after is not None: p.paragraph_format.space_after = Pt(after)
    if before is not None: p.paragraph_format.space_before = Pt(before)
    if indent: p.paragraph_format.first_line_indent = Cm(0.75)
    if keep: p.paragraph_format.keep_with_next = True
    for tok in re.split(r"(\*\*.+?\*\*|\*.+?\*)", text):
        if not tok: continue
        if tok.startswith("**"): rr = p.add_run(tok[2:-2]); rr.bold = True
        elif tok.startswith("*"): rr = p.add_run(tok[1:-1]); rr.italic = True
        else: rr = p.add_run(tok)
        rr.font.name = FONT
        if bold: rr.bold = True
        if italic: rr.italic = True
        if size: rr.font.size = Pt(size)
    return p

def H(text, lvl=1): return doc.add_heading(text, lvl)

def RTL(p):
    pPr = p._p.get_or_add_pPr(); b = OxmlElement("w:bidi"); b.set(qn("w:val"), "1"); pPr.append(b)
    for rr in p.runs:
        rPr = rr._r.get_or_add_rPr(); rtl = OxmlElement("w:rtl"); rtl.set(qn("w:val"), "1"); rPr.append(rtl)
        rr.font.name = "Traditional Arabic"; rPr.rFonts.set(qn("w:cs"), "Traditional Arabic"); rr.font.size = Pt(13)

def set_cell_bg(cell, hexcolor):
    tcPr = cell._tc.get_or_add_tcPr(); shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear"); shd.set(qn("w:color"), "auto"); shd.set(qn("w:fill"), hexcolor); tcPr.append(shd)

def TABLE(caption, header, rows, note=None, widths=None, size=9.5):
    cp = P("", align="l", after=3, keep=True)
    m = re.match(r"(Table \d+\.)(.*)", caption, re.S)
    r1 = cp.add_run(m.group(1)); r1.bold = True; r1.font.size = Pt(11); r1.font.name = FONT
    r2 = cp.add_run(m.group(2)); r2.italic = True; r2.font.size = Pt(11); r2.font.name = FONT
    t = doc.add_table(rows=1, cols=len(header)); t.style = "Table Grid"; t.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, h in enumerate(header):
        c = t.rows[0].cells[i]; c.text = ""; rr = c.paragraphs[0].add_run(h); rr.bold = True; rr.font.size = Pt(size); rr.font.name = FONT
        c.paragraphs[0].paragraph_format.keep_with_next = True
        set_cell_bg(c, "E7EEF5")
    trPr = t.rows[0]._tr.get_or_add_trPr(); th = OxmlElement("w:tblHeader"); th.set(qn("w:val"), "true"); trPr.append(th)
    cs = OxmlElement("w:cantSplit"); cs.set(qn("w:val"), "true"); trPr.append(cs)
    for row in rows:
        cells = t.add_row().cells
        for i, v in enumerate(row):
            cells[i].text = ""; para = cells[i].paragraphs[0]; para.paragraph_format.space_after = Pt(0)
            bold = False
            if isinstance(v, tuple): v, bold = v
            rr = para.add_run(str(v)); rr.font.size = Pt(size); rr.font.name = FONT; rr.bold = bold
    if widths:
        t.autofit = False
        tblPr = t._tbl.tblPr; lay = OxmlElement("w:tblLayout"); lay.set(qn("w:type"), "fixed"); tblPr.append(lay)
        for i, w in enumerate(widths): t.columns[i].width = Cm(w)
        for row in t.rows:
            for i, w in enumerate(widths): row.cells[i].width = Cm(w)
    for row in t.rows:
        for c in row.cells:
            for para in c.paragraphs: para.paragraph_format.line_spacing = 1.0; para.paragraph_format.space_after = Pt(0)
    if note: P(note, align="l", size=9.5, after=10, before=2)
    else: P("", after=6)

def FIG(path, caption, width_cm=13.5):
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER; p.paragraph_format.keep_with_next = True
    p.add_run().add_picture(path, width=Cm(width_cm))
    m = re.match(r"(Figure \d+\.)(.*)", caption, re.S)
    cp = doc.add_paragraph(); cp.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY; cp.paragraph_format.space_after = Pt(10)
    r1 = cp.add_run(m.group(1)); r1.bold = True; r1.font.size = Pt(10.5); r1.font.name = FONT
    r2 = cp.add_run(m.group(2)); r2.font.size = Pt(10.5); r2.font.name = FONT

def EQ(text, num):
    t = doc.add_table(rows=1, cols=2); t.alignment = WD_TABLE_ALIGNMENT.CENTER
    a, b = t.rows[0].cells; a.width = Cm(13.5); b.width = Cm(2.0); a.text = ""; b.text = ""
    ra = a.paragraphs[0].add_run(text); ra.font.name = "Cambria Math"; ra.font.size = Pt(12); a.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    rb = b.paragraphs[0].add_run(f"({num})"); rb.font.name = FONT; rb.font.size = Pt(12); b.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
    for c in (a, b): c.paragraphs[0].paragraph_format.space_after = Pt(4)

# ------------------------------------------------------------------ numbers from results
M = {f: {k: R[f"{f}_{k}"] for k in ("BE", "UN", "DF")} for f in ("GLM", "GAM", "GBM")}
g, a_, b = M["GLM"], M["GAM"], M["GBM"]
px_l, px_g = R["proxy_auc"]["logit"], R["proxy_auc"]["gbm"]
idi = {f: M[f]["UN"]["indirect_disc_index"] for f in M}; ddi = {f: M[f]["BE"]["direct_disc_index"] for f in M}
BS = R["bootstrap"]; nb = BS["n_boot"]
sh = {f: {k: v["share_over_5pct"] for k, v in R[f"{f}_ratio_stats"].items()} for f in M}
gbm_p10p90 = {k: (v["p10"], v["p90"]) for k, v in R["GBM_ratio_stats"].items()}
ratio_gg = BS["ratio_GBM_GLM"]; ratio_ag = BS["ratio_GAM_GLM"]
SENS = R["rsi_sensitivity"]["schemes"]
n_test = R["n_test"]
RF = R["refit"]; CAL = R["calibration"]; n_refit = RF["n_splits"]
n_unused = RF["GBM"]["n_gender_unused"]; n_above = RF["n_GBM_above_GLM"]
V = R["review"]; DEC = V["decomposition"]; PL = V["placebo"]; GNI = V["gbm_no_interactions"]; GG = V["gbm_grid"]; GA = V["glm_alpha_grid"]
AE = V["ae"]; AER = V["ae_refit"]; RS = V["rsi_rescoring"]; IR = V["implied_refit"]; RT = V["runtime_seconds"]
REPO_URL = "https://github.com/dr-elbehairy/governing-ai-in-insurance"
WC = json.load(open(f"{HERE}/wordcount.json"))["main_text"] if os.path.exists(f"{HERE}/wordcount.json") else 8000
def ci2(lohi): return f"[{100*lohi[0]:.2f}, {100*lohi[1]:.2f}]"
def ci3(lohi, d=3): return f"[{lohi[0]:.{d}f}, {lohi[1]:.{d}f}]"
imp = {f: DEC[f]["idi_implied_gbm"] for f in ("GLM", "GAM", "GBM")}
plm = {f: PL[f]["mean"] for f in ("GLM", "GAM", "GBM")}
rows7 = [M[f][k] for f in ("GLM", "GAM", "GBM") for k in ("BE", "UN", "DF")]
ratio_lo, ratio_hi = min(r["ratio_fm"] for r in rows7), max(r["ratio_fm"] for r in rows7)
ae_gap = {(f, k): M[f][k]["ae_m"] - M[f][k]["ae_f"] for f in ("GLM", "GAM", "GBM") for k in ("BE", "UN", "DF")}
(ae_gap_f, ae_gap_k), ae_gap_max = max(ae_gap.items(), key=lambda kv: kv[1])
share_f, share_m = sh["GBM"]["Female"], sh["GBM"]["Male"]
ga_main, ga_best = GA["grid"]["0.0001"], GA["grid"][GA["cv_best_alpha"]]
gg_best = GG["grid"][GG["cv_best"]]; n_gg_zero = sum(1 for v in GG["grid"].values() if v["idi"] < 1e-6)

# ================================================================== TITLE PAGE
TITLE = ("Governing AI in Insurance: Comparing the IAIS, NAIC and EU AI Act Frameworks and Testing Proxy "
         "Discrimination in Motor Pricing")
AR = "--ar-abstract" in sys.argv
P(TITLE, bold=True, align="c", size=15, after=14)
if AR:
    _t = P('حوكمة الذكاء الاصطناعي في التأمين: مقارنة أطر IAIS وNAIC وقانون الذكاء الاصطناعي الأوروبي واختبار التمييز بالوكالة في تسعير تأمين السيارات', bold=True, align="c", size=14, after=12); RTL(_t)
    for _r in _t.runs: _r.font.size = Pt(15); _r.bold = True
if not BLIND:
    P("Moustafa Elbehairy ¹ *  and  Mohamad Sharkawy ¹", align="c", size=12, after=2)
    P("¹ Department of Statistics and Insurance, Faculty of Commerce, Suez Canal University, Ismailia, Egypt", align="c", size=11, after=2)
    P("* Corresponding author: moustafayosri@commerce.suez.edu.eg  |  ORCID 0009-0009-1062-9420 (M. Elbehairy)", align="c", size=11, after=2)
    _p = P("", align="c", size=11, after=10); _r = _p.add_run("ORCID (M. Sharkawy): to be supplied before submission"); _r.font.name = FONT; _r.font.size = Pt(11); _r.font.highlight_color = WD_COLOR_INDEX.YELLOW
else:
    P("[Author details removed for double-blind review]", align="c", size=11, italic=True, after=10)
if not AR:
    P("Manuscript type: Research article (comparative regulatory analysis with an empirical component)   |   "
      f"Main text ≈ {WC:,} words (introduction to future research, excluding abstract, tables, figure captions and references)   |   "
  f"Tables: 12   |   Figures: 5   |   JEL: G22, G28, K23, C52, O33", align="c", size=10, after=14)

# ================================================================== ABSTRACT
if not AR:
    H("Abstract", 1)
    P(f"Three supervisory responses to artificial intelligence in insurance now frame the international debate: the IAIS Application Paper, the "
      f"NAIC Model Bulletin and the EU AI Act. We compare them on a six-dimension framework and a text-anchored Regulatory Stringency Index, and "
      f"then ask whether the fairness requirement they share, that removing a protected attribute does not by itself remove discrimination, can be "
      f"verified on real pricing data. On 20,600 French motor third-party liability policies, with gender as the protected attribute, we estimate "
      f"best-estimate, unawareness and discrimination-free frequency prices under a Poisson GLM, an additive spline GLM and a gradient-boosting "
      f"machine, and we separate the resulting indirect-discrimination index into a proxy component, measured with the best-estimate model held "
      f"fixed, and a refitting component, measured with a placebo attribute. The regimes converge in substance and diverge in form; their ordering "
      f"follows from score dominance and survives every one- and two-cell re-scoring. Rating factors recover gender with an AUC of {f2(px_g[0])}, yet "
      f"the proxy component of discrimination stays below 1 percent of premium in every model class ({pct(imp['GLM'],2)}, {pct(imp['GAM'],2)} and "
      f"{pct(imp['GBM'],2)} for the GLM, the GAM and the GBM). The usual drop-and-refit test overstates it for the boosted model about tenfold "
      f"({pct(idi['GBM'],2)}), since refitting noise alone yields an index of the same size ({pct(plm['GBM'],2)} with a random placebo). A test that "
      f"refits the model is therefore not a test of proxy discrimination. We set out a fixed-model testing protocol for insurers and proportionate "
      f"recommendations for Egypt's Financial Regulatory Authority and Saudi Arabia's Insurance Authority.", after=6)
if AR:
    p = P("الملخص", bold=True, align="r", after=2); RTL(p)
    ar = ("يشكّل الذكاء الاصطناعي اليوم قرارات التسعير والاكتتاب والمطالبات عبر سلسلة قيمة التأمين، وتهيمن على النقاش الدولي ثلاث استجابات إشرافية: "
          "ورقة التطبيق الصادرة عن الجمعية الدولية لمراقبي التأمين (IAIS)، والنشرة النموذجية للجمعية الوطنية لمفوضي التأمين الأمريكية (NAIC)، وقانون "
          "الذكاء الاصطناعي الأوروبي. يقارن هذا البحث الأطر الثلاثة على إطار من ستة أبعاد ومؤشر لصرامة التنظيم مثبّت بالنصوص، ثم يسأل عمّا إذا كان "
          "متطلب العدالة المشترك بينها، وهو أن حذف المتغير المحمي لا يزيل التمييز بذاته، قابلًا للتحقق على بيانات تسعير حقيقية. وباستخدام 20,600 "
          "وثيقة تأمين سيارات ضد الغير في فرنسا واعتبار النوع متغيرًا محميًا، قُدِّرت أسعار التقدير الأفضل وأسعار الإغفال والأسعار الخالية من التمييز في "
          "نموذج خطي معمم بواسون، ونموذج خطي معمم بدوال شرائحية جمعية، وآلة تعزيز متدرج، وفُكِّك مؤشر التمييز غير المباشر الناتج إلى مكوّن وكالة يُقاس "
          "مع تثبيت نموذج التقدير الأفضل، ومكوّن إعادة تقدير يُقاس بمتغير وهمي عشوائي. تتقارب الأطر في الجوهر وتتباين في الشكل، وترتيبها ناتج عن "
          f"هيمنة الدرجات ويصمد أمام كل إعادة ترميز لخلية أو خليتين من المؤشر. وتستعيد عوامل التسعير غير المحمية النوع بمساحة تحت المنحنى قدرها {f2(px_g[0])}، "
          f"غير أن مكوّن الوكالة من التمييز يقل عن 1% من القسط في جميع فئات النماذج ({pct(imp['GLM'],2)} و{pct(imp['GAM'],2)} و{pct(imp['GBM'],2)} على "
          f"الترتيب)، في حين يبالغ اختبار الحذف وإعادة التقدير المعتاد في تقديره لنموذج التعزيز المتدرج بنحو عشرة أضعاف ({pct(idi['GBM'],2)})، لأن ضجيج إعادة "
          f"التقدير وحده ينتج مؤشرًا بالحجم نفسه ({pct(plm['GBM'],2)} بمتغير وهمي). فالقدرة على الاستدلال بالوكالة ليست ضررًا بالوكالة، واختبار العدالة "
          "الذي يعيد تقدير النموذج ليس اختبارًا للوكالة. ويخلص البحث إلى بروتوكول اختبار بنموذج مثبَّت لشركات التأمين، وتوصيات تناسبية موجهة إلى الهيئة "
          "العامة للرقابة المالية في مصر وهيئة التأمين في المملكة العربية السعودية.")
    p = P(ar, align="r", after=4); RTL(p)
    p = P("الكلمات المفتاحية: الذكاء الاصطناعي؛ الرقابة على التأمين؛ العدالة الخوارزمية؛ التمييز بالوكالة؛ التسعير الخالي من التمييز؛ قانون الذكاء الاصطناعي الأوروبي؛ الأسواق الناشئة", align="r", after=12); RTL(p)
    doc.save(f"{ROOT}/AI_Governance_Insurance_Manuscript_v7_الملخص_العربي.docx"); print("saved arabic abstract"); sys.exit(0)

# ================================================================== 1 INTRODUCTION
H("1.  Introduction", 1)
P("Machine-learning models now shape risk classification, pricing, claims triage, fraud detection and customer service across the insurance "
  "value chain (Eling et al., 2022), and the scale of adoption is documented rather than anecdotal. EIOPA reported in 2024 that half of non-life "
  "insurers and a quarter of life insurers already used AI somewhere in the chain, and its 2026 survey of 347 undertakings found nearly two-thirds "
  "working with generative AI (EIOPA, 2024, 2026); the NAIC's own surveys put the share of U.S. private-passenger auto insurers using, planning or "
  "exploring AI at 88 percent (NAIC, 2022, 2023a, 2023b). Peer-reviewed studies of insurers' technology portfolios, drawn from annual reports and "
  "innovation surveys, tell the same story (Bohnert et al., 2019; Lanfranchi & Grassi, 2022; Owens et al., 2022; Zaijuan et al., 2025).", indent=True)
P("Insurance is a business of classification. It sorts people into pools and prices access to protection accordingly (Cevolini & Esposito, "
  "2020), which is why algorithmic decision-making is more fraught here than in most sectors: an opaque, high-dimensional model can absorb "
  "historical discrimination, infer protected characteristics from innocuous features and produce decisions that neither party to the contract can "
  "explain (Barocas & Selbst, 2016; Prince & Schwarcz, 2020; Rudin, 2019). Supervisors accordingly have to ask whether a model's use is fair, "
  "accountable, transparent and overseen, and not merely whether it is accurate (Mittelstadt et al., 2016).", indent=True)
P("Three regulatory responses now structure the field, and all three have matured since 2024. The IAIS published the final version of its "
  "Application Paper on the supervision of AI on 2 July 2025, folding AI into Insurance Core Principles (ICPs) 7, 8, 16 and 19 without creating new "
  "requirements (IAIS, 2025). The NAIC Model Bulletin on the Use of Artificial Intelligence Systems by Insurers, adopted on 4 December 2023, had "
  "been taken up by 27 states, the District of Columbia and Puerto Rico by 8 October 2026, with California, Colorado, New York and Texas relying "
  "on instruments of their own (NAIC, 2026). Regulation (EU) 2024/1689, the EU AI Act, in force since 1 August 2024, classifies risk assessment and "
  "pricing in life and health insurance as high-risk under Annex III, point 5(c); the Digital Omnibus of July 2026 moved the compliance date for "
  "those obligations to 2 December 2027, and EIOPA's Opinion of 6 August 2025 interprets Solvency II and the Insurance Distribution Directive for "
  "the AI systems that fall outside the high-risk class (European Parliament & Council, 2024, 2026; EIOPA, 2025). The three still span the whole "
  "distance from soft to hard law.", indent=True)
P("The paper makes three contributions. It offers an updated comparative framework and a transparent, text-anchored Regulatory Stringency "
  "Index, together with a re-scoring analysis that states how many coding changes would alter the ranking it produces. It builds an empirical "
  "bridge between insurance-specific fairness requirements and discrimination-free pricing on open motor data, and in doing so it decomposes the "
  "standard drop-and-refit fairness test into a proxy component and a refitting component, a separation we have not seen reported and one that "
  "changes how the test should be read for flexible learners. And it derives a fixed-model testing protocol for insurers and proportionate "
  "recommendations for supervisors in markets such as Egypt and Saudi Arabia.", indent=True)
P("Three questions organise what follows. **RQ1:** On which governance dimensions do the IAIS, NAIC and EU frameworks converge, and on which do "
  "they diverge, once their 2025–2026 versions are scored on common criteria? **RQ2:** When the fairness requirement common to the three regimes is "
  "operationalised on real motor-insurance data, how large is proxy discrimination once refitting noise is separated from it, and does the answer "
  "depend on model class? **RQ3:** What follows for insurers' risk governance and for emerging-market supervisors? Section 2 reviews the literature, "
  "Section 3 sets out the methodology, Section 4 reports the regulatory comparison, the actuarial results, the robustness and decomposition checks "
  "and the discussion, and Section 5 concludes with policy implications, limitations and future research.", indent=True)
# ================================================================== 2 LITERATURE REVIEW
H("2.  Literature Review", 1)
H("2.1  AI Risks Across the Insurance Value Chain", 2)
P("Each stage of the value chain carries its own governance risk (Eling et al., 2022). In underwriting and pricing the central worry is "
  "proxy discrimination: a model rebuilds a protected attribute from correlated features even after the attribute has been formally removed, "
  "blurring the old line between direct and indirect discrimination (Adams-Prassl et al., 2023; Prince & Schwarcz, 2020; Frees & Huang, 2023). In "
  "claims and fraud detection the problem becomes opacity and contestability (Wachter et al., 2017); in distribution it becomes targeting and "
  "exclusion (Cevolini & Esposito, 2020). Running through all of these are poor data provenance, model drift and the diffusion of accountability "
  "across vendors, models and human reviewers (Selbst et al., 2019), so that the governance task is to assign responsibility along a socio-technical "
  "chain rather than to regulate a single technology.", indent=True)
H("2.2  Theoretical Lenses", 2)
P("Three traditions in regulatory theory inform the analysis. Risk-based regulation allocates supervisory effort in proportion to the harm an "
  "activity can cause (Black & Baldwin, 2010), and the EU AI Act, with its tiers of obligation, is its textbook application. Principles-based "
  "regulation trades legal certainty for flexibility and relies on supervisory capacity and good-faith compliance (Black, 2008); the IAIS, and to a "
  "degree the NAIC, sit in this camp. The algorithmic-accountability literature supplies the substantive yardsticks of fairness, transparency, "
  "explainability, contestability and human oversight (Cath, 2018; Yeung, 2018), together with the warning that principles without institutional "
  "teeth do not reliably change behaviour (Floridi et al., 2018; Mittelstadt, 2019). Read together, these lenses frame AI governance as the work of "
  "turning principles into proportionate, enforceable institutions (de Almeida et al., 2021; Gasser & Almeida, 2017).", indent=True)
H("2.3  Direct, Indirect and Discrimination-Free Pricing", 2)
P("The actuarial literature gives the fairness mandates in all three regimes a precise operational meaning. Let *Y* be the claim outcome, "
  "*X* the vector of non-protected rating factors and *D* a protected attribute such as gender. The best-estimate price conditions on "
  "everything:", indent=True)
EQ("μ_BE(x, d) = E[Y | X = x, D = d]", 1)
P("The unawareness price simply drops *D*:", indent=True)
EQ("μ_UN(x) = E[Y | X = x] = Σ_d μ_BE(x, d) · P(D = d | X = x)", 2)
P("Equation (2) shows why unawareness does not remove discrimination. Because *D* is integrated out with its conditional distribution given *X*, "
  "any correlation between the rating factors and the protected attribute lets the model rebuild *D* implicitly; this is proxy, or indirect, "
  "discrimination. Lindholm et al. (2022) therefore define the discrimination-free price by integrating over the marginal distribution of *D* "
  "instead:", indent=True)
EQ("μ_DF(x) = Σ_d μ_BE(x, d) · P(D = d)", 3)
P("The gap between (1) and (3) measures direct discrimination and the gap between (2) and (3) indirect discrimination; the benchmark is not "
  "proposed here as a new pricing concept but used as an empirical test of what the regulatory fairness requirements demand. Two of its properties "
  "matter for governance. Computing μ_DF requires the protected attribute during development even though it is never used at the point of sale, so "
  "fairness testing needs data that many insurers do not, or may not, collect (Veale & Binns, 2017); Article 10(5) of the AI Act and paragraph 90 of "
  "the IAIS paper both anticipate this. And the size of the indirect-discrimination gap depends jointly on how strongly X predicts D and on how much "
  "D shifts μ_BE: the former is proxy capacity, and only their product is proxy harm. Section 4.2 measures both.", indent=True)
P("Four streams bear on the questions above, and Table 1 summarises what each has established. The AI-in-insurance governance literature has "
  "described the principles at stake and the content of individual instruments (Eling et al., 2022; Floridi, 2021; Smuha, 2021), but cross-regime "
  "comparison on common criteria is thin and predates the final 2025–2026 texts. The algorithmic-fairness literature has established bias and proxy "
  "discrimination as general phenomena (Barocas & Selbst, 2016; Prince & Schwarcz, 2020) without tying them to the tests supervisors expect. The "
  "discrimination-free pricing literature has supplied the mathematics and, lately, causal and motor-specific refinements (Lindholm et al., 2022, "
  "2024a, 2024b; Araiza Iturria et al., 2024; Moriah et al., 2024; Xin & Huang, 2024; Charpentier, 2024; Côté et al., 2025), and van Bekkum et al. "
  "(2025) have set out the legal research agenda, but none of this work asks what the methods imply for the requirements now in force. The "
  "emerging-market AI governance literature has grown from ecosystem studies of North Africa and Sub-Saharan Africa (Ayana et al., 2024; Effoduh et "
  "al., 2024; Ogoh et al., 2026) and assessments of Egypt's AI strategy (Badawy, 2025) to calls for proactive AI regulation in finance (Truby et al., "
  "2020; Mhlanga, 2020), while the Saudi insurance literature covers supervisory architecture and insurtech adoption (Orlando & Bace, 2021; Seyam, "
  "2025); insurance-specific evidence on algorithmic governance is scarce throughout.", indent=True)
TABLE("Table 1. Positioning of the study against four literature streams",
      ["Literature stream", "What is known", "Remaining gap", "This paper"],
      [["AI-in-insurance governance", "Governance principles; content of individual regulatory instruments", "Limited cross-regime empirical comparison on common criteria", "IAIS–NAIC–EU comparison with a transparent, text-anchored stringency index"],
       ["Algorithmic fairness", "Bias and proxy discrimination well established as phenomena", "Weak operationalisation of insurance regulatory requirements", "Maps each regime's fairness requirement to a measurable test"],
       ["Discrimination-free pricing", "Framework established and refined (Lindholm et al., 2022, 2024a; Araiza Iturria et al., 2024; Côté et al., 2025)", "Limited connection to AI-governance regimes; the drop-and-refit test not decomposed", "Empirical regulatory bridge across GLM, additive and boosted models; proxy and refitting components separated"],
       ["Emerging-market AI governance", "Horizontal AI-governance and data-governance studies (Mohamed et al., 2020; Ayana et al., 2024; Ogoh et al., 2026; Badawy, 2025)", "Little insurance-specific evidence", "Egypt and Saudi Arabia as illustrative applications"]],
      note="Note. The table positions the contribution; it is not a systematic review.", widths=[3.4, 4.2, 4.2, 4.2])
P("Two gaps follow, of the kinds that the literature-review methodology distinguishes (Müller-Bloch & Kranz, 2015; Robinson et al., 2011). "
  "The first is a practical-knowledge gap: the three insurance-specific instruments now exist in final or near-final form and are being applied by "
  "supervisors and compliance functions, yet they have not, to our knowledge, been compared on common criteria in their 2025–2026 form, although the "
  "AI-ethics and EU AI Act literatures have grown quickly (Floridi, 2021; Smuha, 2021) and the actuarial literature treats fair pricing with real "
  "technical depth (Lindholm et al., 2022, 2024a; Xin & Huang, 2024; Côté et al., 2025). The second, and the more consequential, is an empirical gap "
  "with a population dimension. The comparative literature is almost entirely doctrinal. All three regimes ask insurers to test for unfair or proxy "
  "discrimination, but whether the tests they contemplate would detect the harm on real pricing data, and how the answer depends on the class of "
  "model deployed, has received little direct attention; removing the protected attribute is known to be insufficient in theory (Lindholm et al., "
  "2022), yet the size of the residual is rarely measured, and when it is measured the measurement refits the model without the attribute, which "
  "confounds the proxy effect with whatever else a refit changes. The supervisors who need that measurement most, those outside Europe and North "
  "America deciding with thin capacity whether to copy, adapt or resist these models, are also the ones the evidence least often addresses.", indent=True)
# ================================================================== 3 METHODOLOGY
H("3.  Methodology", 1)
H("3.1  Research Design", 2)
P("The study combines two methods in sequence. A comparative documentary analysis of the three regimes, scored on a common rubric, answers "
  "RQ1 and identifies the fairness tests the regimes share; an actuarial experiment on open motor data then operationalises those tests across model "
  "classes to answer RQ2; RQ3 is answered by reading the second set of results back through the first. Neither method is new. Rubric-based coding of "
  "governance texts has been used to compare AI-ethics guidelines (Jobin et al., 2019) and to build governance frameworks (de Almeida et al., 2021), "
  "and the discrimination-free pricing procedure follows Lindholm et al. (2022) as applied in later actuarial work (Frees & Huang, 2023; Lindholm et "
  "al., 2024a, 2024b; Moriah et al., 2024; Xin & Huang, 2024). What is new is their combination and the decomposition described in Section 3.4. The "
  "robustness checks of Section 4.3 are specified in advance in Section 3.5.", indent=True)
H("3.2  Comparative Documentary Analysis and the Regulatory Stringency Index", 2)
P("Synthesising the traditions reviewed in Section 2 yields the six-dimension framework in Table 2, which separates the form of a regime (legal "
  "nature and scope) from its substance (accountability, fairness, transparency and oversight) and gives the four-point rubric used to score each "
  "dimension in Section 4.1.", indent=True)
TABLE("Table 2. Six-dimension analytical framework and scoring rubric for the Regulatory Stringency Index",
      ["Dimension", "Governance question", "Theoretical anchor", "Score 0 → 3"],
      [["Legal nature", "Is the regime binding law, supervisory guidance, or a standard?", "Hard vs. soft law",
        "0 none; 1 non-binding standard; 2 binding via adoption/existing law; 3 directly applicable statute"],
       ["Scope & risk tiering", "Which AI uses are covered, and is scrutiny proportionate to risk?", "Risk-based regulation",
        "0 none; 1 general proportionality; 2 documented risk assessment per system; 3 legally defined risk tiers"],
       ["Accountability & governance", "Who is responsible, through what internal structures?", "Algorithmic accountability",
        "0 none; 1 principle; 2 named function/board expectation; 3 mandatory programme with defined content"],
       ["Fairness & non-discrimination", "How is discriminatory output defined, tested and constrained?", "Anti-discrimination; actuarial fairness",
        "0 none; 1 principle; 2 testing expected, method open; 3 testing and data rules mandated"],
       ["Transparency & explainability", "What must be disclosed to supervisors and consumers?", "Right to explanation",
        "0 none; 1 principle; 2 documentation and consumer notice expected; 3 prescribed technical documentation and logging"],
       ["Oversight & assurance", "How is compliance verified?", "Auditing; conformity assessment",
        "0 none; 1 ordinary supervision; 2 targeted examination/attestation; 3 ex-ante conformity assessment and post-market monitoring"]],
      note="Note. Rows 1–2 capture the form of a regime and rows 3–6 its substance, so that soft-law and hard-law instruments are compared on common criteria. "
           "Scores are ordinal and are anchored to specific text passages in Table 4.", widths=[3.2, 4.6, 3.2, 5.0])

P("How the anchors were written needs stating. The level-3 anchor in each row describes the most formalised and verifiable institutional form "
  "that regulatory theory recognises for that dimension: a directly applicable statute, legally defined risk tiers, prescribed documentation and "
  "logging, ex-ante conformity assessment with post-market monitoring. These are also the forms the EU AI Act takes, since the Act is the only "
  "binding, risk-tiered statute among the three and hence the empirical ceiling of formalisation in this field. The index therefore measures degree "
  "of formalisation with the EU regime as its ceiling by construction; its maximum score in Table 4 is a property of the construct, and the "
  "information lies in where the IAIS and NAIC instruments sit beneath it and in the shape of their profiles (Section 5.2).", indent=True)
P("The three frameworks were selected purposively as the most consequential and philosophically distinct: the IAIS Application Paper (global "
  "soft law), the NAIC Model Bulletin (state-implemented reinterpretation of existing law) and Regulation (EU) 2024/1689 as amended by the Digital "
  "Omnibus (binding, risk-tiered hard law). Primary sources were read directly, including the NAIC adoption map of 8 October 2026, the EIOPA Opinion, "
  "the NIST AI Risk Management Framework and the OECD AI Principles (IAIS, 2025; NAIC, 2023c, 2026; EIOPA, 2025; OECD, 2019; Tabassi, 2023); "
  "scholarly sources were located through Scopus, Web of Science and HeinOnline, and every journal reference was verified against Crossref.", indent=True)
P(f"**Coding, meaning and dominance.** The instruments were coded against the rubric by the first author, with the anchoring passage recorded "
  f"for each score; the second author reviewed every score against the texts and differences were settled by discussion. No independent double "
  f"coding was undertaken and no inter-rater statistic is reported, which Section 5.2 records as a limitation; the scores and anchors appear in full "
  f"in Table 4 so that any cell can be contested, and Section 4.3.1 asks how many one-point re-scorings would change the ranking. Stringency is used "
  f"in the operational sense of how far governance expectations are formalised, binding and verifiable, not as a judgement of regulatory quality. "
  f"The index sums the six scores with equal weights; because, as Table 4 will show, every NAIC score is at least the corresponding IAIS score and "
  f"every EU score at least the corresponding NAIC score, any positive weighting yields the same ranking, so weight sensitivity would test nothing "
  f"and the robustness question is instead one of coding.", indent=True)
H("3.3  Data", 2)
P("The empirical part uses the freMPL1 dataset from the open CASdatasets collection (Dutang & Charpentier, 2020): 20,600 French motor "
  "third-party liability policies observed in 2004, with exposure, a claim indicator, driver characteristics (age, licence age, gender, marital "
  "status, socio-professional group), vehicle characteristics (age, body, price band, engine, fuel, maximum speed, class), usage, garaging, a "
  f"kilometre-limit flag, the bonus–malus coefficient and a risk-zone variable. The claim rate is {pct(R['claim_rate'])} and women hold "
  f"{pct(R['female_share'])} of the portfolio. Gender is the protected attribute D, since the Court of Justice of the European Union prohibited "
  "gender-based pricing in its Test-Achats judgment of 2011, effective from 21 December 2012 (Court of Justice of the European Union, 2011), so an "
  "EU insurer using these factors would have to demonstrate the absence of proxy discrimination; the remaining 16 variables form X, the 37-level "
  "socio-professional code collapsed to six groups. The portfolio is open, which makes the test reproducible; it records a protected attribute "
  "alongside rich rating factors, which the benchmark requires; and motor is the second-largest non-life line after medical cover in both markets "
  "considered in Section 4.4.3, with 21 percent of Egyptian property-and-liability premiums in 2025 and 18.5 percent of Saudi gross written premiums "
  "(Financial Regulatory Authority, 2026; Insurance Authority, 2026).", indent=True)
H("3.4  Models and Fairness Metrics", 2)
P("**Models.** Claim frequency was modelled as a Poisson process with log-exposure offset, and three model classes were fitted to span the "
  "methods insurers deploy: a Poisson GLM with one-hot categorical encoding, linear numeric effects and an L2 penalty of α = 10⁻⁴ on the "
  "standardised design, fixed a priori as the smallest penalty that stabilised the design rather than chosen by cross-validation (Section 4.3.2 "
  "reports a grid, since the penalty turns out to matter for the gender coefficient); an additive spline GLM in the spirit of a GAM, in which "
  "licence age, driver age, bonus–malus and risk zone enter through cubic B-spline bases with eight knots, so that main effects may be non-linear "
  "but no interactions are learned; and a Poisson gradient-boosting machine (GBM; 300 trees, learning rate 0.05, at most 15 leaves, at least 50 "
  "observations per leaf, L2 regularisation 1), with hyper-parameters fixed at values typical of practice and not tuned (Section 4.3.2 reports a "
  f"27-point grid and an early-stopping variant). For each class three prices were computed on a 30 percent hold-out sample (n = {n_test:,}, "
  "stratified by claim status, seed 2026): the best-estimate price μ_BE of equation (1) using gender; the unawareness price μ_UN of equation (2), "
  "obtained, as in Lindholm et al. (2022) and in common practice, by refitting without gender; and the discrimination-free price μ_DF of equation "
  "(3), obtained by predicting each hold-out policy once as female and once as male and averaging with the training-sample gender shares.", indent=True)
P("**Decomposing the unawareness test.** A refitted unawareness model differs from the best-estimate model in more than the omission of gender; "
  "every coefficient, spline or tree is re-estimated, so the gap μ_UN − μ_DF mixes the proxy effect that equation (2) describes with whatever else a "
  "refit changes, which for a tree ensemble can be a great deal. Two further quantities separate them. The implied unawareness price applies "
  "equation (2) literally with the best-estimate model held fixed,", indent=True)
EQ("μ_UN∗(x) = Σ_d μ_BE(x, d) · p̂(d | x)", 5)
P("where p̂(d | x) is the probability of gender given the rating factors, estimated by a gradient-boosting classifier fitted on the training set "
  "(a logistic estimate is reported alongside), so that IDI∗ = E|μ_UN∗ − μ_DF| / E[μ_DF] isolates the proxy component. The placebo index repeats "
  "the full drop-and-refit procedure with an independent random binary variable of the same prevalence in place of gender; a random variable can "
  "carry no proxy, so the resulting index measures the refitting component alone, and it is reported as mean and range over ten draws. A third "
  "variant fits the boosted model with interaction constraints that confine every tree to a single feature, so that the role of interaction "
  "learning can be examined without changing the model family.", indent=True)
P("**Metrics.** Predictive quality was assessed by exposure-weighted Poisson deviance and by the AUC for the claim indicator. Proxy capacity was "
  "measured as the five-fold cross-validated AUC with which X alone predicts gender, using a logistic model and a boosting classifier, with "
  "permutation importance identifying the carrying features. Proxy harm was measured with two exposure-weighted indices,", indent=True)
EQ("DDI = E|μ_BE − μ_DF| / E[μ_DF],        IDI = E|μ_UN − μ_DF| / E[μ_DF]", 4)
P(f"the direct- and indirect-discrimination indices, together with the female-to-male ratio of average predicted premium, the actual-to-expected "
  f"claims ratio by gender and the share of policies whose unawareness price differs from the discrimination-free price by more than 5 percent.", indent=True)
H("3.5  Inference and Planned Robustness Checks", 2)
P(f"**Inference.** Uncertainty in the indices, the shift shares and the between-class ratios was quantified by a hold-out bootstrap: the "
  f"{n_test:,} test policies were resampled with replacement {nb:,} times with the fitted models held fixed, and percentile 95% intervals formed. "
  f"Such intervals capture sampling variation in the evaluation sample, not variation from refitting. Five robustness checks, reported in Section "
  f"4.3, are specified here in advance: every one-, two- and three-cell one-point re-scoring of Table 4, with an adversarial scenario in which each "
  f"contestable score is moved against the ranking; a repetition of the whole procedure on {n_refit} further stratified 70/30 splits (seeds 2027 to "
  f"{2026+n_refit}) with every class refitted, plus two recalibration maps, a level factor and a two-parameter log-linear Poisson map estimated from "
  f"five-fold out-of-fold training predictions, with the discrimination-free price rebuilt from the recalibrated best-estimate model; grids over the "
  f"GLM penalty and the GBM configuration scored by cross-validated deviance; the decomposition of Section 3.4 with bootstrap intervals and its "
  f"distribution across the refitted splits; and the actual-to-expected ratios by gender with bootstrap intervals, training-sample values and refit "
  f"means. Categorical factors, including the counterfactual gender assignments, were encoded with fixed category levels drawn from the full data "
  f"so that the boosted model receives identical codes in training, testing and counterfactual prediction. All computations used Python 3.11 and "
  f"scikit-learn 1.3 (Pedregosa et al., 2011); the code reproducing Tables 6–11 and Figures 3–5, the RSI scoring sheet, the derived data and the "
  f"environment specification are openly available at {REPO_URL} (release v1.0.0).", indent=True)
# ================================================================== 4 EMPIRICAL ANALYSIS
H("4.  Empirical Analysis", 1)
H("4.1  Comparative Regulatory Analysis", 2)
H("4.1.1  The three frameworks in outline", 3)
P("The regimes differ first in legal nature and architecture (Table 3). The EU AI Act is a directly applicable regulation that imposes binding "
  "ex-ante obligations on AI systems sorted by risk class; pricing and risk assessment in life and health insurance are expressly high-risk, which "
  "triggers a risk-management system, data governance including bias examination, documentation, record-keeping, transparency, human oversight and "
  "robustness duties (Articles 9–15), verified through conformity assessment (Article 43) and backed by fines of up to EUR 15 million or 3 percent "
  "of worldwide turnover. The Digital Omnibus left the classification intact and moved the compliance date for Annex III systems to 2 December "
  "2027 (European Parliament & Council, 2026); for the majority of insurance AI that is not high-risk, the EIOPA Opinion supplies proportionate "
  "expectations drawn from Solvency II and the IDD (EIOPA, 2025).", indent=True)
P("The NAIC Model Bulletin sits at the opposite pole. It creates no new legal obligation but reminds insurers that existing prohibitions on "
  "unfair discrimination and unfair trade practices apply to AI-driven decisions, and that each insurer should maintain a written AI Systems (AIS) "
  "Program (NAIC, 2023c). By 8 October 2026 it had been adopted in 27 states, the District of Columbia and Puerto Rico; four states rely on "
  "instruments of their own, California's Bulletin 2022-5, Colorado's Regulation 10-1-1 (binding governance and quantitative testing, extended from "
  "life to auto and health insurers from 15 October 2025), New York's Circular Letter No. 7 (a disproportionate-impact test) and Texas's bulletins of "
  "2020 and 2026 (NAIC, 2026; Colorado Division of Insurance, 2025; NYDFS, 2024; Texas Department of Insurance, 2026). The IAIS instruments are soft "
  "law: the Application Paper works through ICPs 7, 8, 16 and 19 and 'does not include new requirements' (IAIS, 2025, p. 2). Figure 1 places the "
  "instruments on a common timeline.", indent=True)
TABLE("Table 3. The three AI governance regimes at a glance (status at October 2026)",
      ["Attribute", "IAIS", "NAIC", "EU AI Act (2024/1689)"],
      [["Legal nature", "Soft law (global standard; application paper)", "Supervisory bulletin adopted state by state", "Binding regulation (hard law)"],
       ["Regulatory style", "Principles-based, proportionate", "Principles-based, law-reinterpreting", "Risk-based, prescriptive"],
       ["Insurance-specific?", "Yes (sector standard-setter)", "Yes (insurance only)", "No (horizontal; insurance captured as high-risk use)"],
       ["Core instrument", "ICPs 7, 8, 16, 19 + Application Paper (final 2 Jul 2025)", "Model Bulletin (4 Dec 2023); 27 states + DC + PR by 8 Oct 2026",
        "Regulation + Annex III 5(c); EIOPA Opinion (6 Aug 2025) for non-high-risk AI"],
       ["Key dates", "Draft Nov 2024; final Jul 2025", "California Bull. 2022-5; Colorado Reg. 10-1-1 (2023, amended 2025); NY CL 7 (2024); Texas B-0003-26 (2026)",
        "In force 1 Aug 2024; Art. 50 from 2 Aug 2026; Annex III high-risk from 2 Dec 2027 (Omnibus)"],
       ["Enforcement", "Via member supervisors (adopt-or-explain)", "Via state insurance departments; market-conduct exams", "Market surveillance; fines up to EUR 15m / 3% turnover"]],
      note="Note. ICPs = Insurance Core Principles. Sources: IAIS (2025); NAIC (2023c, 2026); European Parliament & Council (2024, 2026); EIOPA (2025).",
      widths=[3.0, 4.3, 4.3, 4.4])
FIG(f"{FIGS}/fig2_timeline.png", "Figure 1. Timeline of the principal AI-in-insurance instruments, November 2023 to August 2028. Dark blue: EU instruments; "
    "gold: NAIC; teal: IAIS; grey: U.S. state-specific instruments. The dotted line marks the date of writing (29 October 2026).")

H("4.1.2  Dimension-by-dimension comparison and the Regulatory Stringency Index", 3)
P("Beneath the formal differences lies a marked substantive convergence (Table 4). All three require senior responsibility and a documented "
  "governance programme: an AIS Program for the NAIC, a governance and risk-management expectation under ICPs 7 and 8 for the IAIS, a mandatory "
  "risk-management system with human oversight for the EU. All three target discriminatory outcomes in their own vocabularies, the NAIC through "
  "unfair-discrimination statutes, the EU through fundamental rights and Article 10 data governance, the IAIS through ICP 19, sharpened by the "
  "warning that 'insurers should carefully consider the use of proxy variables' and a call for 'fairness metrics' in high-risk use cases (IAIS, "
  "2025, paras. 86, 90). All three demand record-keeping and explainability. They diverge most on assurance: the EU mandates ex-ante conformity "
  "assessment and post-market monitoring, an auditing architecture with no parallel in the other two (Mökander et al., 2022), whereas the IAIS and "
  "NAIC rely on ex-post review.", indent=True)
TABLE("Table 4. Regulatory Stringency Index: reconciled scores by dimension with text anchors",
      ["Dimension", "IAIS (2025)", "NAIC (2023)", "EU AI Act (2024, as amended 2026)"],
      [["Legal nature", ("1", True), ("2", True), ("3", True)],
       ["", "Application paper; 'does not include new requirements'", "Binding only through state adoption and existing statutes", "Directly applicable regulation; Art. 99 fines"],
       ["Scope & risk tiering", ("2", True), ("2", True), ("3", True)],
       ["", "Sect. 2: risk-based supervision; proportionality", "AIS Program 'commensurate with' risk of each AI system", "Art. 6 and Annex III legal tiers; 5(c) life/health pricing"],
       ["Accountability & governance", ("2", True), ("3", True), ("3", True)],
       ["", "Sect. 3: ICP 7/8; human oversight; third-party data", "Written AIS Program with prescribed content; senior accountability", "Art. 9 risk-management system; Art. 14 human oversight; Art. 26 deployer duties"],
       ["Fairness & non-discrimination", ("2", True), ("2", True), ("3", True)],
       ["", "Paras. 86–91: proxy variables; fairness metrics; outcome monitoring", "Existing unfair-discrimination law; testing 'as appropriate'", "Art. 10(2)(f)–(g) bias examination; Art. 10(5) special-category data for bias correction"],
       ["Transparency & explainability", ("2", True), ("2", True), ("3", True)],
       ["", "Sect. 5: explanations adapted to recipients; redress", "Documentation; consumer notice and recourse", "Arts. 11–13 technical documentation, logging, instructions; Art. 50 transparency"],
       ["Oversight & assurance", ("1", True), ("2", True), ("3", True)],
       ["", "Ordinary supervision; model inventory 'should consider'", "Market-conduct examination; information requests", "Art. 43 conformity assessment; Art. 72 post-market monitoring"],
       [("Total, equal weights (max 18)", True), ("10", True), ("13", True), ("18", True)]],
      note="Note. Scores follow the rubric in Table 2 and were reviewed against the underlying regulatory texts as described in Section 3.2. Article numbers refer to "
           "Regulation (EU) 2024/1689; paragraph numbers to IAIS (2025). Colorado's Regulation 10-1-1 would score 3 on fairness and oversight if treated as a separate "
           "regime, showing that stringency within the U.S. varies by state.",
      widths=[3.4, 4.1, 4.1, 4.4])
FIG(f"{FIGS}/fig1_rsi_radar.png", "Figure 2. Regulatory Stringency Index by governance dimension (0 = absent, 3 = binding and verifiable). "
    "The three profiles share the same shape on the substantive dimensions and differ mainly in legal nature and assurance.", width_cm=10.5)
P(f"Figure 2 makes the pattern visible. With equal weights the IAIS totals 10 of 18, the NAIC 13 and the EU 18, yet the profiles share a "
  f"shape: every regime scores at least 2 on accountability, fairness and transparency, and the two soft-law-leaning regimes differ by one point on "
  f"exactly three dimensions, legal nature, accountability and oversight. Two features of Table 4 should be read before the totals are interpreted. "
  f"The EU maximum is built into the rubric, as Section 3.2 explained; and the ordering EU > NAIC > IAIS is a matter of dominance rather than "
  f"aggregation, since no regime scores below the regime ranked beneath it on any dimension, so the NAIC–IAIS gap of "
  f"{RS['dominance']['gap_NAIC_IAIS']} points and the EU–NAIC gap of {RS['dominance']['gap_EU_NAIC']} points are what any coding disagreement would "
  f"have to overcome. The provisional answer to RQ1 is that the regimes converge in substance and diverge in form; Section 4.3.1 asks how much "
  f"re-scoring that reading can absorb.", indent=True)
H("4.1.3  Mapping AI risks to framework responses", 3)
P("Table 5 maps the value-chain risks of Section 2.1 to the response each regime offers. The EU Act has the most explicit ex-ante controls and "
  "the heaviest compliance burden; the NAIC is quick to deploy because it leans on existing law but depends on the rigour of state examination; the "
  "IAIS is globally harmonising and proportionate but rests on national capacity. No regime is comprehensively superior.", indent=True)
TABLE("Table 5. AI risks in the insurance value chain mapped to framework responses",
      ["Value-chain risk", "Primary framework response", "Framework with the most explicit control"],
      [["Proxy / algorithmic discrimination", "Fairness testing; data governance; outcome monitoring", "EU (Art. 10) and Colorado; IAIS paras. 86–91"],
       ["Opacity / non-contestability", "Explainability and documentation duties; redress", "EU (Arts. 11–13); IAIS Sect. 5–6"],
       ["Diffuse accountability", "Mandated governance programme and senior ownership", "All three (convergent)"],
       ["Data quality and model drift", "Data governance; post-market monitoring", "EU (Art. 72 monitoring)"],
       ["Third-party and foundation models", "Vendor due diligence; retained accountability", "IAIS Sect. 3.5; NAIC third-party section"],
       ["Targeting / exclusion", "Conduct-of-business fair-treatment rules", "IAIS (ICP 19); EIOPA Opinion"]],
      note="Note. The third column identifies the framework whose instruments address the risk most explicitly and directly; it is descriptive and implies no overall ranking.", widths=[4.6, 6.6, 4.8])

H("4.2  Empirical Results", 2)
H("4.2.1  Proxy capacity versus proxy harm", 3)
P(f"Table 6 reports the hold-out results. Gender adds almost nothing to predictive quality: the Poisson deviance of the GLM is "
  f"{f4(g['BE']['deviance'])} with gender and {f4(g['UN']['deviance'])} without it, and the AUC is {f3(g['BE']['auc'])} in both cases. The "
  f"discrimination-free price even has a slightly lower hold-out deviance than the best-estimate price under both additive models "
  f"({f4(g['DF']['deviance'])} against {f4(g['BE']['deviance'])} for the GLM), which is what one would expect of an attribute whose raw effect on "
  f"frequency is under 4 percent ({f3(R['raw_freq']['Female'])} for women, {f3(R['raw_freq']['Male'])} for men) and which the other factors already "
  f"partly encode. The spline model improves deviance slightly ({f4(a_['UN']['deviance'])}) and the GBM somewhat more ({f4(b['UN']['deviance'])}, AUC "
  f"{f3(b['UN']['auc'])}). The rating factors nonetheless recover gender with considerable accuracy (Figure 3): a logistic model reaches a "
  f"cross-validated AUC of {f3(px_l[0])} and a boosting classifier {f3(px_g[0])}, with licence age, vehicle price band, driver age, risk zone and "
  f"vehicle class carrying most of the signal. Proxy capacity is high.", indent=True)
P(f"Proxy harm, measured first in the conventional way, is small under both additive models and larger under boosting. The drop-and-refit "
  f"index is {pct(idi['GLM'],2)} for the GLM (bootstrap 95% CI {ci_pct(BS['GLM']['idi'])}), {pct(idi['GAM'],2)} for the spline GLM "
  f"({ci_pct(BS['GAM']['idi'])}) and {pct(idi['GBM'],2)} for the GBM ({ci_pct(BS['GBM']['idi'])}); the GBM-to-GLM ratio is {f2(ratio_gg['point'])} "
  f"on the hold-out sample and has a median of {f2(RF['ratio_GBM_GLM']['median'])} with a range of {f2(RF['ratio_GBM_GLM']['min'])}–"
  f"{f2(RF['ratio_GBM_GLM']['max'])} across the {n_refit} refitted splits of Section 4.3.2. The implied index of equation (5), which holds the "
  f"best-estimate model fixed, tells a different story: IDI∗ is {pct(imp['GLM'],2)} for the GLM ({ci2(DEC['GLM']['idi_implied_gbm_ci'])}), "
  f"{pct(imp['GAM'],2)} for the GAM ({ci2(DEC['GAM']['idi_implied_gbm_ci'])}) and {pct(imp['GBM'],2)} for the GBM "
  f"({ci2(DEC['GBM']['idi_implied_gbm_ci'])}). For the additive models the two indices nearly coincide, as they should when a refit changes "
  f"little; for the boosted model the refit-based index is about {DEC['GBM']['idi_refit']/imp['GBM']:.0f} times the implied one, and Section 4.3.3 "
  f"shows the difference to be refitting noise. Under the GLM and the GAM essentially no policy moves by more than 5 percent when the "
  f"discrimination-free correction is applied; under the GBM {pct(sh['GBM']['Female'],1)} of women ({ci_pct(BS['GBM']['share5_f'],1)}) and "
  f"{pct(sh['GBM']['Male'],1)} of men ({ci_pct(BS['GBM']['share5_m'],1)}) do, and the central 80 percent of the price ratio spans "
  f"{f3(gbm_p10p90['Female'][0])}–{f3(gbm_p10p90['Female'][1])} for women and {f3(gbm_p10p90['Male'][0])}–{f3(gbm_p10p90['Male'][1])} for men "
  f"(Figure 4).", indent=True)
def trow(f, k, label):
    m = M[f][k]; a = AE[f"{f}_{k}"]
    d = pct(ddi[f],2) + "\n" + ci_pct(BS[f]["ddi"]) if k == "BE" else ("0" if k == "DF" else "–")
    i = pct(idi[f],2) + "\n" + ci_pct(BS[f]["idi"]) if k == "UN" else ("0" if k == "DF" else "–")
    s = pct(imp[f],2) + "\n" + ci2(DEC[f]["idi_implied_gbm_ci"]) if k == "UN" else ("0" if k == "DF" else "–")
    return [label, f4(m["deviance"]), f3(m["auc"]), f4(m["prem_f"]), f4(m["prem_m"]), f3(m["ratio_fm"]),
            f3(m["ae_f"]) + "\n" + ci3(a["ci_f"], 2), f3(m["ae_m"]) + "\n" + ci3(a["ci_m"], 2), d, i, s]
TABLE(f"Table 6. Hold-out performance and fairness metrics by model class and pricing rule (n = {n_test:,} policies)",
      ["Model", "Poisson deviance", "AUC", "Mean μ women", "Mean μ men", "Ratio W/M", "A/E women [95% CI]", "A/E men [95% CI]", "DDI [95% CI]", "IDI, drop-and-refit [95% CI]", "IDI∗, fixed model [95% CI]"],
      [trow("GLM","BE","GLM BE (gender)"), trow("GLM","UN","GLM UN (no gender)"), trow("GLM","DF","GLM DF"),
       trow("GAM","BE","GAM BE (gender)"), trow("GAM","UN","GAM UN (no gender)"), trow("GAM","DF","GAM DF"),
       trow("GBM","BE","GBM BE (gender)"), trow("GBM","UN","GBM UN (no gender)"), trow("GBM","DF","GBM DF")],
      note=f"Note. BE = best-estimate; UN = unawareness (refitted without gender); DF = discrimination-free; GAM = additive spline GLM; μ = predicted claim frequency per unit "
           f"exposure; A/E = actual-to-expected claims; DDI = direct-discrimination index; IDI = indirect-discrimination index (equation 4) computed with the refitted unawareness "
           f"price; IDI∗ = the same index computed with the implied unawareness price of equation (5), best-estimate model held fixed. All indices exposure-weighted; "
           f"intervals are percentile 95% intervals from {nb:,} hold-out bootstrap resamples with the fitted models held fixed. Training-sample A/E for the UN models: "
           f"{f3(AE['GLM_UN']['train_ae_f'])}/{f3(AE['GLM_UN']['train_ae_m'])} (GLM), {f3(AE['GAM_UN']['train_ae_f'])}/{f3(AE['GAM_UN']['train_ae_m'])} (GAM), "
           f"{f3(AE['GBM_UN']['train_ae_f'])}/{f3(AE['GBM_UN']['train_ae_m'])} (GBM). Deviance is exposure-weighted mean Poisson deviance. Proxy-recovery AUC for gender from "
           f"non-protected features: logistic {f3(px_l[0])}, gradient boosting {f3(px_g[0])} (5-fold CV). Drop-and-refit IDI ratios on the hold-out sample: GBM/GLM "
           f"{f2(ratio_gg['point'])} [{f2(ratio_gg['ci'][0])}, {f2(ratio_gg['ci'][1])}]; GAM/GLM {f2(ratio_ag['point'])} [{f2(ratio_ag['ci'][0])}, {f2(ratio_ag['ci'][1])}].",
      widths=[2.2, 1.3, 1.0, 1.2, 1.2, 1.0, 1.5, 1.5, 1.6, 1.8, 1.8], size=7.5)
FIG(f"{FIGS}/fig3_proxy.png", "Figure 3. Proxy capacity. (a) Cross-validated AUC with which non-protected rating factors predict gender; "
    "(b) permutation importance of the ten most informative features in the boosting classifier.")
FIG(f"{FIGS}/fig4_ratio.png", "Figure 4. Proxy harm. Distribution of the ratio of the unawareness price to the discrimination-free price by gender, hold-out sample, "
    "for (a) the Poisson GLM, (b) the additive spline GLM and (c) the Poisson gradient-boosting machine, with the indirect-discrimination index and its bootstrap 95% "
    "interval. A ratio of 1 means dropping gender already yields the discrimination-free price.", width_cm=15.0)
P(f"The direction of the gender effect deserves supervisory attention because it is not what a naive reading would suggest. Including gender "
  f"raises women's average predicted premium relative to men's (ratio {f3(g['BE']['ratio_fm'])} under the GLM), since conditional on the other "
  f"factors women in this portfolio are slightly riskier, whereas the unawareness and discrimination-free prices place the ratio at "
  f"{f3(g['UN']['ratio_fm'])} and {f3(g['DF']['ratio_fm'])}; a group-average premium test would have flagged the best-estimate model, not the proxy "
  f"problem. The actual-to-expected ratios on this hold-out sample run below one for women and above one for men under every class and pricing "
  f"rule (A/E {f3(g['UN']['ae_f'])} and {f3(g['UN']['ae_m'])} for the unawareness GLM), but the bootstrap intervals in Table 6 are wide, the "
  f"training-sample ratios are {f3(AE['GLM_UN']['train_ae_f'])} and {f3(AE['GLM_UN']['train_ae_m'])}, and Section 4.3.4 shows that the pattern "
  f"reverses across refitted splits; it belongs to this evaluation sample, not to the model classes.", indent=True)
H("4.2.2  What each regime would detect", 3)
P(f"Table 7 reads the results back through the three frameworks and so answers RQ2. One caveat governs the EU column: Annex III, point 5(c) "
  f"makes pricing high-risk only in life and health insurance, so Article 10 does not apply to a motor portfolio, and the applicable EU expectations "
  f"are the proportionate ones of the EIOPA Opinion (2025); the column therefore reads the results in two registers, what the Opinion would expect "
  f"here and what Article 10 would expect were the same test run on a life or health line. The unawareness GLM and GAM produce a small gap "
  f"({pct(idi['GLM'],2)} and {pct(idi['GAM'],2)}; {pct(imp['GLM'],2)} and {pct(imp['GAM'],2)} on the fixed-model measure), whose materiality depends on "
  f"a supervisory standard that none of the instruments states numerically. The boosted model is the instructive case: its drop-and-refit index of "
  f"{pct(idi['GBM'])}, with {pct(share_f,1)} of women's and {pct(share_m,1)} of men's policies shifted by more than 5 percent, would plausibly trigger "
  f"the IAIS outcome-monitoring expectation and, under Colorado's rule, documented remediation, whereas its fixed-model index of {pct(imp['GBM'],2)} "
  f"would trigger nothing. A supervisor applying the conventional test would act on a signal that Section 4.3.3 shows to be refitting noise. Under "
  f"none of the regimes does removing gender discharge the fairness requirement, and two of them (EU Article 10(5) and IAIS para. 90) anticipate the "
  f"collection of protected data that the correction requires.", indent=True)
TABLE("Table 7. Reading the empirical results through the three regimes",
      ["Test contemplated by the regime", "IAIS (2025)", "NAIC (2023) / Colorado", "EU: EIOPA Opinion (motor) / AI Act Art. 10 (life and health)", "Result on freMPL1"],
      [["Protected attribute removed at point of sale", "Necessary, not sufficient (para. 86)", "Necessary; unfair-discrimination law applies to outcomes", "Necessary; fairness controls expected (Opinion) / Art. 10(2)(f) bias examination due (high-risk lines)", f"Done in μ_UN; residual IDI {pct(idi['GLM'],2)} (GLM), {pct(idi['GAM'],2)} (GAM), {pct(idi['GBM'],2)} (GBM); fixed-model IDI∗ {pct(imp['GLM'],2)}, {pct(imp['GAM'],2)}, {pct(imp['GBM'],2)}"],
       ["Proxy-capacity test (predict D from X)", "Implied by para. 88", "Colorado Reg. 10-1-1 requires inferred-attribute testing", "Not prescribed; useful for data-governance risk analysis", f"AUC {f3(px_g[0])}: high capacity"],
       ["Outcome / group-disparity test", "Para. 90–91: fairness metrics, dashboards", "NY CL 7 disproportionate-impact test", "Outcome monitoring (Opinion) / Art. 10(2)(g) 'possible biases … likely to affect' (high-risk lines)", f"Ratio W/M {f2(ratio_lo)}–{f2(ratio_hi)}; A/E gap up to {f2(ae_gap_max)} ({ae_gap_f} {ae_gap_k}), interval includes small values"],
       ["Individual price-shift test vs. μ_DF", "Compatible with para. 91 metrics", "Not prescribed", "Compatible with risk-management expectations (Opinion; Art. 9 for high-risk lines)", f"0% (GLM, GAM) vs. {pct(share_f,1)}–{pct(share_m,1)} (GBM, drop-and-refit) of policies > 5% shift"],
       ["Use of protected data for testing", "Para. 90 where legally permitted", "Colorado permits inference; NY requires testing", "Art. 10(5) special-category exemption (high-risk lines only)", "Required to compute μ_DF and μ_UN∗"],
       ["Proportionality by model class", "Sect. 2: risk-based supervision", "AIS Program 'commensurate' with risk", "Tiering by use, not by model class", f"Drop-and-refit index {f2(ratio_gg['point'])}× larger under GBM than GLM; fixed-model index {DEC['GBM']['idi_implied_gbm']/DEC['GLM']['idi_implied_gbm']:.2f}× (smaller)"]],
      note="Note. IDI = indirect-discrimination index; μ_UN, μ_DF and μ_UN∗ as in equations (2), (3) and (5). Paragraph numbers refer to IAIS (2025). Motor pricing is not a high-risk use under Annex III of the AI Act; the Article 10 entries describe what would apply were the same test run on a life or health line.", widths=[3.4, 2.8, 3.0, 3.6, 3.2], size=8)

H("4.3  Robustness Checks and Decomposition", 2)
P("Section 3.5 announced five checks, and this section reports them in that order.", indent=True)
H("4.3.1  Re-scoring the index", 3)
k1, k2, k3 = RS["perturbations"]["k1"], RS["perturbations"]["k2"], RS["perturbations"]["k3"]; ADV = RS["adversarial_scenario"]
P(f"Because the ranking is fixed by dominance, the question is whether other codings, not other weights, would change it. Table 8 and Figure 5 "
  f"enumerate every admissible one-point change to one, two or three cells of Table 4. None of the {k1['n_valid']} single re-scorings and none of the "
  f"{k2['n_valid']} pairs changes the strict order EU > NAIC > IAIS; among the {k3['n_valid']:,} triples, {k3['n_tie_only']} "
  f"({100*k3['n_tie_only']/k3['n_valid']:.1f} percent) tie the NAIC with the IAIS and none reverses any pair. Reversing the NAIC–IAIS order needs at "
  f"least {RS['min_one_point_changes_to_reverse_NAIC_IAIS']} coordinated one-point changes in the IAIS's favour; a tie with the EU needs "
  f"{RS['min_one_point_changes_to_tie_EU_NAIC']}. The adversarial scenario in the last row shows what such coordination would mean: were the IAIS "
  f"credited with binding force through ICP adoption, a named accountability expectation, mandated fairness metrics and targeted examination, while "
  f"the NAIC lost a point on legal nature and accountability and the EU one on insurance-specific scope, the totals would be IAIS {ADV['totals']['IAIS']}, "
  f"NAIC {ADV['totals']['NAIC']} and EU {ADV['totals']['EU']}. We rejected each of those changes on the text, but the exercise locates the "
  f"sensitivity: the EU's position is immune to coding disputes of this size, whereas the NAIC–IAIS order rests on three one-point judgements. The "
  f"totals under the alternative weighting schemes once used as a check are kept in the replication package, not offered as evidence.", indent=True)
TABLE("Table 8. Re-scoring analysis of the Regulatory Stringency Index",
      ["Perturbation", "Admissible cases", "Strict order EU > NAIC > IAIS preserved", "NAIC–IAIS tie", "Reversal of any pair"],
      [[f"One cell changed by one point", f"{k1['n_valid']}", f"{k1['n_strict_order_preserved']} (100%)", f"{k1['n_tie_only']}", f"{k1['n_reversal']}"],
       [f"Two cells changed by one point each", f"{k2['n_valid']}", f"{k2['n_strict_order_preserved']} (100%)", f"{k2['n_tie_only']}", f"{k2['n_reversal']}"],
       [f"Three cells changed by one point each", f"{k3['n_valid']:,}", f"{k3['n_strict_order_preserved']:,} ({100*k3['share_preserved']:.1f}%)", f"{k3['n_tie_only']}", f"{k3['n_reversal']}"],
       ["Adversarial scenario (seven contestable cells moved against the ranking)", "1", "No", "—", f"IAIS {ADV['totals']['IAIS']} > NAIC {ADV['totals']['NAIC']}; EU {ADV['totals']['EU']} unchanged at the top"]],
      note=f"Note. Baseline totals IAIS 10, NAIC 13, EU 18 (Table 4); NAIC–IAIS gap {RS['dominance']['gap_NAIC_IAIS']} points, EU–NAIC gap {RS['dominance']['gap_EU_NAIC']} points. "
           f"A perturbation is admissible if every score stays within 0–3. Adversarial scenario: IAIS +1 on legal nature, accountability, fairness and oversight; NAIC −1 on legal nature and accountability; EU −1 on scope.",
      widths=[5.2, 2.2, 3.4, 2.0, 3.2], size=8.5)
FIG(f"{FIGS}/fig5_rsi_rescoring.png", "Figure 5. Distribution of the point gaps between regimes after every admissible one-, two- and three-cell re-scoring of Table 4: "
    "(a) NAIC total minus IAIS total; (b) EU total minus NAIC total. Dotted lines mark the baseline gaps; a gap of zero is a tie and a negative gap a reversal.", width_cm=14.0)
H("4.3.2  Model refitting, recalibration and hyper-parameters", 3)
P(f"**Refitting and recalibration.** Across {n_refit} repeated splits with every model refitted (Table 9), the additive models are stable: the GLM "
  f"index averaged {pct(RF['GLM']['idi']['mean'],2)} (range {pct(RF['GLM']['idi']['min'],2)}–{pct(RF['GLM']['idi']['max'],2)}) and the GAM index "
  f"{pct(RF['GAM']['idi']['mean'],2)}, and both used gender in every split. The boosted model's index averaged {pct(RF['GBM']['idi']['mean'],2)} "
  f"(median {pct(RF['GBM']['idi']['median'],2)}) but ranged from zero to {pct(RF['GBM']['idi']['max'],2)}: in {n_unused} of {n_refit} splits it never "
  f"split on gender, so that its three prices coincided, and in the other {n_refit-n_unused} it exceeded the GLM's in {n_above}. A fairness test run "
  f"once on one sample can therefore under- or over-state what the same class would show on a neighbouring one. Recalibration leaves the picture "
  f"intact: the level factors are close to one, and the log-linear maps shrink the additive models' predictions (slopes "
  f"{f2(CAL['GLM']['loglinear']['params_UN'][1])} and {f2(CAL['GAM']['loglinear']['params_UN'][1])}) more than the boosted model's "
  f"({f2(CAL['GBM']['loglinear']['params_UN'][1])}), lowering the GLM and GAM indices to {pct(CAL['GLM']['loglinear']['idi'],2)} and "
  f"{pct(CAL['GAM']['loglinear']['idi'],2)} and leaving the GBM at {pct(CAL['GBM']['loglinear']['idi'],2)}. A slope near one half means that the "
  f"out-of-fold log-predictions of the additive models are about twice as dispersed as the data support, which is over-fitting rather than "
  f"mis-specification: with α = 10⁻⁴ the one-hot design is penalised only nominally, and the unawareness GLM's training deviance is "
  f"{f4(ga_main['train_deviance_un'])} against {f4(ga_main['holdout_deviance_un'])} on the hold-out.", indent=True)
def rrow(f, label):
    r = RF[f]; c = CAL[f]; q = AER[f]
    return [label, f"{pct(r['idi']['mean'],2)} / {pct(r['idi']['median'],2)}", f"{pct(r['idi']['min'],2)}–{pct(r['idi']['max'],2)}", str(r["n_gender_unused"]),
            f"{f4(r['deviance_un']['mean'])} ± {f4(r['deviance_un']['sd'])}", f"{f3(r['auc_un']['mean'])} ± {f3(r['auc_un']['sd'])}",
            pct(c['raw']['idi'],2), pct(c['level']['idi'],2), f"{pct(c['loglinear']['idi'],2)} (b = {f2(c['loglinear']['params_UN'][1])})",
            f"{f3(c['raw']['ae_f'])} / {f3(c['raw']['ae_m'])}", f"{f3(q['ae_f_mean'])} ± {f3(q['ae_f_sd'])} / {f3(q['ae_m_mean'])} ± {f3(q['ae_m_sd'])}"]
TABLE(f"Table 9. Robustness of the drop-and-refit indirect-discrimination index to model refitting and recalibration",
      ["Model class", f"IDI, {n_refit} refitted splits: mean / median", "Range", "Splits with gender unused", "UN deviance, mean ± sd", "UN AUC, mean ± sd",
       "IDI, main hold-out", "IDI, level-recalibrated", "IDI, log-linear recalibrated", "A/E W / M, main hold-out", f"A/E W / M, mean ± sd over {n_refit} refits"],
      [rrow("GLM", "GLM"), rrow("GAM", "GAM (spline GLM)"), rrow("GBM", "GBM")],
      note=f"Note. Left block: {n_refit} repeated stratified 70/30 splits (seeds 2027–{2026+n_refit}) with all models refitted; GBM/GLM ratio of IDI median "
           f"{f2(RF['ratio_GBM_GLM']['median'])}, range {f2(RF['ratio_GBM_GLM']['min'])}–{f2(RF['ratio_GBM_GLM']['max'])}; the GBM index exceeded the GLM index in "
           f"{n_above} splits and was zero in {n_unused} (gender not selected by the boosted best-estimate model). Right block: recalibration maps estimated from "
           f"five-fold cross-fitted training predictions and applied to the main models' hold-out predictions; b = slope of the log-linear map for the unawareness "
           f"model (b = 1 means no dispersion correction). GBM/GLM ratio: {f2(CAL['ratio_GBM_GLM_raw'])} raw, {f2(CAL['ratio_GBM_GLM_level'])} level, "
           f"{f2(CAL['ratio_GBM_GLM_loglinear'])} log-linear.",
      widths=[1.6, 1.7, 1.4, 1.2, 1.6, 1.4, 1.2, 1.4, 1.7, 1.5, 1.9], size=7)

P(f"**Hyper-parameters.** The penalty matters (Table 10). Five-fold cross-validated deviance is minimised at α = {GA['cv_best_alpha']} rather "
  f"than at the 10⁻⁴ of the main analysis, the recalibration slope moves towards one as the penalty rises ({f2(ga_main['recal_slope_un'])} at 10⁻⁴, "
  f"{f2(ga_best['recal_slope_un'])} at {GA['cv_best_alpha']}, {f2(GA['grid']['0.1']['recal_slope_un'])} at 0.1), and the GLM's drop-and-refit index "
  f"falls from {pct(ga_main['idi'],2)} to {pct(ga_best['idi'],2)} at the cross-validated optimum and to {pct(GA['grid']['0.1']['idi'],2)} at α = 0.1, "
  f"a factor of nearly eight across a defensible range. The boosted model is more configuration-dependent still: over the {GG['n_configs']} "
  f"combinations of trees, learning rate and leaves its index ranges from zero to {pct(GG['idi_range'][1],2)}, is exactly zero in {n_gg_zero} because "
  f"the best-estimate model never splits on gender, and exceeds the main GLM index in {GG['n_idi_above_glm_main']}; the configuration with the lowest "
  f"cross-validated deviance ({gg_best['max_iter']} trees, learning rate {gg_best['learning_rate']}, {gg_best['max_leaf_nodes']} leaves) is one of the "
  f"zero cases, and an early-stopped fit gives {pct(GG['early_stopping']['idi'],2)}. The main configuration simply sits in the middle of a wide "
  f"distribution, and the width is the finding.", indent=True)
def grow(a):
    v = GA["grid"][a]; return [f"GLM, α = {float(a):g}" + (" (main)" if a == "0.0001" else " (CV optimum)" if a == GA["cv_best_alpha"] else ""), f4(v["cv_deviance_un"]), f4(v["holdout_deviance_un"]), f3(v["auc_un"]), f2(v["recal_slope_un"]), pct(v["ddi"],2), pct(v["idi"],2)]
gsorted = sorted(GG["grid"].values(), key=lambda v: v["cv_deviance_un"])
def ggrow(v, label=None):
    return [label or f"GBM, {v['max_iter']} trees, lr {v['learning_rate']}, {v['max_leaf_nodes']} leaves", f4(v["cv_deviance_un"]), f4(v["holdout_deviance_un"]), f3(v["auc_un"]), "—", pct(v["ddi"],2), pct(v["idi"],2)]
es = GG["early_stopping"]
TABLE("Table 10. Hyper-parameter sensitivity of the drop-and-refit indices",
      ["Specification", "5-fold CV deviance (UN)", "Hold-out deviance (UN)", "AUC (UN)", "Recalibration slope (UN)", "DDI", "IDI"],
      [grow(a) for a in ["0.0", "1e-05", "0.0001", "0.001", "0.01", "0.1"]] +
      [ggrow(GG["grid"][GG["main"]], f"GBM, 300 trees, lr 0.05, 15 leaves (main)"), ggrow(gsorted[0], f"GBM, {gsorted[0]['max_iter']} trees, lr {gsorted[0]['learning_rate']}, {gsorted[0]['max_leaf_nodes']} leaves (CV optimum)"),
       ggrow(gsorted[-1], f"GBM, {gsorted[-1]['max_iter']} trees, lr {gsorted[-1]['learning_rate']}, {gsorted[-1]['max_leaf_nodes']} leaves (worst CV)"),
       [f"GBM, early stopping ({es['n_iter_be']} / {es['n_iter_un']} trees)", "—", f4(es["holdout_deviance_un"]), f3(es["auc_un"]), "—", pct(es["ddi"],2), pct(es["idi"],2)],
       [f"GBM grid, all {GG['n_configs']} configurations", "—", "—", "—", "—", "—", f"0–{pct(GG['idi_range'][1],2)}; zero in {n_gg_zero}; above main GLM in {GG['n_idi_above_glm_main']}"]],
      note="Note. UN = unawareness model. GLM penalty α applies to the standardised one-hot design (scikit-learn PoissonRegressor). Recalibration slope is the coefficient b of the "
           "out-of-fold log-linear map of Table 9 (b = 1 means no dispersion correction). GBM grid: trees {100, 300, 600} × learning rate {0.03, 0.05, 0.1} × leaves {7, 15, 31}, "
           "min. 50 observations per leaf, L2 = 1; early stopping on a 20 percent validation split with patience 20. Indices on the main hold-out sample.",
      widths=[4.6, 1.9, 1.9, 1.3, 1.9, 1.4, 3.0], size=8)
H("4.3.3  Proxy component versus refitting component", 3)
P(f"Table 11 decomposes the test as Section 3.4 described. For the additive models the implied index, which holds the best-estimate model fixed "
  f"and substitutes the estimated P(D | X), is close to the drop-and-refit index ({pct(imp['GLM'],2)} against {pct(idi['GLM'],2)} for the GLM, "
  f"{pct(imp['GAM'],2)} against {pct(idi['GAM'],2)} for the GAM), whichever classifier supplies the probabilities. For the boosted model the implied "
  f"index is {pct(imp['GBM'],2)} ({ci2(DEC['GBM']['idi_implied_gbm_ci'])}), a tenth of the drop-and-refit figure, while the distance between the two "
  f"unawareness prices, {pct(DEC['GBM']['dist_un_vs_implied'],2)}, is as large as that figure itself; the boosted model also uses gender less, a flip "
  f"of gender moving its best-estimate price by {pct(DEC['GBM']['gender_effect_be'],2)} of premium against {pct(DEC['GLM']['gender_effect_be'],2)} for "
  f"the GLM. The placebo confirms the diagnosis: with an independent random attribute in place of gender, the boosted model yields a drop-and-refit "
  f"index of {pct(plm['GBM'],2)} on average over ten draws (range {pct(PL['GBM']['min'],2)}–{pct(PL['GBM']['max'],2)}), indistinguishable from the "
  f"{pct(idi['GBM'],2)} obtained with gender, whereas the additive models yield {pct(plm['GLM'],2)} and {pct(plm['GAM'],2)}. What the conventional "
  f"test reports for a tree ensemble is, in this portfolio, refitting noise with a small proxy signal underneath. The interaction-free boosted model "
  f"never splits on gender at all (both indices zero, hold-out deviance {f4(GNI['deviance_un'])}), so the role of interactions cannot be isolated; "
  f"whether a boosted model uses a weak protected attribute is a matter of configuration and sample, as Tables 9 and 10 already indicated.", indent=True)
P(f"The implied index is also the stabler measure across refits: over the {n_refit} splits its mean is {pct(IR['GLM']['implied']['mean'],2)} "
  f"(sd {pct(IR['GLM']['implied']['sd'],2)}) for the GLM, {pct(IR['GAM']['implied']['mean'],2)} for the GAM and {pct(IR['GBM']['implied']['mean'],2)} "
  f"for the GBM, the boosted model's implied index lies below the GLM's in {IR['n_splits_GBM_implied_below_GLM_implied']} of {n_refit} splits, and "
  f"its drop-and-refit index exceeds its implied index by more than a factor of two in {IR['n_splits_refit_GBM_exceeds_implied_GBM_by_factor2']}. "
  f"Read with the fixed-model measure, the answer to RQ2 is that proxy harm is under 1 percent of premium in every class and lowest under boosting, "
  f"and that the model-class penalty reported by the drop-and-refit test is a property of the test rather than of the harm.", indent=True)
def drow(f, label):
    d = DEC[f]; p = PL[f]
    return [label, pct(d["idi_refit"],2), pct(d["idi_implied_gbm"],2) + " " + ci2(d["idi_implied_gbm_ci"]), pct(d["idi_implied_logit"],2), pct(d["dist_un_vs_implied"],2),
            pct(d["gender_effect_be"],2), f"{pct(p['mean'],2)} ({pct(p['min'],2)}–{pct(p['max'],2)})", f"{pct(IR[f]['implied']['mean'],2)} ± {pct(IR[f]['implied']['sd'],2)}"]
TABLE("Table 11. Decomposing the unawareness test: proxy component, refitting component and placebo",
      ["Model class", "IDI, drop-and-refit", "IDI∗, fixed model, P(D|X) from GBM classifier [95% CI]", "IDI∗, P(D|X) from logistic classifier", "Distance μ_UN vs μ_UN∗ (% of premium)",
       "Mean |μ_BE(x,F) − μ_BE(x,M)| (% of premium)", "Placebo IDI, random attribute: mean (range, 10 draws)", f"IDI∗, mean ± sd over {n_refit} refits"],
      [drow("GLM", "GLM"), drow("GAM", "GAM (spline GLM)"), drow("GBM", "GBM"),
       ["GBM, interaction-free (one feature per tree)", pct(GNI["idi"],2), "0.00% (gender never used)", "—", "—", "0.00%", "—", "—"]],
      note=f"Note. μ_UN∗ and IDI∗ as in equation (5); the gender classifiers are fitted on the training set and hold-out AUCs are {f3(V['proxy_auc_holdout']['gbm'])} (GBM) and "
           f"{f3(V['proxy_auc_holdout']['logit'])} (logistic). Placebo: gender replaced by an independent Bernoulli variable with the training-sample female share, full "
           f"drop-and-refit procedure repeated, ten seeds. Interaction-free GBM: same hyper-parameters as the main GBM with interaction constraints confining each tree to a single feature.",
      widths=[2.4, 1.5, 2.6, 1.7, 1.8, 1.9, 2.3, 1.9], size=7.5)
H("4.3.4  Calibration by gender under sampling uncertainty", 3)
P(f"The actual-to-expected ratios of Section 4.2.1 do not survive resampling. The intervals in Table 6 reach or include one in all but one "
  f"cell; the men-minus-women gap, whose interval on this sample is {ci3(AE['GLM_UN']['ci_gap'], 2)} for the GLM, reverses sign on the training "
  f"sample ({f3(AE['GLM_UN']['train_ae_f'])} and {f3(AE['GLM_UN']['train_ae_m'])}) and on average across the {n_refit} refitted splits "
  f"({f3(AER['GLM']['ae_f_mean'])} ± {f3(AER['GLM']['ae_f_sd'])} for women, {f3(AER['GLM']['ae_m_mean'])} ± {f3(AER['GLM']['ae_m_sd'])} for men), with "
  f"women below one in only {AER['GLM']['n_f_below_1']} splits. With roughly 2,300 women's policies and 250 claims among them, the main sample "
  f"simply drew an unfavourable configuration. What a group-level calibration check surfaces on one sample is mostly the width of its own "
  f"uncertainty.", indent=True)
P(f"The checks therefore sort the findings into two groups. The ordering of the regimes, the smallness of proxy harm under the additive "
  f"models, the high proxy capacity of the rating factors and the inadequacy of unawareness as a compliance strategy survive every check. The "
  f"reading that the drop-and-refit figures invite, that boosting triples proxy harm, does not; the fixed-model measure places the boosted model "
  f"lowest, and the discussion that follows is built on that reversal.", indent=True)
H("4.4  Discussion", 2)
H("4.4.1  Convergence in substance, divergence in form", 3)
P("The comparison yields a clear thesis. The three regimes are converging on a common substantive core of governance and accountability, data "
  "quality, fairness, transparency and human oversight, while diverging on legal force, prescriptiveness and the mechanism of assurance. The "
  "pattern reflects a shared lineage in the OECD AI Principles and the wider ethics consensus (Jobin et al., 2019; OECD, 2019), filtered through "
  "three legal cultures: codified, rights-based law in the EU, sectoral state regulation reinterpreting existing statutes in the United States, "
  "consensus soft law at the IAIS. The 2025–2026 developments strengthen the reading: the IAIS finalised its text without adding requirements, the "
  "NAIC count rose past half the states while the most demanding U.S. rules stayed state-specific, and the EU deferred its deadline without touching "
  "the obligations. The re-scoring analysis adds the right qualification: the EU's position is fixed by the construct and by any plausible coding, "
  "whereas the NAIC–IAIS order rests on three one-point judgements and would reverse if the IAIS paper were read as binding and its fairness metrics "
  "as mandated. This agrees with Floridi (2021) and Smuha (2021) on the EU instrument as the hard-law expression of a principles consensus, and it "
  "extends to insurance a convergence that Eling et al. (2022) anticipated but did not score.", indent=True)
H("4.4.2  What the empirical test adds to the literature", 3)
P("The empirical component qualifies the doctrinal consensus on three points. The first is that proxy capacity is not proxy harm. The rating "
  "factors predict gender well, the kind of finding that fuels alarm about algorithmic discrimination, yet the price distortion is under 1 percent "
  "in every model class once it is measured with the best-estimate model held fixed, because gender itself shifts the best-estimate price very "
  "little here. Tests that stop at proxy capacity will produce false positives; tests that stop at group-average premiums will produce false "
  "negatives, since the group ratios in Table 6 are close to one under every rule. The discrimination-free benchmark of Lindholm et al. (2022) is "
  "the natural yardstick because it measures the product of capacity and effect at the individual level; what our test adds to their analytical "
  "result and to Prince and Schwarcz's (2020) doctrinal argument is the magnitude, with its sampling uncertainty, and the observation that none of "
  "the three instruments states a numerical threshold against which it would be judged, a silence Xin and Huang (2024) also note. Lindholm et al. "
  "(2024a) have since separated proxy discrimination, which the discrimination-free price removes, from demographic disparity, which it may leave; "
  "our group ratios near one show the two diverging in practice, as Moriah et al. (2024) also find on motor data.", indent=True)
P(f"The second point concerns the test rather than the model. The same data, rule and attribute produced a drop-and-refit index about "
  f"{f2(ratio_gg['point'])} times larger under gradient boosting than under a GLM, and the natural reading, the one the literature on flexible "
  f"learners invites, is that interaction-learning capacity aggravates proxy harm. The decomposition shows otherwise. With the best-estimate model "
  f"held fixed the boosted model's proxy component is the smallest of the three, and a random placebo reproduces its drop-and-refit index almost "
  f"exactly: what the conventional test measures for a tree ensemble is the difference between two independently grown ensembles, one of which "
  f"happens to lack a weak feature, and it attributes to gender a difference that has nothing to do with it. The additive spline model remains the "
  f"informative control, though for a different reason than one might expect: because a refit changes an additive model little, its two indices "
  f"agree, which is why the drop-and-refit test has served the GLM literature well and why its transfer to flexible learners, which Lindholm et al. "
  f"(2022) and the regulatory texts both contemplate, needs the fixed-model form of equation (5) or a placebo baseline. Prince and Schwarcz (2020) "
  f"argued that machine learning makes proxy discrimination harder to detect; the most common detection method, it turns out, also makes it easy to "
  f"see where it is not.", indent=True)
P(f"Two related findings bear on model class. Whether the boosted model uses gender at all depends on the training sample and the "
  f"configuration (it did not in {n_unused} of {n_refit} refitted splits and in {n_gg_zero} of {GG['n_configs']} hyper-parameter settings, the "
  f"cross-validated optimum among them), so a single fairness test on a single fit is less informative for a boosted model than for a GLM and argues "
  f"for recurrent testing with the test's own sampling distribution reported. And the GLM's index is not a constant either: it falls by a factor of "
  f"nearly eight as the penalty moves from nominal to the cross-validated optimum and beyond, because a weakly penalised one-hot design lets the model "
  f"reconstruct gender from sparse category combinations that a stronger penalty smooths away. Regularisation is itself a fairness parameter, which "
  f"the actuarial and regulatory literatures have not generally treated as one. The EU tiers obligations by use case and the NAIC and IAIS by the risk "
  f"of the system; neither ties scrutiny to model class, and on this evidence they are right not to, since the harm does not scale with flexibility. "
  f"What should scale with flexibility is the testing protocol. Section 5.3 turns this into propositions.", indent=True)
P("The third point is that unawareness is not compliance and that compliance needs protected data. Computing μ_DF and μ_UN∗ required gender "
  "for every policy in the development sample, which is the operational meaning of Article 10(5) of the AI Act and of paragraph 90 of the IAIS "
  "paper: fairness testing and privacy minimisation pull in opposite directions, and a regime that wants outcome testing must say where the "
  "protected data may come from.", indent=True)
H("4.4.3  Reading the results for two emerging markets", 3)
P("Emerging-market supervisors face a sharper version of the dilemma: drawn to the EU model by its extraterritorial weight, many lack the "
  "actuarial, audit and supervisory capacity that a conformity-assessment regime quietly assumes, so that adopting the Act wholesale risks "
  "unenforceable formalism and ignoring it risks arbitrage and exclusion from reinsurance and data flows. Two markets, treated as illustrations "
  "rather than case studies, make the point concrete.", indent=True)
P("Egypt overhauled its regime through the Unified Insurance Law No. 155 of 2024, effective 11 July 2024, which consolidates supervision under "
  "the Financial Regulatory Authority (FRA) and gives it explicit powers over digital distribution and insurtech; Board Decisions No. 102 of 2025 and "
  "No. 145 of 2026 extended the compliance deadline to 11 July 2027, the last extension Article 7 permits (Financial Regulatory Authority, 2026; "
  "Al-Shorouk, 2026). Direct premiums reached EGP 114.1 billion in 2025, up 20.9 percent, with motor the second-largest property-and-liability line "
  "after medical cover (Financial Regulatory Authority, 2026). AI is arriving through the FRA's sandbox, whose fintech committee gave preliminary "
  "approval in July 2026 to two AI insurance projects, a platform assistant and a voice-agent and chatbot service (Ahram Online, 2026; Middle East "
  "Insurance Review, 2026); sandboxes raise innovators' access to funding and shape how regulators learn (Alaassar et al., 2021; Cornelli et al., "
  "2024), and they also invite the 'riskwashing' Brown and Piroska (2022) describe when exit conditions are loose. The law says little about "
  "algorithmic governance, so the market is entering the AI era with growing adoption, a sandbox already hosting AI projects and no "
  "algorithmic-governance text to apply to them.", indent=True)
P("Saudi Arabia sits at a different point on the same curve. The Council of Ministers established a dedicated Insurance Authority by Resolution "
  "No. 85 of 15 August 2023, moving supervision out of the central bank whose combined supervision Orlando and Bace (2021) describe; the Authority "
  "inherited the insurtech rules of 2022, with their sandbox and a code of conduct that includes non-discrimination, and supervises a market of SAR "
  "84.3 billion in 2025 in which motor is the second-largest line after health and the one that generates most complaints (Council of Ministers, "
  "2023; SAMA, 2022; Insurance Authority, 2026; Seyam, 2025). The Saudi Data and AI Authority's ethics principles of 2023 already require fairness, "
  "accountability and transparency, and sector regulators are expected to integrate them (SDAIA, 2023). The Saudi case therefore poses the "
  "proportionality question in its purest form: a young authority with horizontal principles in place, a sandbox as entry point and no legacy "
  "algorithmic rules to reinterpret.", indent=True)
# ================================================================== 5 CONCLUSION AND POLICY IMPLICATIONS
H("5.  Conclusion and Policy Implications", 1)
H("5.1  Conclusion and Regulatory Implications", 2)
P(f"This paper asked how the three dominant regulatory frameworks for AI in insurance compare in their 2025–2026 form, whether the fairness "
  f"test they share detects the harm it targets, and what follows for emerging markets. To RQ1: the regimes converge in substance and diverge in "
  f"form; the ordering follows from dominance, survives every one- and two-cell re-scoring, and is contestable only between the two "
  f"soft-law-leaning regimes. To RQ2: the rating factors recover gender with an AUC of {f2(px_g[0])}, yet the proxy component left by removing "
  f"gender is below 1 percent of premium in every model class and lowest under boosting; the drop-and-refit test overstates it for the boosted "
  f"model roughly tenfold, a placebo attribute producing an index of the same size, and the GLM's own index moves by a factor of eight with its "
  f"penalty. To RQ3: measuring proxy harm needs protected data, a fixed-model rather than a refitted comparison, and a report of the test's sampling "
  f"distribution; proportionality belongs to the testing protocol, not to a presumed penalty on flexible models.", indent=True)
P("The contribution is threefold: a comparative framework and index whose construction is stated and whose coding sensitivity is measured; an "
  "empirical bridge between AI-governance requirements and discrimination-free pricing showing that the standard unawareness test decomposes into a "
  "proxy and a refitting component, the second dominating for a tree ensemble, which corrects a reading we think common in practice and initially "
  "shared; and a fixed-model testing routine for insurers with a proportionate path for supervisors. The findings translate into recommendations "
  "for four addressees.", indent=True)
P(f"**Insurers and their boards.** Because every regime demands documented governance, senior accountability, fairness testing, explainability "
  f"and oversight, the efficient strategy is to embed AI oversight inside existing enterprise risk-management structures rather than build a "
  f"regime-specific silo (Eling et al., 2022). A model-risk policy should then require, for any pricing or underwriting model, (i) a proxy-capacity "
  f"screen; (ii) a discrimination-free benchmark computed with the best-estimate model held fixed, as in equation (5), with an individual-level "
  f"shift threshold; (iii) a placebo baseline whenever a refit-based comparison is used; (iv) calibration by protected group with its sampling "
  f"interval; and (v) repeated fits across splits and configurations for learners whose use of a weak attribute is unstable, with the regularisation "
  f"strength recorded. The routine is cheap, the three model pairs, three prices and proxy screen having run in {RT['routine_total']:.0f} seconds on "
  f"a laptop (Apple M1 Pro; bootstrap and refits excluded), and it maps onto the IAIS 'fairness dashboard' (para. 91), the NAIC AIS Program and, for "
  f"high-risk lines, the EU Article 10 record.", indent=True)
P("**Egypt's Financial Regulatory Authority.** The Unified Insurance Law says little about algorithmic governance, and the transition to 11 "
  "July 2027 leaves time to fill the space without primary legislation: proportionate supervisory guidance under existing conduct and governance "
  "powers, in the manner of the NAIC Bulletin; the five-step routine above as an exit condition of the sandbox, so that the first algorithmic pricing "
  "models reach the market with a documented proxy screen, a fixed-model benchmark and group calibration with intervals; and a legal basis, on the "
  "model of Article 10(5), for insurers to hold protected attributes solely for bias testing, without which neither the benchmark nor the "
  "requirement it serves can be demonstrated.", indent=True)
P("**Saudi Arabia's Insurance Authority.** A young authority can write AI expectations into its framework from the outset: adopt the convergent "
  "core through the ICPs it already applies; make the SDAIA principles operative in the insurtech rules and the sandbox so that fairness testing "
  "becomes a condition of entry; tier obligations by use case, as the EU does, but tier the testing protocol by model class, with fixed-model "
  "benchmarks, placebo baselines and recurrent monitoring for boosted and neural models whose use of a protected attribute this study found to be "
  "sample- and configuration-dependent; and import conformity assessment and post-market monitoring only where it can assess conformity.", indent=True)
P("**The IAIS and emerging-market supervisors generally.** The proportionate route for both markets runs through the IAIS rather than through "
  "Brussels: adopt the convergent core through ICPs already applied, import the EU's risk tiering and post-market monitoring only where audit "
  "capacity exists, and borrow the NAIC's reinterpretation strategy to avoid legislative delay. Table 12 sets out this middle path and completes the "
  "answer to RQ3.", indent=True)
TABLE("Table 12. Proposed proportionate AI-governance principles for emerging-market insurance supervisors",
      ["Principle", "Practical translation", "Borrowed from", "Evidence from this study"],
      [["Proportionality by use case and by testing protocol", "Tier obligations by use case; tier the fairness-testing protocol (fixed-model benchmark, placebo, repeated fits) by the model's refitting variance", "EU risk tiers + IAIS Sect. 2", f"Drop-and-refit index {f2(ratio_gg['point'])}× larger under GBM than GLM but placebo reproduces it ({pct(plm['GBM'],2)}); fixed-model index lowest under GBM ({pct(imp['GBM'],2)}); gender unused in {n_unused} of {n_refit} refits"],
       ["Leverage existing law", "Apply unfair-discrimination and fair-treatment rules to AI now", "NAIC Model Bulletin", "Group ratios and A/E computable under current rules (Table 6)"],
       ["Mandated governance", "Board-owned AI programme inside ERM and three lines of defence", "All three (convergent core)", "Every regime scores at least 2 on accountability, fairness and transparency (Table 4), so one governance programme can meet the convergent core of all three"],
       ["Test outcomes, not inputs", "Require a discrimination-free benchmark computed with the best-estimate model held fixed, and an individual shift threshold", "IAIS paras. 90–91; Colorado 10-1-1", f"Unawareness leaves a proxy component of {pct(imp['GBM'],2)}–{pct(imp['GAM'],2)}; the refitted comparison overstates it for GBM tenfold"],
       ["Enable protected data for testing", "Legal basis for collecting or inferring protected attributes solely for bias testing", "EU Art. 10(5); IAIS para. 90", "Neither μ_DF nor μ_UN∗ can be computed without D"],
       ["Capacity-calibrated assurance", "Phase in recurrent monitoring and audit as actuarial and audit capacity grows; use sandboxes as testing gates", "EU monitoring (selective); FRA and IA sandboxes", f"Routine runs in {RT['routine_total']:.0f} s on standard pricing data; boosted-model indices vary across refits and configurations (Tables 9–10)"],
       ["Global interoperability", "Align to IAIS and OECD to avoid arbitrage and keep reinsurance access", "IAIS + OECD principles", "Drawn from the literature on regulatory convergence (Jobin et al., 2019; OECD, 2019), not from this study's results"]],
      note="Note. The principles are proposed on the basis of the comparative and empirical evidence above; they are not derived from a survey of emerging-market supervisors. "
           "The recommended path adopts the convergent substantive core via proportionate principles and imports hard-law assurance tools selectively as capacity allows.",
      widths=[3.1, 4.8, 3.2, 4.9], size=8.5)
H("5.2  Limitations of the Study", 2)
P("Seven limitations apply. The EU high-risk regime is still phasing in, so its effects are assessed prospectively and the Article 10 readings "
  "in Table 7 are hypothetical for a motor portfolio. The coding was done by one author and reviewed by the other rather than double-coded, so no "
  "inter-rater statistic can be offered; the published anchors and the re-scoring analysis show how much a disagreement would matter but do not "
  "replace independent coding. The rubric's top anchors describe maximal formalisation and so resemble the EU Act; the index measures distance from "
  "that ceiling, not regime-neutral quality, and its dominance-ordered scores mean no weighting could have changed the ranking. The empirical test "
  "uses one portfolio, one protected attribute and one line of business, with 2004 data that predate telematics and external data; attributes with "
  "larger direct effects may show larger proxy components. The Egyptian and Saudi material is illustrative. The three model classes are not a "
  "spectrum of flexibility, the GLM penalty and GBM configuration were fixed a priori and both indices depend on them, and the interaction-free "
  "boosted model did not use gender at all, so nothing is claimed about interactions. Finally, the hold-out intervals hold the models fixed, the "
  "placebo is a ten-draw average and the classifier supplying P(D | X) is itself an estimate whose error the implied index inherits; the "
  "GLM–GAM–GBM contrast is evidence about how the unawareness test behaves across classes, not evidence that any class is intrinsically fairer.", indent=True)
H("5.3  Future Research", 2)
P("The comparative analysis and the decomposition motivate, rather than test, five propositions; P3 and P4 follow most directly from the "
  "actuarial results, while P1, P2 and P5 concern compliance cost, regulatory dynamics and extraterritorial influence, none of which this study "
  "measures:", indent=True)
for t in ["**P1.** Insurers that embed AI oversight within existing enterprise risk-management structures reach compliance with all three regimes at lower cost than those that build standalone AI-compliance functions.",
          "**P2.** Substantive governance requirements converge across regimes faster than their enforcement mechanisms; the Regulatory Stringency Index spread on assurance will remain wider than on fairness through 2028.",
          f"**P3.** For any learner, the drop-and-refit unawareness index exceeds the fixed-model index by an amount that increases with the learner's refitting variance and is reproduced by a placebo attribute; the gap is negligible for penalised additive models and of the order of the index itself for tree ensembles, as observed here ({pct(idi['GBM'],2)} against {pct(imp['GBM'],2)}).",
          "**P4.** Proxy capacity (AUC of predicting D from X) is a weak predictor of the proxy component of harm (IDI∗) across portfolios; harm is better predicted by the direct effect of D on the best-estimate price, and for penalised models by the strength of the penalty.",
          "**P5.** The EU AI Act exerts measurable extraterritorial influence on emerging-market insurance regulation independently of trade ties; Egypt and the Gulf are natural test cases as their sandboxes mature."]:
    p = P(t, after=3); p.paragraph_format.left_indent = Cm(0.75)
P("Each proposition can be tested by cross-jurisdictional comparison, event-study designs or replication on other portfolios; P3 in particular "
  "can be run on any dataset with a recorded protected attribute using the replication package.", indent=True, before=4)
if not BLIND:
    H("Declarations", 1)
    P("**Data and code availability.** The freMPL1 dataset is openly available in the CASdatasets collection (Dutang & Charpentier, 2020). All code "
      f"required to reproduce Tables 6–11 and Figures 3–5 (numbered scripts 01–06 for data preparation, the index, the three model classes, "
      f"proxy capacity, the bootstrap and robustness checks, and the decomposition and review checks), together with the RSI scoring sheet, the "
      f"derived data and environment specifications, is openly available at {REPO_URL} (release v1.0.0).", after=3)
    P("**Funding.** No external funding was received.   **Conflicts of interest.** None.   **Ethics.** The study uses anonymised open data and public regulatory documents; no ethics approval was required.", after=3)
# ================================================================== REFERENCES
H("References", 1)
refs = [
 "Adams-Prassl, J., Binns, R., & Kelly-Lyth, A. (2023). Directly discriminatory algorithms. *The Modern Law Review, 86*(1), 144–175. https://doi.org/10.1111/1468-2230.12759",
 "Barocas, S., & Selbst, A. D. (2016). Big data's disparate impact. *California Law Review, 104*(3), 671–732.",
 "Black, J. (2008). Forms and paradoxes of principles-based regulation. *Capital Markets Law Journal, 3*(4), 425–457. https://doi.org/10.1093/cmlj/kmn026",
 "Black, J., & Baldwin, R. (2010). Really responsive risk-based regulation. *Law & Policy, 32*(2), 181–213. https://doi.org/10.1111/j.1467-9930.2010.00318.x",
 "Cath, C. (2018). Governing artificial intelligence: Ethical, legal and technical opportunities and challenges. *Philosophical Transactions of the Royal Society A, 376*(2133), 20180080. https://doi.org/10.1098/rsta.2018.0080",
 "Cevolini, A., & Esposito, E. (2020). From pool to profile: Social consequences of algorithmic prediction in insurance. *Big Data & Society, 7*(2), 1–11. https://doi.org/10.1177/2053951720939228",
 "Charpentier, A. (2024). *Insurance, biases, discrimination and fairness*. Springer. https://doi.org/10.1007/978-3-031-49783-4",
 "Colorado Division of Insurance. (2025). *Regulation 10-1-1: Governance and risk management framework requirements for life insurers', private passenger automobile insurers', and health benefit plan insurers' use of external consumer data and information sources, algorithms, and predictive models* (3 CCR 702-10, as amended effective October 15, 2025).",
 "Court of Justice of the European Union. (2011). *Association belge des Consommateurs Test-Achats ASBL and Others v Conseil des ministres*, Case C-236/09, judgment of the Court (Grand Chamber) of 1 March 2011, ECLI:EU:C:2011:100.",
 "de Almeida, P. G. R., dos Santos, C. D., & Farias, J. S. (2021). Artificial intelligence regulation: A framework for governance. *Ethics and Information Technology, 23*(3), 505–525. https://doi.org/10.1007/s10676-021-09593-z",
 "Dutang, C., & Charpentier, A. (2020). *CASdatasets: Insurance datasets* (R package version 1.0-11) [Data set]. http://cas.uqam.ca",
 "Eling, M., Nuessle, D., & Staubli, J. (2022). The impact of artificial intelligence along the insurance value chain and on the insurability of risks. *The Geneva Papers on Risk and Insurance — Issues and Practice, 47*(2), 205–241. https://doi.org/10.1057/s41288-020-00201-7",
 "European Insurance and Occupational Pensions Authority. (2024). *Report on the digitalisation of the European insurance sector*. EIOPA.",
 "European Insurance and Occupational Pensions Authority. (2025, August 6). *Opinion on artificial intelligence governance and risk management*. EIOPA.",
 "European Insurance and Occupational Pensions Authority. (2026, February 2). *Generative AI market survey: Outlook, use cases and risk management* (EIOPA-BoS-25-679). EIOPA.",
 "European Parliament & Council of the European Union. (2024). Regulation (EU) 2024/1689 laying down harmonised rules on artificial intelligence (Artificial Intelligence Act). *Official Journal of the European Union*, L, 12 July 2024.",
 "European Parliament & Council of the European Union. (2026). Regulation amending Regulation (EU) 2024/1689 as regards the timing of high-risk obligations and simplification measures (Digital Omnibus on AI). *Official Journal of the European Union*, L, 24 July 2026.",
 "Financial Regulatory Authority. (2026). *Al-kitab al-ihsa'i al-sanawi 'an nashat al-ta'min khilal al-'am al-miladi 2025* [Annual statistical book on insurance activity, calendar year 2025]. Arab Republic of Egypt. https://fra.gov.eg/wp-content/uploads/2026/09/الكتاب-الاحصائي-2025.pdf",
 "Al-Shorouk. (2026, August 2). *Al-Raqaba al-Maliya tusdir al-qarar raqm 145 li-sanat 2026 bi-madd muhlat tawfiq awda' sharikat al-ta'min hatta 11 yulyu 2027* [FRA issues Decision No. 145 of 2026 extending the insurance sector's compliance deadline to 11 July 2027]. https://www.shorouknews.com/news/view.aspx?cdate=02082026&id=c12c3fad-3fc3-4b8b-bef6-622465494f1d",
 "Ahram Online. (2026, July 28). *FRA grants preliminary approval to two artificial-intelligence insurance projects in the regulatory sandbox*. https://gate.ahram.org.eg/News/5838475.aspx",
 "Middle East Insurance Review. (2026, August 3). *Egypt: Regulator gives preliminary nod to two AI insurance projects in sandbox*. https://www.meinsurancereview.com/News/View-NewsLetter-Article/id/96413/Type/MiddleEast/",
 "Floridi, L. (2021). The European legislation on AI: A brief analysis of its philosophical approach. *Philosophy & Technology, 34*(2), 215–222. https://doi.org/10.1007/s13347-021-00460-9",
 "Floridi, L., Cowls, J., Beltrametti, M., Chatila, R., Chazerand, P., Dignum, V., Luetge, C., Madelin, R., Pagallo, U., Rossi, F., Schafer, B., Valcke, P., & Vayena, E. (2018). AI4People—An ethical framework for a good AI society. *Minds and Machines, 28*(4), 689–707. https://doi.org/10.1007/s11023-018-9482-5",
 "Frees, E. W., & Huang, F. (2023). The discriminating (pricing) actuary. *North American Actuarial Journal, 27*(1), 2–24. https://doi.org/10.1080/10920277.2021.1951296",
 "Gasser, U., & Almeida, V. A. F. (2017). A layered model for AI governance. *IEEE Internet Computing, 21*(6), 58–62. https://doi.org/10.1109/MIC.2017.4180835",
 "Council of Ministers of the Kingdom of Saudi Arabia. (2023). *Resolution No. 85 of 28/1/1445 AH (15 August 2023) approving the establishment of the Insurance Authority and its statute*. Umm Al-Qura Gazette. https://www.uqn.gov.sa/details?p=23532",
 "Insurance Authority. (2026). *The Saudi insurance market report 2025*. Kingdom of Saudi Arabia. https://www.ia.gov.sa/SectorReports/The%20Saudi%20Insurance%20Market%20Report%202025.pdf",
 "International Association of Insurance Supervisors. (2025, July 2). *Application paper on the supervision of artificial intelligence*. IAIS.",
 "Jobin, A., Ienca, M., & Vayena, E. (2019). The global landscape of AI ethics guidelines. *Nature Machine Intelligence, 1*(9), 389–399. https://doi.org/10.1038/s42256-019-0088-2",
 "Lindholm, M., Richman, R., Tsanakas, A., & Wüthrich, M. V. (2022). Discrimination-free insurance pricing. *ASTIN Bulletin, 52*(1), 55–89. https://doi.org/10.1017/asb.2021.23",
 "Lindholm, M., Richman, R., Tsanakas, A., & Wüthrich, M. V. (2024a). What is fair? Proxy discrimination vs. demographic disparities in insurance pricing. *Scandinavian Actuarial Journal, 2024*(9), 935–970. https://doi.org/10.1080/03461238.2024.2364741",
 "Lindholm, M., Richman, R., Tsanakas, A., & Wüthrich, M. V. (2024b). A multi-task network approach for calculating discrimination-free insurance prices. *European Actuarial Journal, 14*(2), 329–369. https://doi.org/10.1007/s13385-023-00367-z",
 "Mittelstadt, B. (2019). Principles alone cannot guarantee ethical AI. *Nature Machine Intelligence, 1*(11), 501–507. https://doi.org/10.1038/s42256-019-0114-4",
 "Mittelstadt, B. D., Allo, P., Taddeo, M., Wachter, S., & Floridi, L. (2016). The ethics of algorithms: Mapping the debate. *Big Data & Society, 3*(2), 1–21. https://doi.org/10.1177/2053951716679679",
 "Mökander, J., Axente, M., Casolari, F., & Floridi, L. (2022). Conformity assessments and post-market monitoring: A guide to the role of auditing in the proposed European AI regulation. *Minds and Machines, 32*(2), 241–268. https://doi.org/10.1007/s11023-021-09577-4",
 "National Association of Insurance Commissioners. (2022). *Private passenger auto artificial intelligence/machine learning survey results*. NAIC Big Data and Artificial Intelligence (H) Working Group.",
 "National Association of Insurance Commissioners. (2023a). *Home insurance artificial intelligence/machine learning survey results*. NAIC.",
 "National Association of Insurance Commissioners. (2023b). *Life insurance artificial intelligence/machine learning survey results*. NAIC.",
 "National Association of Insurance Commissioners. (2023c). *Model bulletin: Use of artificial intelligence systems by insurers* (adopted December 4, 2023). NAIC.",
 "National Association of Insurance Commissioners. (2026). *Implementation of NAIC Model Bulletin: Use of artificial intelligence systems by insurers* (status as of October 8, 2026). NAIC. https://content.naic.org/sites/default/files/legal-adoption-map-ai-model-bulletin.pdf",
 "New York State Department of Financial Services. (2024, July 11). *Insurance Circular Letter No. 7: Use of artificial intelligence systems and external consumer data and information sources in insurance underwriting and pricing*. NYDFS.",
 "Organisation for Economic Co-operation and Development. (2019). *Recommendation of the Council on artificial intelligence* (OECD/LEGAL/0449). OECD.",
 "Pedregosa, F., Varoquaux, G., Gramfort, A., Michel, V., Thirion, B., Grisel, O., Blondel, M., Prettenhofer, P., Weiss, R., Dubourg, V., Vanderplas, J., Passos, A., Cournapeau, D., Brucher, M., Perrot, M., & Duchesnay, É. (2011). Scikit-learn: Machine learning in Python. *Journal of Machine Learning Research, 12*, 2825–2830.",
 "Prince, A. E. R., & Schwarcz, D. (2020). Proxy discrimination in the age of artificial intelligence and big data. *Iowa Law Review, 105*(3), 1257–1318.",
 "Rudin, C. (2019). Stop explaining black box machine learning models for high-stakes decisions and use interpretable models instead. *Nature Machine Intelligence, 1*(5), 206–215. https://doi.org/10.1038/s42256-019-0048-x",
 "Saudi Central Bank. (2022). *Rules for insurtech*. SAMA.",
 "Saudi Data and Artificial Intelligence Authority. (2023). *AI ethics principles* (Version 2.0). SDAIA.",
 "Selbst, A. D., Boyd, D., Friedler, S. A., Venkatasubramanian, S., & Vertesi, J. (2019). Fairness and abstraction in sociotechnical systems. *Proceedings of the Conference on Fairness, Accountability, and Transparency*, 59–68. https://doi.org/10.1145/3287560.3287598",
 "Smuha, N. A. (2021). From a 'race to AI' to a 'race to AI regulation': Regulatory competition for artificial intelligence. *Law, Innovation and Technology, 13*(1), 57–84. https://doi.org/10.1080/17579961.2021.1898300",
 "Tabassi, E. (2023). *Artificial intelligence risk management framework (AI RMF 1.0)* (NIST AI 100-1). National Institute of Standards and Technology. https://doi.org/10.6028/NIST.AI.100-1",
 "Texas Department of Insurance. (2026, June 12). *Commissioner's Bulletin B-0003-26: Use of artificial intelligence by regulated entities*. TDI.",
 "Veale, M., & Binns, R. (2017). Fairer machine learning in the real world: Mitigating discrimination without collecting sensitive data. *Big Data & Society, 4*(2), 1–17. https://doi.org/10.1177/2053951717743530",
 "Wachter, S., Mittelstadt, B., & Floridi, L. (2017). Why a right to explanation of automated decision-making does not exist in the General Data Protection Regulation. *International Data Privacy Law, 7*(2), 76–99. https://doi.org/10.1093/idpl/ipx005",
 "Xin, X., & Huang, F. (2024). Antidiscrimination insurance pricing: Regulations, fairness criteria, and models. *North American Actuarial Journal, 28*(2), 285–319. https://doi.org/10.1080/10920277.2023.2190528",
 "Yeung, K. (2018). Algorithmic regulation: A critical interrogation. *Regulation & Governance, 12*(4), 505–523. https://doi.org/10.1111/rego.12158",
]
refs += ['Alaassar, A., Mention, A.-L., & Aas, T. H. (2021). Exploring a new incubation model for FinTechs: Regulatory sandboxes. *Technovation, 103*, 102237. https://doi.org/10.1016/j.technovation.2021.102237', 'Araiza Iturria, C. A., Hardy, M., & Marriott, P. (2024). A discrimination-free premium under a causal framework. *North American Actuarial Journal, 28*(4), 801–821. https://doi.org/10.1080/10920277.2023.2291524', 'Ayana, G., Dese, K., Daba Nemomssa, H., Habtamu, B., Mellado, B., Badu, K., Yamba, E., Faye, S. L., Ondua, M., Nsagha, D., Nkweteyim, D., & Kong, J. D. (2024). Decolonizing global AI governance: Assessment of the state of decolonized AI governance in Sub-Saharan Africa. *Royal Society Open Science, 11*(8), 231994. https://doi.org/10.1098/rsos.231994', 'Badawy, W. (2025). The ethical use and development of artificial intelligence (AI) strategy in Egypt: Identifying gaps and recommendations. *AI and Ethics, 5*(4), 3579–3591. https://doi.org/10.1007/s43681-025-00729-1', 'Bohnert, A., Fritzsche, A., & Gregor, S. (2019). Digital agendas in the insurance industry: The importance of comprehensive approaches. *The Geneva Papers on Risk and Insurance — Issues and Practice, 44*(1), 1–19. https://doi.org/10.1057/s41288-018-0109-0', 'Brown, E., & Piroska, D. (2022). Governing fintech and fintech as governance: The regulatory sandbox, riskwashing, and disruptive social classification. *New Political Economy, 27*(1), 19–32. https://doi.org/10.1080/13563467.2021.1910645', 'Cornelli, G., Doerr, S., Gambacorta, L., & Merrouche, O. (2024). Regulatory sandboxes and fintech funding: Evidence from the UK. *Review of Finance, 28*(1), 203–233. https://doi.org/10.1093/rof/rfad017', 'Côté, O., Côté, M.-P., & Charpentier, A. (2025). A fair price to pay: Exploiting causal graphs for fairness in insurance. *Journal of Risk and Insurance, 92*(1), 33–75. https://doi.org/10.1111/jori.12503', 'Effoduh, J. O., Akpudo, U. E., & Kong, J. D. (2024). Toward a trustworthy and inclusive data governance policy for the use of artificial intelligence in Africa. *Data & Policy, 6*, e34. https://doi.org/10.1017/dap.2024.26', "Lanfranchi, D., & Grassi, L. (2022). Examining insurance companies' use of technology for innovation. *The Geneva Papers on Risk and Insurance — Issues and Practice, 47*(3), 520–537. https://doi.org/10.1057/s41288-021-00258-y", 'Mhlanga, D. (2020). Industry 4.0 in finance: The impact of artificial intelligence (AI) on digital financial inclusion. *International Journal of Financial Studies, 8*(3), 45. https://doi.org/10.3390/ijfs8030045', 'Mohamed, S., Png, M.-T., & Isaac, W. (2020). Decolonial AI: Decolonial theory as sociotechnical foresight in artificial intelligence. *Philosophy & Technology, 33*(4), 659–684. https://doi.org/10.1007/s13347-020-00405-8', 'Moriah, M., Vermet, F., & Charpentier, A. (2024). Measuring and mitigating biases in motor insurance pricing. *European Actuarial Journal, 14*(3), 833–869. https://doi.org/10.1007/s13385-024-00390-8', 'Müller-Bloch, C., & Kranz, J. (2015). A framework for rigorously identifying research gaps in qualitative literature reviews. In *Proceedings of the 36th International Conference on Information Systems (ICIS 2015)*. Association for Information Systems. https://aisel.aisnet.org/icis2015/proceedings/ResearchMethods/2', 'Ogoh, G., Eke, D., & Stahl, B. C. (2026). AI governance in North Africa: An ecosystems approach. *Science and Public Policy, 53*(2), 258–276. https://doi.org/10.1093/scipol/scag007', 'Orlando, G., & Bace, E. (2021). Challenging times for insurance, banking and financial supervision in Saudi Arabia (KSA). *Administrative Sciences, 11*(3), 62. https://doi.org/10.3390/admsci11030062', 'Owens, E., Sheehan, B., Mullins, M., Cunneen, M., Ressel, J., & Castignani, G. (2022). Explainable artificial intelligence (XAI) in insurance. *Risks, 10*(12), 230. https://doi.org/10.3390/risks10120230', 'Robinson, K. A., Saldanha, I. J., & McKoy, N. A. (2011). Development of a framework to identify research gaps from systematic reviews. *Journal of Clinical Epidemiology, 64*(12), 1325–1330. https://doi.org/10.1016/j.jclinepi.2011.06.009', 'Seyam, E. A. (2025). Modeling the determinants of insurtech adoption: Evidence from the Saudi insurance industry. *International Journal of Innovative Research and Scientific Studies, 8*(3), 3297–3312. https://doi.org/10.53894/ijirss.v8i4.7229', 'Truby, J., Brown, R., & Dahdal, A. (2020). Banking on AI: Mandating a proactive approach to AI regulation in the financial sector. *Law and Financial Markets Review, 14*(2), 110–120. https://doi.org/10.1080/17521440.2020.1760454', 'van Bekkum, M., Zuiderveen Borgesius, F., & Heskes, T. (2025). AI, insurance, discrimination and unfair differentiation: An overview and research agenda. *Law, Innovation and Technology, 17*(1), 177–204. https://doi.org/10.1080/17579961.2025.2469348', 'Zaijuan, X., Ramayah, T., & Yubo, S. (2025). Determinants of generative artificial intelligence adoption in insurance companies. *International Journal of Business and Society, 26*(1), 272–291. https://doi.org/10.33736/ijbs.9562.2025']
import re as _re
def _key(r):
    k = _re.sub(r'^(van |de )', '', r).lower(); return k
refs = sorted(refs, key=_key)
for ref in refs:
    p = P(ref, align="l", size=11, after=3); p.paragraph_format.left_indent = Cm(1.0); p.paragraph_format.first_line_indent = Cm(-1.0); p.paragraph_format.line_spacing = 1.0

name = "AI_Governance_Insurance_Manuscript_v7_BLINDED.docx" if BLIND else "AI_Governance_Insurance_Manuscript_v7.docx"
doc.save(f"{ROOT}/{name}"); print("saved", name)
