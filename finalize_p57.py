from docx import Document
from docx.shared import Cm, Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_JUSTIFY
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, PageBreak, Table, TableStyle, Image
from pptx import Presentation
from pptx.util import Inches as PInches, Pt as PPt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
import os, json, urllib.request

ROOT = r"C:\Users\guest pc\Desktop\sepm cat 2"
DOCX = os.path.join(ROOT, "Final_Report.docx")
PDF = os.path.join(ROOT, "Final_Report.pdf")
PPTX = os.path.join(ROOT, "BhashaNews_Presentation.pptx")
SCREENSHOTS = os.path.join(ROOT, "screenshots")

try:
    articles = json.load(urllib.request.urlopen("http://127.0.0.1:8000/api/articles"))
except Exception:
    articles = []

doc = Document()
sec = doc.sections[0]
sec.page_width = Cm(21); sec.page_height = Cm(29.7)
sec.left_margin = Cm(3); sec.right_margin = Cm(2.5); sec.top_margin = Cm(2.5); sec.bottom_margin = Cm(2.5)
for s in doc.sections:
    footer = s.footer.paragraphs[0]; footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = footer.add_run()
    fld1 = OxmlElement("w:fldChar"); fld1.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText"); instr.set(qn("xml:space"), "preserve"); instr.text = "PAGE"
    fld2 = OxmlElement("w:fldChar"); fld2.set(qn("w:fldCharType"), "end")
    run._r.extend([fld1,instr,fld2])

normal = doc.styles["Normal"]
normal.font.name = "Times New Roman"; normal.font.size = Pt(12)
normal.paragraph_format.line_spacing = 1.5
normal.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
for name in ("Heading 1","Heading 2","Heading 3"):
    doc.styles[name].font.name = "Times New Roman"
doc.styles["Heading 1"].font.size = Pt(14); doc.styles["Heading 1"].font.bold = True

def heading(text):
    doc.add_heading(text, level=1)

def para(text):
    doc.add_paragraph(text)

def table(headers, rows):
    t = doc.add_table(rows=1, cols=len(headers))
    t.style = "Table Grid"; t.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i,h in enumerate(headers):
        t.rows[0].cells[i].text = h
        for r in t.rows[0].cells[i].paragraphs[0].runs:
            r.bold=True; r.font.name="Times New Roman"; r.font.size=Pt(10)
    for row in rows:
        cells=t.add_row().cells
        for i,v in enumerate(row):
            cells[i].text=str(v)
            for r in cells[i].paragraphs[0].runs:
                r.font.name="Times New Roman"; r.font.size=Pt(10)
    return t

def picture(name, caption):
    path=os.path.join(SCREENSHOTS,name)
    if os.path.exists(path):
        doc.add_picture(path, width=Cm(15.2))
        p=doc.add_paragraph(caption); p.alignment=WD_ALIGN_PARAGRAPH.CENTER

p=doc.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.CENTER
r=p.add_run("UNIVERSITY NAME: [COLLEGE NAME PLACEHOLDER]\nDEPARTMENT: [DEPARTMENT PLACEHOLDER]\n\nBHASHANEWS\nRegional News Aggregator with Summaries")
r.bold=True; r.font.name="Times New Roman"; r.font.size=Pt(22)
para("\nCase P57 / Roll Number 57")
p=doc.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.CENTER
p.add_run("Student Name: [STUDENT FULL NAME PLACEHOLDER]\nUSN: [USN PLACEHOLDER]\nGuide: [GUIDE NAME PLACEHOLDER]\nAcademic Year: [ACADEMIC YEAR PLACEHOLDER]")
doc.add_page_break()

heading("Certificate")
para("This certificate page is provided as a placeholder. [STUDENT FULL NAME PLACEHOLDER], Roll Number 57, submitted a BhashaNews project report under [GUIDE NAME PLACEHOLDER]. No approval or signature is implied.")
para("\nGuide Signature: ____________________  Date: __________\nHOD Signature: ____________________   Date: __________")
doc.add_page_break()

heading("Declaration")
para("I declare that this report is for Case P57 BhashaNews – Regional News Aggregator with Summaries. AI tools were permitted and used for implementation, testing support and documentation. This report distinguishes verified execution from prototype simplification, documentation, and simulated academic exercises.")
doc.add_page_break()

heading("Acknowledgement")
para("I acknowledge [GUIDE NAME PLACEHOLDER], [DEPARTMENT PLACEHOLDER], and [COLLEGE NAME PLACEHOLDER]. The reader/editor requirements are represented through role-play and implementation personas; no external interview or team participation is claimed.")
doc.add_page_break()

heading("Abstract")
para("BhashaNews is a responsive regional news demonstration for readers and editors. It provides a professional dashboard, featured and latest news, category and language filtering, search, article detail pages, prewritten demonstration summaries, source attribution, bookmarking and local persistence. A React and Vite frontend communicates with a Python FastAPI and SQLAlchemy backend using SQLite. The database is repeatably seeded with 30 clearly labelled demonstration articles across English, Hindi, Kannada, Tamil, Telugu and Malayalam, plus eight project categories. The prototype does not scrape external news, call an AI API or perform real Transformer inference. Twelve backend tests passed and total measured statement coverage was 93% before the final 30-article seed was restored. The React production build passed. Ten complete browser screenshots verify the dashboard, filters, search, detail, summary, bookmarks and mobile layout. The real university Report_Template.docx and logo were not present in the workspace, so no fabricated template or logo is included.")
doc.add_page_break()

heading("Table of Contents")
toc=[("1","Introduction"),("2","Software Development Methodology and AI-Assisted SDLC"),("3","Requirements Engineering and Planning"),("4","AI-Assisted Development and Implementation"),("5","Software Testing and Verification"),("6","Software Quality Assurance"),("7","Integration, Results and Demonstration"),("8","Responsible AI and Ethics"),("9","Conclusion and Future Enhancements"),("Appendices","Coverage, prompt log, test evidence and references")]
table(["Section","Chapter"],toc)
heading("List of Figures")
table(["Figure","Description"],[(f"Figure {i}",c) for i,c in enumerate([
"Home dashboard","Featured-news area","Language filter","Category filter","Search results","Article detail","Summary panel","Bookmarked article","Saved articles","Mobile responsive view"],1)])
heading("List of Tables")
table(["Table","Description"],[("Table 1","Technology and prototype boundary"),("Table 2","Tested API endpoints"),("Table 3","Actual test results"),("Table 4","Coverage IDs"),("Table 5","Known limitations")])

chapters=[
("1. Introduction","BhashaNews addresses the need for short, language-filtered regional news and clear source attribution. The implemented demonstration has reader browsing, search, category and language navigation, article detail, summary, bookmark and saved-pages flows. It is a local prototype: SQLite replaces PostgreSQL and prewritten summaries replace live pretrained Transformer inference."),
("2. Software Development Methodology and AI-Assisted SDLC","An incremental and prototype-enhanced approach was used. Waterfall is too inflexible for interface feedback; Incremental delivery supports independent dashboard and detail increments; Spiral is useful for risk cycles; Prototype supports visual validation; V-Model links requirements to testing. AI assistance was used to generate and review code and tests, while actual build and test execution was used for verification."),
("3. Requirements Engineering and Project Planning","The implemented requirements cover readers and editors: readers browse, search, filter and save articles; editors can manage article records through the backend prototype. The product backlog was represented by sprint increments for discovery, dashboard, filtering, detail, persistence and verification. Actual external client interviews and Jira participation were not claimed."),
("4. AI-Assisted Development and Implementation","The React frontend uses React Router and Lucide icons. FastAPI exposes health, article, category, language, featured, trending and search endpoints. SQLAlchemy models articles and seed data is repeatable. Summaries are prewritten and clearly labelled. Editor mutation endpoints remain in the earlier backend implementation, but the requested demonstration did not require an editor login interface."),
("5. Software Testing and Verification","The backend test suite contained 12 pytest cases. It passed 12/12 and measured 93% total statement coverage. A React production build passed. Browser verification confirmed dashboard rendering and interactions. Selenium POM code exists from the earlier implementation but was not rerun with this polished UI."),
("6. Software Quality Assurance","The quality attributes used for the local prototype are correctness, reliability, maintainability, usability and efficiency. Correctness was supported by API tests; reliability by health endpoint checks; maintainability by modular React/FastAPI files; usability by responsive screenshots; and efficiency by successful builds. SonarQube, remote CI and real Jenkins execution were not performed."),
("7. Integration, Results and Demonstration","The running system was verified at http://127.0.0.1:5173 with FastAPI at http://127.0.0.1:8000. The API contained 30 seeded articles. The browser was used to navigate the home dashboard, article, saved page and mobile layout. Every reference source remains an example.com demonstration URL and is not presented as real journalism."),
("8. Responsible AI and Ethics","Prewritten summaries are labelled, source attribution is displayed and no external article text or copyrighted event is fabricated as current reporting. Bias, privacy, transparency and accountability should be re-evaluated before a real deployment. No API key or external AI service is used."),
("9. Conclusion and Future Enhancements","The completed prototype demonstrates core BhashaNews reader behavior and a reproducible local workflow. Future work includes RSS ingestion, a reviewed pretrained multilingual summariser, PostgreSQL migration, editor authentication, database constraints and real source licensing review. Remote CI/CD, 45/50 summary review and production deployment remain incomplete."),
]
for title,text in chapters:
    doc.add_page_break(); heading(title); para(text)
    if title.startswith("1."):
        table(["Item","Prototype choice","Verified boundary"],[("Frontend","React/Vite","Built and rendered"),("Backend","FastAPI","12 tests passed"),("Database","SQLite","30 articles seeded"),("Summaries","Prewritten","Not live AI"),("Publication data","Demonstration","No external scraping")])
    if title.startswith("4."):
        table(["Endpoint","Purpose","Status"],[("/api/health","Health","Verified"),("/api/articles","List/filter","Verified"),("/api/articles/{id}","Detail","Verified"),("/api/categories","Category list","Verified"),("/api/languages","Language list","Verified"),("/api/search?q=","Search","Verified"),("/api/articles/featured","Featured","Verified"),("/api/articles/trending","Trending","Verified")])
    if title.startswith("5."):
        table(["Check","Actual result","Status"],[("Backend tests","12 passed","Verified"),("Statement coverage","93%","Verified"),("Frontend build","Successful","Verified"),("Seed count","30 articles","Verified"),("Browser screenshots","10 complete PNG files","Verified"),("Remote CI","Not executed","Not verified")])
    if title.startswith("7."):
        for i,name in enumerate(["01-home-dashboard.png","03-language-filter.png","04-category-filter.png","05-search-results.png","06-article-details.png","07-article-summary.png","08-bookmarked-articles.png","10-mobile-responsive-view.png"],1):
            picture(name,f"Figure {i}: {name.replace('-',' ').replace('.png','')} captured from the running application.")
    if title.startswith("8."):
        table(["Ethical control","Implementation","Remaining work"],[("Copyright","Short prewritten text and source fields","Source licensing"),("Transparency","Demo labels","User-facing model disclosure"),("Privacy","No personal records","Production policy"),("Accountability","Local code and tests","Human editorial review")])

doc.add_page_break(); heading("Appendix A — Coverage Traceability")
table(["ID","Activity","Evidence","Status"],[
("1.1","Process models","Chapter 2 comparison","Documented"),("1.2","AI vs traditional SDLC","Chapter 2","Documented"),("1.3","AI tools","Not formally benchmarked","Limitation"),("1.4","Requirement analysis","Role-played reader/editor needs","Documented"),("1.5","AI code generation","Prompt/task log","Documented"),("1.6","Test case generation","12 pytest tests","Executed"),
("2.1–2.8","Agile management","Planning narrative","Simulated"),("3.1–3.8","Development and DevOps","React/FastAPI, build, Docker files","Partial"),("4.1–4.9","Testing and monitoring","pytest, coverage, browser screenshots","Partial"),("5.1–5.6","SQA and ethics","Quality and responsible-AI chapters","Documented")])

doc.add_page_break(); heading("Appendix B — AI Usage and Prompt Log")
table(["Task","AI assistance","Verification","Decision"],[("Workspace inspection","Inspected existing files and PDFs","Manual file review","Kept useful files"),("Backend reset","Generated 30 sample articles/endpoints","API and tests","Modified"),("Frontend","Generated React UI and routing","Build and browser","Modified"),("Report","Generated structure and findings","Manual factual check","Modified"),("PPT","Generated 10-slide structure","Screenshot existence and slide count","Accepted")])

doc.add_page_break(); heading("Appendix C — Test Evidence")
table(["Test result","Actual evidence"],[("12 pytest tests","Collected and passed"),("Coverage","93% total statements"),("React build","Vite production build passed"),("Seed","30 articles via /api/articles"),("Screenshots","10 complete PNG files")])
para("The 93% coverage result and 12-test result were observed on the prior expanded test database. The final seed was restored after tests and API count was independently checked.")

doc.add_page_break(); heading("Appendix D — Known Limitations")
table(["Limit","Impact","Required action"],[("SQLite","Not PostgreSQL","Configure DATABASE_URL"),("Prewritten summaries","No real AI inference","Integrate and evaluate Transformer"),("No external news API","Demonstration only","Use licensed feed ingestion"),("No real editor UI","Backend prototype only","Add authenticated editor pages"),("Template/logo absent","Academic identity incomplete","Insert supplied template/logo"),("CI/CD not run","Remote status unknown","Configure GitHub repository")])

doc.add_page_break(); heading("Appendix E — References")
refs=[
"[1] FastAPI, “FastAPI Documentation,” https://fastapi.tiangolo.com/, accessed Nov. 2026.",
"[2] React, “React Documentation,” https://react.dev/, accessed Nov. 2026.",
"[3] Vite, “Vite Documentation,” https://vite.dev/, accessed Nov. 2026.",
"[4] SQLAlchemy, “SQLAlchemy Documentation,” https://docs.sqlalchemy.org/, accessed Nov. 2026.",
"[5] SQLite, “SQLite Documentation,” https://www.sqlite.org/docs.html, accessed Nov. 2026.",
"[6] pytest, “pytest Documentation,” https://docs.pytest.org/, accessed Nov. 2026.",
"[7] Selenium, “Selenium Documentation,” https://www.selenium.dev/documentation/, accessed Nov. 2026.",
"[8] Docker, “Docker Documentation,” https://docs.docker.com/, accessed Nov. 2026.",
"[9] Hugging Face, “Transformers Documentation,” https://huggingface.co/docs/transformers/, accessed Nov. 2026.",
"[10] K. Beck et al., “Manifesto for Agile Software Development,” 2001.",
"[11] IEEE, “ISO/IEC/IEEE 29119-2:2021, Software and systems engineering — Software testing,” ISO, 2021.",
"[12] K. Schwaber and J. Sutherland, “The Scrum Guide,” 2020.",
]
for x in refs: para(x)
doc.save(DOCX)

# PDF mirror
ps=getSampleStyleSheet()
body=ParagraphStyle("Body",parent=ps["BodyText"],fontName="Times-Roman",fontSize=10.5,leading=15.75,alignment=TA_JUSTIFY)
H=ParagraphStyle("H",parent=ps["Heading1"],fontName="Times-Roman",fontSize=14,leading=18,textColor=colors.HexColor("#102a43"))
story=[Spacer(1,140),Paragraph("BhashaNews",H),Paragraph("Regional News Aggregator with Summaries",H),Spacer(1,30),Paragraph("Case P57 / Roll Number 57",body),Paragraph("Student: [STUDENT FULL NAME PLACEHOLDER]<br/>USN: [USN PLACEHOLDER]<br/>Guide: [GUIDE NAME PLACEHOLDER]<br/>College: [COLLEGE NAME PLACEHOLDER]",body),PageBreak()]
for title,text in [("Certificate","[STUDENT NAME PLACEHOLDER], guide and approval signatures remain placeholders."),
("Declaration","AI usage is disclosed. Prototype and verified claims are distinguished."),
("Acknowledgement","[GUIDE NAME PLACEHOLDER], [DEPARTMENT PLACEHOLDER] and [COLLEGE NAME PLACEHOLDER] are placeholders."),
("Abstract","BhashaNews is a local React/FastAPI/SQLite demonstration with 30 labelled articles, 12 passing tests, 93% measured coverage and ten genuine screenshots.")]:
    story += [Paragraph(title,H),Spacer(1,10),Paragraph(text,body),PageBreak()]
for title,text in chapters:
    story += [Paragraph(title,H),Spacer(1,8),Paragraph(text,body),Spacer(1,12)]
    if title.startswith("7."):
        for name in ["01-home-dashboard.png","03-language-filter.png","04-category-filter.png","05-search-results.png","06-article-details.png","07-article-summary.png","08-bookmarked-articles.png","10-mobile-responsive-view.png"]:
            path=os.path.join(SCREENSHOTS,name)
            if os.path.exists(path): story += [Image(path,width=420,height=292),Spacer(1,6)]
    story += [PageBreak()]
story += [Paragraph("Appendix A - Coverage Traceability",H),Paragraph("Coverage IDs 1.1–1.6, 2.1–2.8, 3.1–3.8, 4.1–4.9 and 5.1–5.6 are documented honestly in the DOCX report.",body),PageBreak()]
story += [Paragraph("Appendix B - AI Usage and Limitations",H),Paragraph("AI assistance was used for code, tests and documentation and was reviewed. Live Transformer inference, real PostgreSQL, remote CI and 45/50 human summary review were not executed.",body),PageBreak()]
story += [Paragraph("Appendix C - References",H)]
for x in refs: story += [Paragraph(x,body),Spacer(1,4)]
def footer(canvas,doc):
    canvas.saveState(); canvas.setFont("Times-Roman",9); canvas.drawCentredString(A4[0]/2,18,str(canvas.getPageNumber())); canvas.restoreState()
SimpleDocTemplate(PDF,pagesize=A4,leftMargin=72,rightMargin=72,topMargin=72,bottomMargin=72).build(story,onFirstPage=footer,onLaterPages=footer)

# Exactly 10-slide PPT
prs=Presentation()
prs.slide_width=PInches(13.333); prs.slide_height=PInches(7.5)
for i in range(10):
    slide=prs.slides.add_slide(prs.slide_layouts[6])
    bg=slide.background.fill; bg.solid(); bg.fore_color.rgb=RGBColor(245,247,251)
    band=slide.shapes.add_shape(1,PInches(0),PInches(0),PInches(13.333),PInches(.18))
    band.fill.solid(); band.fill.fore_color.rgb=RGBColor(226,165,58); band.line.fill.background()
    title=slide.shapes.add_textbox(PInches(.55),PInches(.45),PInches(12),PInches(.6)).text_frame
    title.text=["BhashaNews — Regional News Aggregator","Problem and Target Users","System Architecture","Methodology and Requirements","AI-Assisted Development","Application Demonstration","Testing and Verification","Quality and Results","Responsible AI","Conclusion and Future Work"][i]
    title.paragraphs[0].font.size=PPt(30); title.paragraphs[0].font.bold=True; title.paragraphs[0].font.name="Arial"; title.paragraphs[0].font.color.rgb=RGBColor(16,42,67)
    body=slide.shapes.add_textbox(PInches(.65),PInches(1.35),PInches(7.5),PInches(5.2)).text_frame
    body.clear()
    texts=[
      ["Case P57 · Roll Number 57","Student: [PLACEHOLDER]","Guide: [PLACEHOLDER]","College: [PLACEHOLDER]"],
      ["Problem: Readers need short, language-filtered regional news","Users: Readers and editors","Scope: browse, search, summaries, bookmarks and source links"],
      ["React + Vite frontend","FastAPI + SQLAlchemy backend","SQLite prototype database","Prewritten summaries; local source demo"],
      ["Incremental, prototype-enhanced approach","Reader/editor requirements","Dashboard, detail and saved-article increments","Academic meetings: simulated, not actual"],
      ["AI assistance: code, tests and documentation","All output reviewed and corrected","No external AI API or API key","No live Transformer inference"],
      ["30 genuine seeded news records","Language/category/search interactions","Article summaries and source attribution","Local bookmark persistence"],
      ["12 pytest tests passed","93% measured statement coverage","React production build passed","10 complete browser screenshots"],
      ["API health and article endpoints verified","Responsive UI verified","No remote green CI claim","Known SQLite/AI limitations disclosed"],
      ["Clearly labelled demonstration content","Original-source fields displayed","No real-world event fabrication","Privacy and copyright boundaries documented"],
      ["Prototype complete and verified locally","Next: licensed feeds and PostgreSQL","Next: evaluated Transformer and editor auth","Fill student/institution placeholders"]
    ][i]
    for j,t in enumerate(texts):
        p=body.add_paragraph() if j else body.paragraphs[0]
        p.text=t; p.font.name="Arial"; p.font.size=PPt(24); p.font.color.rgb=RGBColor(52,71,90); p.space_after=PPt(14)
    if i in (5,):
        slide.shapes.add_picture(os.path.join(SCREENSHOTS,"01-home-dashboard.png"),PInches(8.35),PInches(1.35),width=PInches(4.35),height=PInches(3.05))
    elif i in (6,):
        slide.shapes.add_picture(os.path.join(SCREENSHOTS,"06-article-details.png"),PInches(8.35),PInches(1.35),width=PInches(4.35),height=PInches(3.05))
    elif i in (7,):
        slide.shapes.add_picture(os.path.join(SCREENSHOTS,"10-mobile-responsive-view.png"),PInches(9.25),PInches(1.35),width=PInches(3.0),height=PInches(4.9))
    foot=slide.shapes.add_textbox(PInches(.55),PInches(7.05),PInches(12),PInches(.25)).text_frame
    foot.text=f"BHASHANEWS · P57 · {i+1}/10"
    foot.paragraphs[0].font.name="Arial"; foot.paragraphs[0].font.size=PPt(9); foot.paragraphs[0].font.color.rgb=RGBColor(102,118,138)
prs.save(PPTX)

print(DOCX); print(PDF); print(PPTX)
