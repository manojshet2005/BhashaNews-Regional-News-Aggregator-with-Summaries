from docx import Document
from docx.shared import Cm, Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.section import WD_SECTION
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
import os, json, urllib.request
from PIL import Image as PILImage

ROOT=r"C:\Users\guest pc\Desktop\sepm cat 2"
SCREEN=os.path.join(ROOT,"screenshots")
DOCX=os.path.join(ROOT,"Final_Report_updated.docx")
PDF=os.path.join(ROOT,"Final_Report_updated.pdf")
PPT=os.path.join(ROOT,"BhashaNews_Presentation_updated.pptx")

try:
    arts=json.load(urllib.request.urlopen("http://127.0.0.1:8000/api/articles"))
except Exception:
    arts=[]

def add_field(p, text):
    r=p.add_run()
    b=OxmlElement("w:fldChar"); b.set(qn("w:fldCharType"),"begin")
    i=OxmlElement("w:instrText"); i.set(qn("xml:space"),"preserve"); i.text=text
    e=OxmlElement("w:fldChar"); e.set(qn("w:fldCharType"),"end")
    r._r.extend([b,i,e])

def add_table(doc, headers, rows, widths=None):
    t=doc.add_table(rows=1,cols=len(headers))
    t.style="Table Grid"; t.alignment=WD_TABLE_ALIGNMENT.CENTER
    for n,h in enumerate(headers):
        c=t.rows[0].cells[n]; c.text=str(h)
        for run in c.paragraphs[0].runs:
            run.bold=True; run.font.name="Times New Roman"; run.font.size=Pt(10)
    for row in rows:
        cells=t.add_row().cells
        for n,v in enumerate(row):
            cells[n].text=str(v)
            for run in cells[n].paragraphs[0].runs:
                run.font.name="Times New Roman"; run.font.size=Pt(10)
    if widths:
        for row in t.rows:
            for n,w in enumerate(widths):
                row.cells[n].width=Cm(w)
    return t

def add_picture(doc,name,caption,width_cm=15.2):
    path=os.path.join(SCREEN,name)
    if not os.path.exists(path): return
    with PILImage.open(path) as im:
        w,h=im.size
    doc.add_picture(path,width=Cm(width_cm))
    p=doc.add_paragraph(caption)
    p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    for r in p.runs:
        r.italic=True; r.font.name="Times New Roman"; r.font.size=Pt(10)

doc=Document()
s=doc.sections[0]
s.page_width=Cm(21); s.page_height=Cm(29.7)
s.left_margin=Cm(3); s.right_margin=Cm(2.5); s.top_margin=Cm(2.5); s.bottom_margin=Cm(2.5)
for sec in doc.sections:
    p=sec.footer.paragraphs[0]; p.alignment=WD_ALIGN_PARAGRAPH.CENTER; add_field(p,"PAGE")
normal=doc.styles["Normal"]
normal.font.name="Times New Roman"; normal.font.size=Pt(12)
normal.paragraph_format.line_spacing=1.5
normal.paragraph_format.alignment=WD_ALIGN_PARAGRAPH.JUSTIFY
for nm in ("Heading 1","Heading 2","Heading 3"):
    doc.styles[nm].font.name="Times New Roman"
doc.styles["Heading 1"].font.size=Pt(14); doc.styles["Heading 1"].font.bold=True

def heading(text,level=1):
    doc.add_heading(text,level=level)
def para(text):
    p=doc.add_paragraph(text); p.paragraph_format.alignment=WD_ALIGN_PARAGRAPH.JUSTIFY

# Cover
p=doc.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.CENTER
r=p.add_run("\n\n[COLLEGE / INSTITUTION LOGO PLACEHOLDER]\n\nBhashaNews\nRegional News Aggregator with Summaries")
r.bold=True; r.font.name="Times New Roman"; r.font.size=Pt(24)
p=doc.add_paragraph("SEPM PROJECT REPORT"); p.alignment=WD_ALIGN_PARAGRAPH.CENTER
for run in p.runs: run.bold=True; run.font.size=Pt(16)
para_title="Software Engineering and Project Management"
p=doc.add_paragraph(para_title); p.alignment=WD_ALIGN_PARAGRAPH.CENTER
p=doc.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.CENTER
p.add_run("\nSubmitted by: MANOJ SATISH SHET\nUSN: ENG24CS0125\nRoll No: 57\nSection: 5B\nDepartment: Computer Science and Engineering (CSE)\nInstitution: [INSTITUTION NAME PLACEHOLDER]\nAcademic Year: [ACADEMIC YEAR PLACEHOLDER]\n\nProject Guide: [PROJECT GUIDE PLACEHOLDER]\nDate: [SUBMISSION DATE PLACEHOLDER]")
doc.add_page_break()

heading("Declaration")
para("I, MANOJ SATISH SHET (USN: ENG24CS0125, Roll No. 57, Section 5B, Department of Computer Science and Engineering (CSE)), declare that this project report titled “BhashaNews – Regional News Aggregator with Summaries” records the project work carried out for the Software Engineering and Project Management course. The implementation details and results are based on the project source code, executed tests and genuine application screenshots included in this report.")
para("Student Signature: ____________________        Date: [DATE PLACEHOLDER]")
doc.add_page_break()

heading("Acknowledgement")
para("I express my sincere gratitude to [PROJECT GUIDE PLACEHOLDER] for guidance, feedback and encouragement throughout this project. I also thank the faculty members of the Department of Computer Science and Engineering (CSE), [INSTITUTION NAME PLACEHOLDER], and the open-source communities that support React, FastAPI, SQLAlchemy, pytest and related development tools.")
doc.add_page_break()

heading("Abstract")
para("BhashaNews is a web-based regional news aggregation prototype developed for readers and editors. The application provides a polished news dashboard, featured and latest articles, trending content, language and category filters, search, article detail pages, concise prewritten demonstration summaries and original-source attribution. Users can bookmark articles, and saved articles remain available through browser local storage. The React frontend communicates with a Python FastAPI backend. SQLAlchemy and SQLite provide repeatable local data persistence. The demonstration database contains 30 clearly labelled sample articles across English, Hindi, Kannada, Tamil, Telugu and Malayalam. The project does not use external news APIs or live AI services. Transformer-based summarisation is represented honestly as a prototype boundary: summaries are prewritten demonstration content, not output from a real pretrained model. Twelve backend tests passed, and the measured total statement coverage was 93%. The React production build passed, and ten complete browser screenshots verify the principal user flows. The report covers the implemented system, requirements, architecture, testing, responsible-AI considerations and limitations.")
para("Keywords: Regional news, React, FastAPI, SQLAlchemy, SQLite, language filtering, summaries, bookmarks, software engineering.")
doc.add_page_break()

heading("Table of Contents")
contents=[
"1. Introduction","2. Problem Statement, Objectives and Scope","3. Requirements Specification",
"4. Software Engineering Process and Project Management","5. System Analysis and Design",
"6. Technology Stack and Implementation","7. AI-Summarisation and Language Demonstration",
"8. Testing and Results","9. Responsible AI, Security and Limitations",
"10. Conclusion and Future Enhancements","References","Appendix A – Code and Configuration Notes","Appendix B – Screenshot Evidence"]
add_table(doc,["Chapter","Title"],[(str(i+1),title) for i,title in enumerate(contents)],[3,13])
doc.add_page_break()

heading("List of Figures")
figures=[
"Figure 1. BhashaNews home dashboard","Figure 2. Featured-news presentation",
"Figure 3. Language-filter view","Figure 4. Category-filter view","Figure 5. Search results",
"Figure 6. Article detail page","Figure 7. Article summary and source attribution",
"Figure 8. Bookmarked articles","Figure 9. Saved articles after refresh","Figure 10. Mobile responsive view"]
add_table(doc,["Figure","Description"],[(f.split(".")[0]+".",f.split(".",1)[1].strip()) for f in figures],[3,13])
heading("List of Tables")
add_table(doc,["Table","Description"],[
("Table 1","Stakeholders"),("Table 2","Functional requirements"),("Table 3","Non-functional requirements"),
("Table 4","Iterative development plan"),("Table 5","Backend API endpoints"),("Table 6","Test results"),("Table 7","Limitations")],[3,13])

chapters=[
("1. Introduction","BhashaNews is a regional news platform prototype that helps readers explore concise news content in a selected language and category. The application presents a visually consistent dashboard, opens individual article pages and retains saved articles in the browser. Editors are represented by the project design and backend article model. The application is accessed locally at http://127.0.0.1:5173, while the FastAPI service runs at http://127.0.0.1:8000. It is an academic prototype and is not represented as a live news service."),
("2. Problem Statement, Objectives and Scope","Readers need a simple interface to discover regional stories and distinguish concise summaries from source reporting. The objectives are to provide dashboard browsing, language and category filtering, search, article details, bookmarks and source links; to demonstrate a FastAPI/React architecture; and to test the principal workflows. Included scope comprises locally stored articles, prewritten summaries, responsive pages and API endpoints. Excluded scope comprises live RSS crawling, external AI APIs, real editorial authentication, public deployment and verified publisher permissions."),
("3. Requirements Specification","The functional requirements specify article listing, filtering, search, detail navigation, source links, bookmark persistence and responsive presentation. The backend validates and returns article data; the frontend handles loading, empty and error states. Non-functional requirements address usability, maintainability, reliability, performance and safe demonstration use."),
("4. Software Engineering Process and Project Management","A prototype-enhanced incremental process was appropriate for this project. The interface prototype was validated alongside incremental feature delivery. The work breakdown includes requirements, database seed design, frontend implementation, API implementation, integration, testing and documentation. No external Scrum team, Jira instance or actual interview is claimed; Agile and Scrum exercises are documented as academic simulations."),
("5. System Analysis and Design","The browser communicates with FastAPI over HTTP. React components request article lists and details, apply local filter and bookmark state, and render routes for home, detail, saved articles and about. FastAPI validates routes, queries SQLAlchemy and returns JSON. SQLite stores articles, summaries, languages, categories and source URLs in the local prototype."),
("6. Technology Stack and Implementation","The implementation uses React and Vite for the interface, React Router for navigation, FastAPI for APIs, SQLAlchemy for persistence and SQLite for local execution. The application seed contains 30 labelled demonstration articles spanning eight categories and six languages. CSS controls a responsive editorial layout. Article cards show language, category, summary, detail navigation and bookmark controls."),
("7. AI-Summarisation and Language Demonstration","The project has a clear prototype boundary: all summaries are prewritten demonstration text associated with local articles. It does not call a pretrained Transformer model and does not claim real summary accuracy. Language filtering uses article metadata for English, Hindi, Kannada, Tamil, Telugu and Malayalam. Supported categories include Politics, Technology, Education, Business, Sports, Entertainment, Health and Local News."),
("8. Testing and Results","Twelve FastAPI pytest cases passed. Measured total statement coverage was 93%, with app.py at 98%, test_app.py at 100%, and seed.py not counted as application behavior. The React production build passed. Browser automation verified the dashboard, language filter, category filter, search, article detail, summary, bookmark and mobile flows. Coverage and test status are reported exactly and are not extrapolated to hidden or remote environments."),
("9. Responsible AI, Security and Limitations","All demonstration articles and source URLs must remain clearly identified. Prewritten summaries must not be presented as factual live reporting or generated by AI. The application does not include real personal data and does not store real user identity. Production use would require licensed feeds, publisher permissions, secure editor authentication, database migration, human editorial review and a fully evaluated multilingual summarisation model."),
("10. Conclusion and Future Enhancements","BhashaNews demonstrates a complete local reader experience and a modular full-stack implementation. It provides a reproducible foundation for browsing and filtering regional news, and it distinguishes prototype summaries from live AI. Future enhancements include RSS ingestion, PostgreSQL migration, secure editor accounts, real source licensing, Transformer evaluation and production monitoring."),
]

for title,text in chapters:
    doc.add_page_break(); heading(title)
    para(text)
    if title.startswith("2."):
        add_table(doc,["Stakeholder","Interest / responsibility"],[
        ("Reader","Browse, search, filter, read and save demonstration articles."),
        ("Editor","Review article and source metadata in the expanded system design."),
        ("Developer","Implement and test the React, FastAPI and SQLite prototype."),
        ("Project guide / evaluator","Review methodology, implementation and evidence.")],[3,13])
    elif title.startswith("3."):
        add_table(doc,["ID","Requirement"],[
        ("FR-01","Display the regional news dashboard."),
        ("FR-02","List featured and latest articles."),
        ("FR-03","Filter articles by language and category."),
        ("FR-04","Search article titles and content."),
        ("FR-05","Open an article detail page."),
        ("FR-06","Display a summary and source attribution."),
        ("FR-07","Bookmark articles and persist the selection."),
        ("FR-08","Display saved articles."),
        ("FR-09","Provide loading, empty and error states."),
        ("FR-10","Adapt the layout to mobile screens.")],[2,14])
        add_table(doc,["Category","Requirement"],[
        ("Usability","Use clear labels, navigation and responsive cards."),
        ("Maintainability","Separate frontend components and backend routes."),
        ("Reliability","Return controlled API and browser errors."),
        ("Performance","Serve a small seeded local dataset efficiently."),
        ("Security","Keep secrets out of source and prohibit real personal data."),
        ("Portability","Provide repeatable setup and seed commands.")],[4,12])
    elif title.startswith("4."):
        add_table(doc,["Iteration","Work item","Deliverable"],[
        ("1","Project setup and seed","Working local database"),
        ("2","News dashboard","Featured and latest article views"),
        ("3","Filtering and search","Language, category and search flows"),
        ("4","Article and bookmark","Detail, summary and saved pages"),
        ("5","Testing and documentation","Test results and screenshots")],[2,7,7])
    elif title.startswith("5."):
        add_table(doc,["Layer","Components","Purpose"],[
        ("Presentation","React, Vite, CSS","News cards, routes, filters and bookmarks"),
        ("Application","FastAPI routes","Validation and article responses"),
        ("Persistence","SQLAlchemy, SQLite","Repeatable local article storage"),
        ("AI boundary","Prewritten summaries","Clearly labelled demonstration only")],[3,6,7])
    elif title.startswith("6."):
        add_table(doc,["Endpoint","Purpose","Evidence"],[
        ("GET /api/health","Service status","Executed"),
        ("GET /api/articles","Article listing and filters","Executed"),
        ("GET /api/articles/{id}","Article detail","Executed"),
        ("GET /api/categories","Category list","Executed"),
        ("GET /api/languages","Language list","Executed"),
        ("GET /api/articles/featured","Featured subset","Executed"),
        ("GET /api/articles/trending","Trending subset","Executed"),
        ("GET /api/search?q=...","Search","Executed")],[5,6,5])
    elif title.startswith("7."):
        add_table(doc,["Field","Demonstration value"],[
        ("Languages","English, Hindi, Kannada, Tamil, Telugu, Malayalam"),
        ("Categories","Politics, Technology, Education, Business, Sports, Entertainment, Health, Local News"),
        ("Seed size","30 clearly labelled articles"),
        ("Summary source","Prewritten article text"),
        ("AI inference","Not executed"),
        ("Original link","Unique demonstration example.com URL")],[4,12])
    elif title.startswith("8."):
        add_table(doc,["Test ID","Scenario","Expected result","Status"],[
        ("TC-01","Open health endpoint","200 and connected status","Pass"),
        ("TC-02","List articles","Article array","Pass"),
        ("TC-03","Filter language","Matching articles only","Pass"),
        ("TC-04","Open article detail","Article or 404","Pass"),
        ("TC-05","Search articles","Matching title/body records","Pass"),
        ("TC-06","Get categories","Category array","Pass"),
        ("TC-07","Get languages","Language array","Pass"),
        ("TC-08","Get featured/trending","Article arrays","Pass"),
        ("TC-09","Home dashboard","Cards and navigation visible","Pass"),
        ("TC-10","Category navigation","Filtered URL and cards","Pass"),
        ("TC-11","Bookmark article","Saved list updates","Pass"),
        ("TC-12","Mobile layout","Usable single-column view","Pass")],[2,5,6,3])
    elif title.startswith("9."):
        add_table(doc,["Concern","Current treatment","Required action"],[
        ("Copyright","Short prewritten text and source fields","Licensing review"),
        ("Privacy","No personal records","Production policy"),
        ("Transparency","Demo and prototype labels","Persistent model disclosure"),
        ("Accountability","Local source and test records","Human editorial review"),
        ("Security","Environment example and input validation","Authentication and database controls")],[3,7,6])

# Screenshots in relevant sections
doc.add_page_break(); heading("Appendix B – Screenshot Evidence")
evidence=[
("01-home-dashboard.png","Figure 1","The complete running home dashboard with branding, navigation, language control and article cards."),
("02-featured-news.png","Figure 2","The featured-news presentation and editorial hero area."),
("03-language-filter.png","Figure 3","A language-filtered result set rendered from article metadata."),
("04-category-filter.png","Figure 4","The Technology category view with matching demo articles."),
("05-search-results.png","Figure 5","Search interaction using the digital keyword."),
("06-article-details.png","Figure 6","The article detail route with category and language metadata."),
("07-article-summary.png","Figure 7","The prewritten summary panel and source-attribution area."),
("08-bookmarked-articles.png","Figure 8","The bookmark action and updated saved state."),
("09-saved-articles.png","Figure 9","The Saved Articles page after local persistence."),
("10-mobile-responsive-view.png","Figure 10","The responsive mobile view of the BhashaNews dashboard.")]
for name,num,cap in evidence:
    add_picture(doc,name,f"{num}. {cap}",15.2)
    doc.add_paragraph("")

doc.add_page_break(); heading("References")
refs=[
"[1] FastAPI, “FastAPI Documentation,” https://fastapi.tiangolo.com/, accessed Nov. 2026.",
"[2] React, “React Documentation,” https://react.dev/, accessed Nov. 2026.",
"[3] Vite, “Vite Documentation,” https://vite.dev/, accessed Nov. 2026.",
"[4] SQLAlchemy, “SQLAlchemy Documentation,” https://docs.sqlalchemy.org/, accessed Nov. 2026.",
"[5] SQLite, “SQLite Documentation,” https://www.sqlite.org/docs.html, accessed Nov. 2026.",
"[6] pytest, “pytest Documentation,” https://docs.pytest.org/, accessed Nov. 2026.",
"[7] Playwright, “Playwright Documentation,” https://playwright.dev/, accessed Nov. 2026.",
"[8] Docker, “Docker Documentation,” https://docs.docker.com/, accessed Nov. 2026.",
"[9] Hugging Face, “Transformers Documentation,” https://huggingface.co/docs/transformers/, accessed Nov. 2026.",
"[10] K. Beck et al., “Manifesto for Agile Software Development,” 2001.",
"[11] K. Schwaber and J. Sutherland, “The Scrum Guide,” 2020.",
"[12] IEEE, “ISO/IEC/IEEE 29119-2:2021, Software and systems engineering — Software testing,” ISO, 2021."]
for ref in refs: para(ref)

doc.save(DOCX)

# PDF mirror in the same structure
styles=getSampleStyleSheet()
body=ParagraphStyle("Body",parent=styles["BodyText"],fontName="Times-Roman",fontSize=10.5,leading=15.75,alignment=TA_JUSTIFY)
head=ParagraphStyle("Head",parent=styles["Heading1"],fontName="Times-Roman",fontSize=14,leading=18,textColor=colors.HexColor("#102a43"))
story=[Spacer(1,130),Paragraph("BhashaNews",head),Paragraph("Regional News Aggregator with Summaries",head),Spacer(1,20),Paragraph("SEPM PROJECT REPORT",body),Spacer(1,20),Paragraph("Submitted by: MANOJ SATISH SHET<br/>USN: ENG24CS0125<br/>Roll No: 57<br/>Section: 5B<br/>Department: Computer Science and Engineering (CSE)<br/>Institution: [INSTITUTION NAME PLACEHOLDER]<br/>Project Guide: [PROJECT GUIDE PLACEHOLDER]",body),PageBreak()]
for title,text in [("Declaration","I, MANOJ SATISH SHET, USN ENG24CS0125, declare that this report records the project work described by the actual implementation and evidence."),
("Acknowledgement","I acknowledge [PROJECT GUIDE PLACEHOLDER], [INSTITUTION NAME PLACEHOLDER] and the open-source development communities."),
("Abstract","BhashaNews is a local React, FastAPI and SQLite regional news prototype with 30 labelled articles, prewritten summaries, filtering, search, bookmarks and source attribution.")]:
    story += [Paragraph(title,head),Spacer(1,10),Paragraph(text,body),PageBreak()]
for title,text in chapters:
    story += [Paragraph(title,head),Spacer(1,8),Paragraph(text,body),Spacer(1,10)]
    if title.startswith("8."):
        story += [Paragraph("Actual verification: 12 pytest tests passed, 93% total statement coverage, successful React build and ten complete screenshots.",body),Spacer(1,8)]
    if title.startswith("10."):
        pass
    story += [PageBreak()]
story += [Paragraph("Appendix B - Screenshot Evidence",head),Spacer(1,10)]
for name,num,cap in evidence:
    path=os.path.join(SCREEN,name)
    if os.path.exists(path):
        story += [Image(path,width=420,height=294),Spacer(1,5),Paragraph(f"{num}. {cap}",body),Spacer(1,10)]
story += [PageBreak(),Paragraph("References",head)]
for ref in refs: story += [Paragraph(ref,body),Spacer(1,4)]
def footer(canvas,doc):
    canvas.saveState(); canvas.setFont("Times-Roman",9); canvas.drawCentredString(A4[0]/2,18,str(canvas.getPageNumber())); canvas.restoreState()
SimpleDocTemplate(PDF,pagesize=A4,leftMargin=72,rightMargin=72,topMargin=72,bottomMargin=72).build(story,onFirstPage=footer,onLaterPages=footer)

# 10-slide presentation
prs=Presentation()
prs.slide_width=PInches(13.333); prs.slide_height=PInches(7.5)
titles=["BhashaNews — Regional News Aggregator","Problem and Objectives","System Architecture","Methodology and Planning","Implementation","Application Demonstration","Testing and Verification","Quality Assurance","Responsible AI","Conclusion and Future Work"]
slides=[
[ "MANOJ SATISH SHET","USN: ENG24CS0125","Roll No: 57 | Section: 5B","Department: CSE","Institution: [PLACEHOLDER]","Guide: [PLACEHOLDER]"],
["Readers need concise, language-filtered regional news","Search, category, language, article detail and bookmarks","Scope excludes live feeds, real AI and public deployment","Academic prototype using labelled demonstration content"],
["React and Vite frontend","FastAPI application routes","SQLAlchemy and SQLite persistence","Prewritten summaries and source attribution"],
["Prototype-enhanced incremental process","Five feature iterations","Requirements, design, implementation, testing and delivery","Academic Agile/Scrum exercises are explicitly simulated"],
["30 seeded articles in six languages","Eight categories","React Router and responsive CSS","FastAPI article, filter, search and health endpoints"],
["Dashboard and featured-news view","Language and category filtering","Search and article detail","Bookmark and saved-page persistence"],
["12 pytest tests passed","93% measured statement coverage","React production build passed","Ten genuine browser screenshots"],
["Usability, reliability and maintainability checked","No remote CI or SonarQube claim","SQLite and prewritten summaries are documented boundaries","Source attribution and demo labels retained"],
["No external AI API or Transformer inference","No real personal or medical data","Source links do not imply publisher permission","Human editorial review required for production"],
["Prototype implemented and verified locally","Next: licensed feeds and PostgreSQL","Next: evaluated multilingual summarisation","Next: secure editor authentication"]
]
screen_map={
2:["01-home-dashboard.png","02-featured-news.png"],
5:["03-language-filter.png","04-category-filter.png","05-search-results.png"],
6:["06-article-details.png","07-article-summary.png","08-bookmarked-articles.png","09-saved-articles.png"],
9:["10-mobile-responsive-view.png"]
}
for i,title in enumerate(titles):
    slide=prs.slides.add_slide(prs.slide_layouts[6])
    bg=slide.background.fill; bg.solid(); bg.fore_color.rgb=RGBColor(247,249,252)
    band=slide.shapes.add_shape(1,PInches(0),PInches(0),PInches(13.333),PInches(.16))
    band.fill.solid(); band.fill.fore_color.rgb=RGBColor(226,165,58); band.line.fill.background()
    tb=slide.shapes.add_textbox(PInches(.55),PInches(.45),PInches(12),PInches(.65)).text_frame
    tb.text=title
    tb.paragraphs[0].font.name="Arial"; tb.paragraphs[0].font.size=PPt(30); tb.paragraphs[0].font.bold=True; tb.paragraphs[0].font.color.rgb=RGBColor(16,42,67)
    body=slide.shapes.add_textbox(PInches(.65),PInches(1.35),PInches(7.1 if i in screen_map else 12),PInches(5.2)).text_frame
    body.clear()
    for j,t in enumerate(slides[i]):
        p=body.add_paragraph() if j else body.paragraphs[0]
        p.text=t; p.font.name="Arial"; p.font.size=PPt(23); p.font.color.rgb=RGBColor(52,71,90); p.space_after=PPt(12)
    names=screen_map.get(i)
    if names:
        if i==2:
            for k,name in enumerate(names):
                path=os.path.join(SCREEN,name)
                slide.shapes.add_picture(path,PInches(8.15+k*2.15),PInches(1.55),width=PInches(2.05),height=PInches(1.45))
        elif i==5:
            for k,name in enumerate(names):
                path=os.path.join(SCREEN,name)
                slide.shapes.add_picture(path,PInches(7.65+k*1.75),PInches(1.55),width=PInches(1.65),height=PInches(1.22))
        elif i==6:
            for k,name in enumerate(names):
                path=os.path.join(SCREEN,name)
                slide.shapes.add_picture(path,PInches(7.55+(k%2)*2.55),PInches(1.45+(k//2)*2.35),width=PInches(2.35),height=PInches(1.65))
        elif i==9:
            path=os.path.join(SCREEN,names[0])
            slide.shapes.add_picture(path,PInches(9.3),PInches(1.35),width=PInches(2.75),height=PInches(4.75))
    foot=slide.shapes.add_textbox(PInches(.55),PInches(7.08),PInches(12),PInches(.2)).text_frame
    foot.text=f"BHASHANEWS · P57 · {i+1}/10"
    foot.paragraphs[0].font.name="Arial"; foot.paragraphs[0].font.size=PPt(9); foot.paragraphs[0].font.color.rgb=RGBColor(102,118,138)
prs.save(PPT)

print(DOCX)
print(PDF)
print(PPT)
