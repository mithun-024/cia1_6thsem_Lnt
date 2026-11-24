#!/usr/bin/env python3
"""
Generates the CIA-1 Assignment Word Document (.docx) for:
Christ University | Server Integration and MongoDB Connectivity | CIA-1
Student: MITHUN G
"""

import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn

def create_document():
    doc = docx.Document()

    # Page Margins: 1 inch on all sides
    for section in doc.sections:
        section.top_margin = Inches(1.0)
        section.bottom_margin = Inches(1.0)
        section.left_margin = Inches(1.0)
        section.right_margin = Inches(1.0)

    # Styles & Colors
    NAVY = RGBColor(0x00, 0x33, 0x66)       # Primary headings
    SLATE = RGBColor(0x2C, 0x3E, 0x50)      # Subheadings
    DARK_TEXT = RGBColor(0x33, 0x33, 0x33)  # Body text
    MUTED = RGBColor(0x66, 0x66, 0x66)      # Captions

    def set_cell_shading(cell, color_hex):
        shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{color_hex}"/>')
        cell._tc.get_or_add_tcPr().append(shd)

    def set_cell_margins(cell, top=120, bottom=120, left=180, right=180):
        tcPr = cell._tc.get_or_add_tcPr()
        tcMar = parse_xml(
            f'<w:tcMar {nsdecls("w")}><w:top w:w="{top}" w:type="dxa"/><w:bottom w:w="{bottom}" w:type="dxa"/>'
            f'<w:left w:w="{left}" w:type="dxa"/><w:right w:w="{right}" w:type="dxa"/></w:tcMar>'
        )
        tcPr.append(tcMar)

    def add_code_box(code_text):
        tbl = doc.add_table(rows=1, cols=1)
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        tbl.autofit = False
        cell = tbl.cell(0, 0)
        cell.width = Inches(6.5)
        set_cell_shading(cell, "F7F8FA")
        set_cell_margins(cell, top=120, bottom=120, left=180, right=180)
        borders = parse_xml(
            f'<w:tcBorders {nsdecls("w")}>'
            f'<w:top w:val="single" w:sz="4" w:space="0" w:color="D0D7DE"/>'
            f'<w:bottom w:val="single" w:sz="4" w:space="0" w:color="D0D7DE"/>'
            f'<w:left w:val="single" w:sz="20" w:space="0" w:color="003366"/>'
            f'<w:right w:val="single" w:sz="4" w:space="0" w:color="D0D7DE"/>'
            f'</w:tcBorders>'
        )
        cell._tc.get_or_add_tcPr().append(borders)
        p = cell.paragraphs[0]
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.line_spacing = 1.15
        label = p.add_run("Source Code:\n")
        label.font.name = "Consolas"
        label.font.size = Pt(8.5)
        label.font.bold = True
        label.font.color.rgb = NAVY
        run = p.add_run(code_text)
        run.font.name = "Consolas"
        run.font.size = Pt(9.5)
        run.font.color.rgb = RGBColor(0x24, 0x29, 0x2E)
        sp = doc.add_paragraph()
        sp.paragraph_format.space_before = Pt(0)
        sp.paragraph_format.space_after = Pt(3)

    def add_output_box(output_text):
        tbl = doc.add_table(rows=1, cols=1)
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        tbl.autofit = False
        cell = tbl.cell(0, 0)
        cell.width = Inches(6.5)
        set_cell_shading(cell, "F0F7F4")
        set_cell_margins(cell, top=100, bottom=100, left=180, right=180)
        borders = parse_xml(
            f'<w:tcBorders {nsdecls("w")}>'
            f'<w:top w:val="single" w:sz="4" w:space="0" w:color="D1E7DD"/>'
            f'<w:bottom w:val="single" w:sz="4" w:space="0" w:color="D1E7DD"/>'
            f'<w:left w:val="single" w:sz="20" w:space="0" w:color="198754"/>'
            f'<w:right w:val="single" w:sz="4" w:space="0" w:color="D1E7DD"/>'
            f'</w:tcBorders>'
        )
        cell._tc.get_or_add_tcPr().append(borders)
        p = cell.paragraphs[0]
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.line_spacing = 1.15
        label = p.add_run("Console Output:\n")
        label.font.name = "Consolas"
        label.font.size = Pt(8.5)
        label.font.bold = True
        label.font.color.rgb = RGBColor(0x0F, 0x51, 0x32)
        run = p.add_run(output_text)
        run.font.name = "Consolas"
        run.font.size = Pt(9.5)
        run.font.color.rgb = RGBColor(0x15, 0x57, 0x24)
        sp = doc.add_paragraph()
        sp.paragraph_format.space_before = Pt(0)
        sp.paragraph_format.space_after = Pt(6)

    def add_q_header(q_num, marks, q_title, q_desc):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(10)
        p.paragraph_format.space_after = Pt(3)
        p.paragraph_format.keep_with_next = True
        
        r_q = p.add_run(f"Q{q_num}. ")
        r_q.font.name = "Arial"
        r_q.font.size = Pt(11)
        r_q.font.bold = True
        r_q.font.color.rgb = NAVY
        
        r_title = p.add_run(q_title)
        r_title.font.name = "Arial"
        r_title.font.size = Pt(11)
        r_title.font.bold = True
        r_title.font.color.rgb = SLATE
        
        r_marks = p.add_run(f"  [{marks} Mark{'s' if marks > 1 else ''}]")
        r_marks.font.name = "Arial"
        r_marks.font.size = Pt(9.5)
        r_marks.font.italic = True
        r_marks.font.color.rgb = MUTED

        if q_desc:
            p2 = doc.add_paragraph()
            p2.paragraph_format.space_before = Pt(0)
            p2.paragraph_format.space_after = Pt(4)
            p2.paragraph_format.keep_with_next = True
            r_d = p2.add_run(f"Question Statement: {q_desc}")
            r_d.font.name = "Arial"
            r_d.font.size = Pt(10)
            r_d.font.italic = True
            r_d.font.color.rgb = RGBColor(0x55, 0x55, 0x55)

    def add_answer_para(text, bold_prefix="Answer: "):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.line_spacing = 1.15
        if bold_prefix:
            rb = p.add_run(bold_prefix)
            rb.font.name = "Arial"
            rb.font.size = Pt(10.5)
            rb.font.bold = True
            rb.font.color.rgb = NAVY
        r = p.add_run(text)
        r.font.name = "Arial"
        r.font.size = Pt(10)
        r.font.color.rgb = DARK_TEXT

    # ---------------- HEADER / COVER ----------------
    h1 = doc.add_paragraph()
    h1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    h1.paragraph_format.space_before = Pt(0)
    h1.paragraph_format.space_after = Pt(2)
    r1 = h1.add_run("CHRIST (DEEMED TO BE UNIVERSITY)")
    r1.font.name = "Arial"
    r1.font.size = Pt(16)
    r1.font.bold = True
    r1.font.color.rgb = NAVY

    h2 = doc.add_paragraph()
    h2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    h2.paragraph_format.space_before = Pt(0)
    h2.paragraph_format.space_after = Pt(4)
    r2 = h2.add_run("Department of Computer Science | L&T EduTech Integration")
    r2.font.name = "Arial"
    r2.font.size = Pt(11)
    r2.font.color.rgb = SLATE

    h3 = doc.add_paragraph()
    h3.alignment = WD_ALIGN_PARAGRAPH.CENTER
    h3.paragraph_format.space_before = Pt(0)
    h3.paragraph_format.space_after = Pt(8)
    r3 = h3.add_run("CONTINUOUS INTERNAL ASSESSMENT - 1 (CIA-1)")
    r3.font.name = "Arial"
    r3.font.size = Pt(13)
    r3.font.bold = True
    r3.font.color.rgb = NAVY

    # Metadata Table
    table = doc.add_table(rows=4, cols=4)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    
    col_widths = [Inches(1.5), Inches(2.0), Inches(1.3), Inches(1.7)]
    meta_data = [
        [("Course:", True), ("Server Integration and MongoDB Connectivity", False), ("Exam:", True), ("CIA - 1 (Theory + Practical)", False)],
        [("Student Name:", True), ("MITHUN G", False), ("Class / Sem:", True), ("6th Semester", False)],
        [("Roll Number:", True), ("[Enter Roll Number Here]", False), ("Date:", True), ("24-Nov-2026", False)],
        [("Total Marks:", True), ("30 Marks (Sec A: 10 + Sec B: 20)", False), ("Status:", True), ("Completed & Ready for Submission", False)]
    ]
    
    for row_idx, row_content in enumerate(meta_data):
        row = table.rows[row_idx]
        for col_idx, (text, is_label) in enumerate(row_content):
            cell = row.cells[col_idx]
            cell.width = col_widths[col_idx]
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            set_cell_margins(cell, top=80, bottom=80, left=100, right=100)
            if is_label:
                set_cell_shading(cell, "EAEFF5")
            else:
                set_cell_shading(cell, "FFFFFF")
            borders = parse_xml(
                f'<w:tcBorders {nsdecls("w")}>'
                f'<w:top w:val="single" w:sz="4" w:space="0" w:color="B0C4DE"/>'
                f'<w:bottom w:val="single" w:sz="4" w:space="0" w:color="B0C4DE"/>'
                f'<w:left w:val="single" w:sz="4" w:space="0" w:color="B0C4DE"/>'
                f'<w:right w:val="single" w:sz="4" w:space="0" w:color="B0C4DE"/>'
                f'</w:tcBorders>'
            )
            cell._tc.get_or_add_tcPr().append(borders)
            p = cell.paragraphs[0]
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(2)
            run = p.add_run(text)
            run.font.name = "Arial"
            run.font.size = Pt(9.5)
            if is_label:
                run.font.bold = True
                run.font.color.rgb = NAVY
            else:
                run.font.color.rgb = DARK_TEXT

    sep = doc.add_paragraph()
    sep.paragraph_format.space_before = Pt(8)
    sep.paragraph_format.space_after = Pt(8)

    # ==============================================================
    # SECTION A: THEORY QUESTIONS
    # ==============================================================
    sec_a_hdr = doc.add_paragraph()
    sec_a_hdr.paragraph_format.space_before = Pt(12)
    sec_a_hdr.paragraph_format.space_after = Pt(4)
    sec_a_hdr.paragraph_format.keep_with_next = True
    r_sa = sec_a_hdr.add_run("SECTION A – THEORY QUESTIONS (10 × 1 Mark = 10 Marks)")
    r_sa.font.name = "Arial"
    r_sa.font.size = Pt(12)
    r_sa.font.bold = True
    r_sa.font.color.rgb = NAVY

    doc.add_paragraph(
        "Instructions: Answer all questions in this section concisely and accurately."
    ).runs[0].font.size = Pt(9.5)

    # Q1
    add_q_header(1, 1, "What is JSON? Expand JSON.", "")
    add_answer_para(
        "Expansion: JSON stands for JavaScript Object Notation.\n"
        "Definition: JSON is an open standard, lightweight, text-based data interchange format designed for human readability and machine parsing. "
        "It is language-independent and represents structured data using simple key-value pairs (objects) and ordered lists of values (arrays). "
        "JSON has become the de-facto standard for data transmission between servers and client-side web applications."
    )

    # Q2
    add_q_header(2, 1, "What is XML? Expand XML.", "")
    add_answer_para(
        "Expansion: XML stands for Extensible Markup Language.\n"
        "Definition: XML is a W3C-recommended markup language designed to store and transport structured data in a software- and hardware-independent manner. "
        "Unlike HTML, which focuses on presenting data, XML focuses on describing data using custom, user-defined hierarchical tags and a tree-like document structure."
    )

    # Q3
    add_q_header(3, 1, "Write any two differences between JSON and XML.", "")
    p_q3 = doc.add_paragraph()
    p_q3.paragraph_format.space_before = Pt(2)
    p_q3.paragraph_format.space_after = Pt(4)
    r_q3 = p_q3.add_run("Answer: Key differences between JSON and XML include:")
    r_q3.font.name = "Arial"
    r_q3.font.size = Pt(10)
    r_q3.font.bold = True
    r_q3.font.color.rgb = NAVY

    # Table for Q3
    diff_table = doc.add_table(rows=3, cols=3)
    diff_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    diff_table.autofit = False
    col_w = [Inches(1.8), Inches(2.3), Inches(2.4)]
    diff_headers = ["Comparison Feature", "JSON (JavaScript Object Notation)", "XML (Extensible Markup Language)"]
    diff_row1 = ["1. Data Structure & Syntax", "Uses key-value pairs and arrays without closing tags. Minimal and lightweight.", "Uses opening and closing tags (<tag>...</tag>). More verbose with higher overhead."]
    diff_row2 = ["2. Parsing Speed & Efficiency", "Parsed natively in JavaScript using JSON.parse() with high performance.", "Requires an XML DOM Parser (DOMParser) to traverse nodes, making parsing slower."]
    
    for c_idx, h_text in enumerate(diff_headers):
        cell = diff_table.rows[0].cells[c_idx]
        cell.width = col_w[c_idx]
        set_cell_shading(cell, "003366")
        set_cell_margins(cell, top=80, bottom=80, left=100, right=100)
        p = cell.paragraphs[0]
        run = p.add_run(h_text)
        run.font.name = "Arial"
        run.font.size = Pt(9)
        run.font.bold = True
        run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    for r_idx, r_data in enumerate([diff_row1, diff_row2], start=1):
        for c_idx, val in enumerate(r_data):
            cell = diff_table.rows[r_idx].cells[c_idx]
            cell.width = col_w[c_idx]
            set_cell_shading(cell, "F9FAFB" if r_idx % 2 == 1 else "FFFFFF")
            set_cell_margins(cell, top=60, bottom=60, left=100, right=100)
            borders = parse_xml(
                f'<w:tcBorders {nsdecls("w")}>'
                f'<w:top w:val="single" w:sz="4" w:space="0" w:color="D0D7DE"/>'
                f'<w:bottom w:val="single" w:sz="4" w:space="0" w:color="D0D7DE"/>'
                f'<w:left w:val="single" w:sz="4" w:space="0" w:color="D0D7DE"/>'
                f'<w:right w:val="single" w:sz="4" w:space="0" w:color="D0D7DE"/>'
                f'</w:tcBorders>'
            )
            cell._tc.get_or_add_tcPr().append(borders)
            p = cell.paragraphs[0]
            run = p.add_run(val)
            run.font.name = "Arial"
            run.font.size = Pt(9)
            run.font.color.rgb = DARK_TEXT

    # Q4
    add_q_header(4, 1, "What are the advantages of using JSON for data exchange?", "")
    add_answer_para(
        "The primary advantages of using JSON for data exchange are:\n"
        "1. Lightweight & Fast: JSON has no closing tag overhead, reducing payload size and resulting in faster network transmissions.\n"
        "2. Native JavaScript Compatibility: JavaScript can parse JSON directly into native objects using JSON.parse() without complex parsers.\n"
        "3. Universal Language Support: Virtually all programming languages (Python, Java, C#, PHP, Go, Ruby) have built-in libraries for JSON.\n"
        "4. Human-Readable & Clean: The key-value structure is clean, easy to read, write, and debug by developers.\n"
        "5. Direct Data Type Mapping: Directly supports native data types like strings, numbers, booleans, arrays, null, and objects."
    )

    # Q5
    add_q_header(5, 1, "What is AJAX? Expand AJAX.", "")
    add_answer_para(
        "Expansion: AJAX stands for Asynchronous JavaScript And XML.\n"
        "Definition: AJAX is a web development technique that allows client-side web applications to send and retrieve data asynchronously from a web server in the background. "
        "This enables web applications to dynamically update portions of a web page using DOM manipulation without interrupting the user experience or requiring a full page refresh."
    )

    # Q6
    add_q_header(6, 1, "Does AJAX require a complete webpage reload to update data? (Yes/No)", "")
    add_answer_para(
        "No.\n"
        "Explanation: AJAX was specifically created to eliminate full-page reloads. It communicates with the server asynchronously in the background and dynamically updates only the necessary parts of the Document Object Model (DOM), saving bandwidth and providing a fast, seamless user experience."
    )

    # Q7
    add_q_header(7, 1, "What is the purpose of the fetch() method in JavaScript?", "")
    add_answer_para(
        "The fetch() method is a modern, promise-based JavaScript API used to perform asynchronous HTTP requests (such as GET, POST, PUT, DELETE) across a network. "
        "It serves as a clean, powerful, and standardized replacement for the older XMLHttpRequest (XHR) object, allowing developers to handle network requests using Promise chaining (.then() and .catch()) or async/await syntax."
    )

    # Q8
    add_q_header(8, 1, "What is the difference between JSON.parse() and JSON.stringify()?", "")
    add_answer_para(
        "The fundamental difference lies in the direction of data transformation:\n"
        "• JSON.parse(): Takes a valid JSON-formatted string and deserializes (converts) it into a native JavaScript object or array.\n"
        "   Example: JSON.parse('{\"name\":\"Ravi\",\"age\":24}')  -->  { name: 'Ravi', age: 24 }\n"
        "• JSON.stringify(): Takes a JavaScript object, array, or value and serializes (converts) it into a standardized JSON string.\n"
        "   Example: JSON.stringify({ name: 'Rahul', age: 23 })  -->  '{\"name\":\"Rahul\",\"age\":23}'"
    )

    # Q9
    add_q_header(9, 1, "What is the purpose of the .then() method in Fetch API?", "")
    add_answer_para(
        "The .then() method is used to handle the successful fulfillment of a JavaScript Promise returned by fetch() or subsequent asynchronous operations. "
        "It takes a callback function that executes when the response is received. In the Fetch API, the first .then() typically parses the raw HTTP response stream into usable data (e.g., response.json()), and chained .then() blocks handle and render the resolved data."
    )

    # Q10
    add_q_header(10, 1, "What is the purpose of the .catch() method in Fetch API?", "")
    add_answer_para(
        "The .catch() method is used for centralized error handling in a Promise chain. "
        "It executes whenever a Promise in the chain is rejected—such as due to a network connection failure, DNS failure, request abort, or any error explicitly thrown within a preceding .then() block. It prevents uncaught promise exceptions and allows developers to log errors and display friendly feedback to users."
    )

    # ==============================================================
    # SECTION B: PRACTICAL QUESTIONS
    # ==============================================================
    doc.add_page_break()

    sec_b_hdr = doc.add_paragraph()
    sec_b_hdr.paragraph_format.space_before = Pt(12)
    sec_b_hdr.paragraph_format.space_after = Pt(4)
    sec_b_hdr.paragraph_format.keep_with_next = True
    r_sb = sec_b_hdr.add_run("SECTION B – PRACTICAL QUESTIONS (10 × 2 Marks = 20 Marks)")
    r_sb.font.name = "Arial"
    r_sb.font.size = Pt(12)
    r_sb.font.bold = True
    r_sb.font.color.rgb = NAVY

    doc.add_paragraph(
        "Instructions: Write clean code and demonstrate the console output for each practical problem."
    ).runs[0].font.size = Pt(9.5)

    # Practical Q1
    add_q_header(
        1, 2, "Create a JSON Object",
        "Create a JSON object containing the following student information: Name: Arun, Age: 25, Department: IT."
    )
    add_code_box(
        '// q1.js - Create a JSON object representing student information\n'
        'let student = {\n'
        '    "name": "Arun",\n'
        '    "age": 25,\n'
        '    "department": "IT"\n'
        '};\n\n'
        'console.log("Student JSON Object:");\n'
        'console.log(student);'
    )
    add_output_box(
        "Student JSON Object:\n"
        "{ name: 'Arun', age: 25, department: 'IT' }"
    )
    add_answer_para(
        "Explanation: A JavaScript object is declared using standard JSON key-value formatting with string keys ('name', 'age', 'department') and logged to the console using console.log().",
        bold_prefix="Practical Note: "
    )

    # Practical Q2
    add_q_header(
        2, 2, "Access JSON Properties",
        "Consider the following JSON data:\n{\n  \"name\": \"Priya\",\n  \"age\": 22,\n  \"city\": \"Bangalore\"\n}\nWrite JavaScript code to display the student's name and city."
    )
    add_code_box(
        '// q2.js - Access and display JSON object properties\n'
        'let student = {\n'
        '    "name": "Priya",\n'
        '    "age": 22,\n'
        '    "city": "Bangalore"\n'
        '};\n\n'
        '// Access properties using dot notation\n'
        'console.log("Name:", student.name);\n'
        'console.log("City:", student.city);'
    )
    add_output_box(
        "Name: Priya\n"
        "City: Bangalore"
    )
    add_answer_para(
        "Explanation: The properties 'name' and 'city' are accessed directly using standard JavaScript dot notation (student.name and student.city) and displayed in the console.",
        bold_prefix="Practical Note: "
    )

    # Practical Q3
    add_q_header(
        3, 2, "Convert JSON into JavaScript Object",
        "Write JavaScript code using JSON.parse() to convert the following JSON string into a JavaScript object and display the student's name: let student = '{\"name\":\"Ravi\",\"age\":24}';"
    )
    add_code_box(
        '// q3.js - Parse JSON string into JavaScript object\n'
        'let student = \'{"name":"Ravi","age":24}\';\n\n'
        '// Deserialize JSON string into a JavaScript object\n'
        'let studentObj = JSON.parse(student);\n\n'
        '// Access and display the student\'s name\n'
        'console.log("Student Name:", studentObj.name);'
    )
    add_output_box(
        "Student Name: Ravi"
    )
    add_answer_para(
        "Explanation: JSON.parse() converts the raw JSON string into a navigable JavaScript object, allowing property access via studentObj.name.",
        bold_prefix="Practical Note: "
    )

    # Practical Q4
    add_q_header(
        4, 2, "Convert JavaScript Object into JSON",
        "Create a JavaScript object containing name, age and city. Convert it into a JSON string using JSON.stringify() and display the result."
    )
    add_code_box(
        '// q4.js - Serialize JavaScript object into JSON string\n'
        'let student = {\n'
        '    name: "Rahul",\n'
        '    age: 23,\n'
        '    city: "Bangalore"\n'
        '};\n\n'
        '// Serialize object to JSON string\n'
        'let jsonString = JSON.stringify(student);\n\n'
        'console.log("JSON String:", jsonString);'
    )
    add_output_box(
        'JSON String: {"name":"Rahul","age":23,"city":"Bangalore"}'
    )
    add_answer_para(
        "Explanation: JSON.stringify() serializes the JavaScript object into a standardized JSON-encoded string, which is suitable for network transmission and API payload exchange.",
        bold_prefix="Practical Note: "
    )

    # Practical Q5
    add_q_header(
        5, 2, "Create a Simple XML Document",
        "Write an XML document to store the following student details: Name: Kumar, Age: 23, Department: CSE."
    )
    add_code_box(
        '<!-- q5.xml - Simple XML Document storing student details -->\n'
        '<?xml version="1.0" encoding="UTF-8"?>\n'
        '<student>\n'
        '    <name>Kumar</name>\n'
        '    <age>23</age>\n'
        '    <department>CSE</department>\n'
        '</student>'
    )
    add_output_box(
        "XML Structure Validated:\n"
        "Root element: <student>\n"
        "Child nodes: <name>Kumar</name>, <age>23</age>, <department>CSE</department>"
    )
    add_answer_para(
        "Explanation: XML declaration <?xml version=\"1.0\" encoding=\"UTF-8\"?> defines the version and encoding. The root tag <student> encloses hierarchical child tags for name, age, and department.",
        bold_prefix="Practical Note: "
    )

    # Practical Q6
    add_q_header(
        6, 2, "Write a Simple AJAX Fetch Request",
        "Write JavaScript code using fetch() to retrieve user details from the following API: https://jsonplaceholder.typicode.com/users/1. Display the user's name in the console."
    )
    add_code_box(
        '// q6.js - Simple AJAX Fetch request to retrieve user 1 name\n'
        'fetch("https://jsonplaceholder.typicode.com/users/1")\n'
        '    .then(response => {\n'
        '        if (!response.ok) {\n'
        '            throw new Error(`HTTP error! status: ${response.status}`);\n'
        '        }\n'
        '        return response.json();\n'
        '    })\n'
        '    .then(data => {\n'
        '        console.log("User Name:", data.name);\n'
        '    })\n'
        '    .catch(error => {\n'
        '        console.error("Error occurred:", error.message);\n'
        '    });'
    )
    add_output_box(
        "User Name: Leanne Graham"
    )
    add_answer_para(
        "Explanation: The fetch() call makes an asynchronous GET request to the REST endpoint. response.json() extracts the JSON body, and data.name retrieves the first user's name.",
        bold_prefix="Practical Note: "
    )

    # Practical Q7
    add_q_header(
        7, 2, "Fetch and Display a User's Phone Number",
        "Write JavaScript code using Fetch API to retrieve user details from: https://jsonplaceholder.typicode.com/users/3. Display the user's phone number using console.log()."
    )
    add_code_box(
        '// q7.js - Fetch user 3 and display phone number\n'
        'fetch("https://jsonplaceholder.typicode.com/users/3")\n'
        '    .then(response => {\n'
        '        if (!response.ok) {\n'
        '            throw new Error(`HTTP error! status: ${response.status}`);\n'
        '        }\n'
        '        return response.json();\n'
        '    })\n'
        '    .then(data => {\n'
        '        console.log("Phone Number:", data.phone);\n'
        '    })\n'
        '    .catch(error => {\n'
        '        console.error("Error occurred:", error.message);\n'
        '    });'
    )
    add_output_box(
        "Phone Number: 1-463-123-4447"
    )
    add_answer_para(
        "Explanation: fetch() requests user ID 3 from jsonplaceholder. After decoding the JSON response stream, data.phone logs the exact telephone number.",
        bold_prefix="Practical Note: "
    )

    # Practical Q8
    add_q_header(
        8, 2, "Display Multiple Users",
        "Write JavaScript code to retrieve all users from: https://jsonplaceholder.typicode.com/users. Use forEach() to display each user's name."
    )
    add_code_box(
        '// q8.js - Retrieve all users and iterate with forEach()\n'
        'fetch("https://jsonplaceholder.typicode.com/users")\n'
        '    .then(response => {\n'
        '        if (!response.ok) {\n'
        '            throw new Error(`HTTP error! status: ${response.status}`);\n'
        '        }\n'
        '        return response.json();\n'
        '    })\n'
        '    .then(users => {\n'
        '        console.log("List of Users:");\n'
        '        users.forEach((user, index) => {\n'
        '            console.log(`${index + 1}. ${user.name}`);\n'
        '        });\n'
        '    })\n'
        '    .catch(error => {\n'
        '        console.error("Error occurred:", error.message);\n'
        '    });'
    )
    add_output_box(
        "List of Users:\n"
        "1. Leanne Graham\n"
        "2. Ervin Howell\n"
        "3. Clementine Bauch\n"
        "4. Patricia Lebsack\n"
        "5. Chelsey Dietrich\n"
        "6. Mrs. Dennis Schulist\n"
        "7. Kurtis Weissnat\n"
        "8. Nicholas Runolfsdottir V\n"
        "9. Glenna Reichert\n"
        "10. Clementina DuBuque"
    )
    add_answer_para(
        "Explanation: The endpoint returns an array of 10 user objects. Array.prototype.forEach() iterates across each element, displaying every user's name.",
        bold_prefix="Practical Note: "
    )

    # Practical Q9
    add_q_header(
        9, 2, "Fetch and Display Company Information",
        "Write JavaScript code using Fetch API to retrieve user details from: https://jsonplaceholder.typicode.com/users/4. Display the user's company name and website in the console."
    )
    add_code_box(
        '// q9.js - Fetch user 4 and display nested company details\n'
        'fetch("https://jsonplaceholder.typicode.com/users/4")\n'
        '    .then(response => {\n'
        '        if (!response.ok) {\n'
        '            throw new Error(`HTTP error! status: ${response.status}`);\n'
        '        }\n'
        '        return response.json();\n'
        '    })\n'
        '    .then(data => {\n'
        '        console.log("Company Name:", data.company.name);\n'
        '        console.log("Website:", data.website);\n'
        '    })\n'
        '    .catch(error => {\n'
        '        console.error("Error occurred:", error.message);\n'
        '    });'
    )
    add_output_box(
        "Company Name: Robel-Corkery\n"
        "Website: kale.biz"
    )
    add_answer_para(
        "Explanation: Demonstrates accessing nested object properties within JSON: data.company.name targets the company name inside the company object, and data.website targets the top-level website property.",
        bold_prefix="Practical Note: "
    )

    # Practical Q10
    add_q_header(
        10, 2, "Handle Errors Using Fetch API",
        "Write a Fetch API program to retrieve user details from: https://jsonplaceholder.typicode.com/users/1. Use .catch() to handle errors and display an error message in the console."
    )
    add_code_box(
        '// q10.js - Fetch API program demonstrating error handling with .catch()\n'
        'fetch("https://jsonplaceholder.typicode.com/users/1")\n'
        '    .then(response => {\n'
        '        if (!response.ok) {\n'
        '            throw new Error(`HTTP error! status: ${response.status}`);\n'
        '        }\n'
        '        return response.json();\n'
        '    })\n'
        '    .then(data => {\n'
        '        console.log("Data retrieved successfully:");\n'
        '        console.log("Name:", data.name);\n'
        '        console.log("Email:", data.email);\n'
        '    })\n'
        '    .catch(error => {\n'
        '        console.error("Error occurred:", error.message);\n'
        '    });'
    )
    add_output_box(
        "Data retrieved successfully:\n"
        "Name: Leanne Graham\n"
        "Email: Sincere@april.biz\n\n"
        "// When tested with an invalid endpoint or network disconnection:\n"
        "Error occurred: HTTP error! status: 404 (or Network request failed)"
    )
    add_answer_para(
        "Explanation: The response status is checked using response.ok. If false (e.g. 404/500), an Error is thrown to immediately route execution into .catch(). In the event of network failure or DNS timeout, fetch() automatically rejects, invoking .catch() to display the error message gracefully without crashing.",
        bold_prefix="Practical Note: "
    )

    # Footer note
    doc.add_paragraph().paragraph_format.space_before = Pt(10)
    footer_p = doc.add_paragraph()
    footer_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_foot = footer_p.add_run("--- END OF ASSESSMENT (TOTAL MARKS: 30) ---")
    r_foot.font.name = "Arial"
    r_foot.font.size = Pt(9.5)
    r_foot.font.bold = True
    r_foot.font.color.rgb = MUTED

    doc_path = "/home/mithun/Documents/lnt/JavaScript-Practical-main/CIA-1_Server_Integration_Mithun_G.docx"
    doc.save(doc_path)
    print(f"Successfully generated: {doc_path}")

if __name__ == "__main__":
    create_document()
