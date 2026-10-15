import io
import re
import csv
from pathlib import Path

import pandas as pd
import streamlit as st
from docx import Document

st.set_page_config(
    page_title="Word to CSV Question Converter",
    page_icon="📝",
    layout="wide",
)

st.title("📝 Word to CSV Question Converter")
st.write(
    "Upload a Microsoft Word question document and an Excel template. "
    "The app extracts questions, options A–D, and answers, maps the answer to "
    "`option_1`–`option_4`, then creates a downloadable CSV."
)

DEFAULTS = {
    "question_type": "single_choice",
    "question_group": "119",
    "question_level": "easy",
    "marks": "1",
}

def normalize_label(text):
    return re.sub(r"\s+", " ", text.strip().lower())

def read_docx_text(uploaded_file):
    """Read paragraphs and tables in document order as best as python-docx allows."""
    doc = Document(uploaded_file)
    lines = []
    for paragraph in doc.paragraphs:
        value = paragraph.text.strip()
        if value:
            lines.append(value)
    for table in doc.tables:
        for row in table.rows:
            cells = [cell.text.strip().replace("\n", " / ") for cell in row.cells]
            line = " | ".join(cell for cell in cells if cell)
            if line:
                lines.append(line)
    return lines

def parse_questions(lines):
    """
    Parse common layouts such as:
      1. Question text
      Option A: ...
      Option B: ...
      Option C: ...
      Option D: ...
      Answer: Option B
    Also supports question/options/answer labels on a single line.
    """
    text = "\n".join(lines)
    # Standardize Word smart punctuation and line endings.
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    text = re.sub(r"[ \t]+", " ", text)

    # A question starts with a numbered prefix (e.g. 1., 1), Question 1:).
    starts = list(re.finditer(
        r"(?im)^\s*(?:(?:question)\s*)?\d+\s*[\.\):\-]\s*",
        text
    ))
    chunks = []
    if starts:
        for i, match in enumerate(starts):
            end = starts[i + 1].start() if i + 1 < len(starts) else len(text)
            chunks.append(text[match.start():end].strip())
    else:
        # Fallback: split around explicit Question N labels.
        starts = list(re.finditer(r"(?im)^\s*question\s+\d+\s*[:\-]\s*", text))
        for i, match in enumerate(starts):
            end = starts[i + 1].start() if i + 1 < len(starts) else len(text)
            chunks.append(text[match.start():end].strip())

    parsed = []
    for chunk in chunks:
        # Remove the leading question number/label.
        q_match = re.match(
            r"(?is)^\s*(?:(?:question)\s*)?\d+\s*[\.\):\-]\s*(.*)$",
            chunk
        )
        if not q_match:
            q_match = re.match(r"(?is)^\s*question\s+\d+\s*[:\-]\s*(.*)$", chunk)
        body = q_match.group(1).strip() if q_match else chunk

        option_matches = list(re.finditer(
            r"(?i)\boption\s*([A-D])\s*[:.)\-]\s*",
            body
        ))
        answer_match = re.search(
            r"(?i)\banswer\s*[:\-]\s*(?:option\s*)?([A-D])\b",
            body
        )

        if not option_matches:
            continue

        question = body[:option_matches[0].start()].strip(" \n:-")
        question = re.sub(r"(?i)^\s*question\s*[:\-]\s*", "", question).strip()

        options = {}
        for idx, match in enumerate(option_matches):
            next_start = option_matches[idx + 1].start() if idx + 1 < len(option_matches) else len(body)
            segment = body[match.end():next_start]
            # The answer label is not part of the last option.
            segment = re.split(r"(?i)\banswer\s*[:\-]", segment)[0]
            options[match.group(1).upper()] = segment.strip(" \n;|")

        if not answer_match:
            # Handle answer on a separate line that may have been split/normalized.
            answer_match = re.search(r"(?i)\banswer\s*[:\-]\s*(?:option\s*)?([A-D])\b", body)

        answer_letter = answer_match.group(1).upper() if answer_match else ""
        row = {
            "question_type": DEFAULTS["question_type"],
            "question_group": DEFAULTS["question_group"],
            "question_level": DEFAULTS["question_level"],
            "marks": DEFAULTS["marks"],
            "question": question,
            "option_1": options.get("A", ""),
            "option_2": options.get("B", ""),
            "option_3": options.get("C", ""),
            "option_4": options.get("D", ""),
            "answer": f"option_{ord(answer_letter) - ord('A') + 1}" if answer_letter else "",
        }
        row["_answer_letter"] = answer_letter
        parsed.append(row)
    return parsed

def find_header_columns(df):
    return {str(col).strip().lower(): col for col in df.columns}

def make_output(parsed_rows, template_file):
    if template_file is not None:
        template_file.seek(0)
        template = pd.read_excel(template_file)
        headers = list(template.columns)
    else:
        headers = [
            "question_type", "question_group", "question_level", "marks",
            "question", "option_1", "option_2", "option_3", "option_4", "answer"
        ]

    output_rows = []
    for item in parsed_rows:
        clean = {k: v for k, v in item.items() if not k.startswith("_")}
        # Match headers case-insensitively; preserve template's exact header names/order.
        normalized = {str(h).strip().lower(): h for h in headers}
        out = {}
        for key, value in clean.items():
            if key.lower() in normalized:
                out[normalized[key.lower()]] = value
        output_rows.append(out)

    # Include any template columns that aren't populated as blank.
    return pd.DataFrame(output_rows, columns=headers)

with st.sidebar:
    st.header("Default values")
    q_type = st.text_input("question_type", "single_choice")
    q_group = st.text_input("question_group", "119")
    q_level = st.text_input("question_level", "easy")
    marks = st.text_input("marks", "1")
    st.caption("These values are applied to every imported question.")

st.subheader("1. Upload your files")
word_file = st.file_uploader("Word document with questions (.docx)", type=["docx"])
template_file = st.file_uploader(
    "Excel template (.xlsx, optional)", type=["xlsx"],
    help="If supplied, the CSV will use the template's exact column names and order."
)

if word_file:
    if not word_file.name.lower().endswith(".docx"):
        st.error("Please upload a .docx file. Legacy .doc files need to be saved as .docx first.")
    else:
        try:
            lines = read_docx_text(word_file)
            parsed = parse_questions(lines)

            # Apply values currently set in the sidebar.
            for row in parsed:
                row["question_type"] = q_type
                row["question_group"] = q_group
                row["question_level"] = q_level
                row["marks"] = marks

            st.subheader("2. Review extracted questions")
            if not parsed:
                st.warning(
                    "No questions were detected. Please check the document layout. "
                    "For best results, use numbered questions and labels like "
                    "'Option A: ...' through 'Option D: ...', followed by 'Answer: Option A'."
                )
                with st.expander("Show extracted document text"):
                    st.text("\n".join(lines))
            else:
                review_df = pd.DataFrame([{k: v for k, v in row.items() if not k.startswith("_")} for row in parsed])
                missing_answers = sum(1 for row in parsed if not row.get("answer"))
                missing_options = sum(
                    1 for row in parsed
                    if any(not row.get(f"option_{i}") for i in range(1, 5))
                )
                c1, c2, c3 = st.columns(3)
                c1.metric("Questions detected", len(parsed))
                c2.metric("Missing answers", missing_answers)
                c3.metric("Questions with missing options", missing_options)

                edited_df = st.data_editor(
                    review_df,
                    use_container_width=True,
                    num_rows="dynamic",
                    key="review_editor",
                )
                st.caption("Review the extracted text and correct any parsing issues before downloading.")

                st.subheader("3. Download CSV")
                csv_bytes = edited_df.to_csv(index=False, quoting=csv.QUOTE_MINIMAL).encode("utf-8-sig")
                st.download_button(
                    label="⬇️ Download converted CSV",
                    data=csv_bytes,
                    file_name="converted_questions.csv",
                    mime="text/csv",
                    type="primary",
                )
                st.download_button(
                    label="Download review as Excel",
                    data=edited_df.to_csv(index=False).encode("utf-8-sig"),
                    file_name="converted_questions_review.csv",
                    mime="text/csv",
                )
        except Exception as exc:
            st.error(f"Could not process the document: {exc}")
            st.info("Try opening the document in Microsoft Word and saving it as a .docx file.")

st.divider()
st.caption(
    "Privacy note: processing happens in this Streamlit app session. "
    "If you deploy it publicly, avoid uploading confidential or sensitive question banks."
)
