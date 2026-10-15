# Word to CSV Question Converter (Streamlit)

A Streamlit web app that extracts multiple-choice questions from a `.docx` file, maps Option A–D to `option_1`–`option_4`, and maps answers to the corresponding option column.

## Run locally

1. Install Python 3.10 or newer.
2. Open a terminal in this folder.
3. Install dependencies:

   ```bash
   pip install -r requirements.txt
   ```

4. Start the app:

   ```bash
   streamlit run app.py
   ```

## Deploy to Streamlit Community Cloud

1. Create a GitHub repository and upload `app.py`, `requirements.txt`, and `README.md`.
2. Open https://share.streamlit.io/ and sign in.
3. Select **Create app**, choose the repository, branch, and `app.py`.
4. Deploy.

## Expected Word format

Use numbered questions, for example:

1. What is the capital of France?
   Option A: Berlin
   Option B: Paris
   Option C: Rome
   Option D: Madrid
   Answer: Option B

The parser supports `.docx` files. Save legacy `.doc` files as `.docx` before uploading.

## Excel template

The Excel template is optional. If supplied, its column headers and order are preserved. Expected columns include:

`question_type`, `question_group`, `question_level`, `marks`, `question`, `option_1`, `option_2`, `option_3`, `option_4`, `answer`

The sidebar lets you change the default question type, group, level, and marks.
