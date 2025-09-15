import streamlit as st
import pandas as pd
import google.generativeai as genai
import io
import time
import re

# --- Configuration ---
st.set_page_config(layout="wide", page_title="Advanced AI Excel Interviewer", page_icon="🤖")

try:
    # This will be set in Streamlit Community Cloud's secrets
    genai.configure(api_key=st.secrets["GEMINI_API_KEY"])
    model = genai.GenerativeModel('gemini-1.5-flash')
except Exception as e:
    st.error(f"API Key not found or invalid. Please ensure you have set the GEMINI_API_KEY secret in your Streamlit app settings.", icon="🚨")
    st.stop()

# --- Custom CSS for Advanced UI ---
st.markdown("""
<style>
    .stApp {
        background-color: #1a253c; /* Dark blue background */
        color: #f0f2f6;
    }
    .st-chat-message {
        background-color: #2c3e50; /* Slightly lighter blue for messages */
        border-radius: 10px;
        padding: 1rem;
        box-shadow: 0 4px 8px rgba(0,0,0,0.15);
        margin-bottom: 1rem;
        border: 1px solid #34495e;
    }
    .st-chat-message.user {
        background-color: #004d40; /* Dark green for user */
    }
    .st-chat-message.assistant {
        background-color: #2c3e50;
    }
    .stButton>button {
        border-radius: 25px;
        border: 2px solid #00bfa5; /* Teal border */
        background-color: transparent;
        color: #00bfa5;
        padding: 0.5rem 1.5rem;
        font-weight: bold;
        transition: all 0.3s ease-in-out;
    }
    .stButton>button:hover {
        background-color: #00bfa5;
        color: #1a253c;
    }
    .stProgress > div > div > div > div {
        background-color: #00bfa5;
    }
    h1, h2, h3, .stMarkdown {
        color: #f0f2f6;
    }
    .st-emotion-cache-16txtl3 { /* Main content area */
        background-color: rgba(255, 255, 255, 0.05);
        padding: 2rem;
        border-radius: 15px;
    }
</style>
""", unsafe_allow_html=True)


# --- Sample Data Generation ---
@st.cache_data
def create_sample_excel():
    """Creates an in-memory Excel file with more complex sample data."""
    sales_data = {
        'Date': pd.to_datetime(['2023-01-05', '2023-01-06', '2023-01-07', '2023-01-08', '2023-02-10', '2023-02-11', '2023-03-15', '2023-03-16', '2023-03-17']),
        'Region': ['North', 'West', 'North', 'South', 'West', 'East', 'South', 'West', 'East'],
        'Category': ['Electronics', 'Apparel', 'Electronics', 'Books', 'Apparel', 'Books', 'Books', 'Electronics', 'Apparel'],
        'Product_ID': ['P101', 'P201', 'P102', 'P301', 'P202', 'P302', 'P303', 'P103', 'P201'],
        'Sales': [1200, 300, 800, 50, 200, 75, 60, 450, 150]
    }
    product_data = {
        'Product_ID': ['P101', 'P102', 'P103', 'P201', 'P202', 'P301', 'P302', 'P303'],
        'Product_Name': ['Laptop', 'Monitor', 'Keyboard', 'T-Shirt', 'Jeans', 'Novel', 'Cookbook', 'History'],
        'Unit_Cost': [800, 250, 50, 10, 25, 8, 12, 10]
    }
    df_sales = pd.DataFrame(sales_data)
    df_products = pd.DataFrame(product_data)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_sales.to_excel(writer, sheet_name='SalesData', index=False)
        df_products.to_excel(writer, sheet_name='Products', index=False)
    output.seek(0)
    return output.read()

# --- Advanced Interview Questions & Adaptive Logic ---
INTERVIEW_QUESTIONS = {
    "1": { # Unique ID for state tracking
        "difficulty": "easy",
        "type": "conceptual",
        "text": "Let's start with a conceptual question: What is the primary purpose of the `IF` function in Excel?",
        "evaluation_prompt": """
            Evaluate the user's answer about the IF function based on its core purpose: conditional logic.
            1. **Core Concept**: Does it mention making decisions or returning different values based on a condition being true or false?
            2. **Clarity**: Is the explanation clear?
            User's Answer: "{user_answer}"
            Provide a brief, one-sentence evaluation and a score out of 10.
            Format: Evaluation: [Your evaluation] | Score: [Score]/10
        """
    },
    "2": {
        "difficulty": "easy",
        "type": "practical_value",
        "text": "Using the `SalesData` sheet, what are the total sales for the 'North' region?",
        "correct_answer": 2000,
        "retries": 1
    },
    "3": {
        "difficulty": "medium",
        "type": "practical_value",
        "text": "On the `SalesData` sheet, a new column `Profit` needs to be added. Use the `Products` sheet to look up the `Unit_Cost` for each `Product_ID` and calculate the profit for each sale (`Sales` - `Unit_Cost`). What is the total profit for the 'Electronics' category?",
        "correct_answer": 1350,
        "retries": 1
    },
    "4": {
        "difficulty": "hard",
        "type": "practical_file",
        "text": "Excellent. For the final task, create a Pivot Table in a new sheet named 'Summary'. It should show the average `Sales` for each `Region`. Then, add a slicer to filter the Pivot Table by `Category`. Please upload the file once completed.",
        "evaluation_logic": "evaluate_advanced_pivot_table",
        "retries": 0
    }
}

# --- Helper Functions ---
def normalize_answer(answer):
    """Cleans and converts a string answer to a float for comparison."""
    try:
        cleaned_answer = re.sub(r'[$,\s]', '', str(answer))
        return float(cleaned_answer)
    except (ValueError, TypeError):
        return None

def evaluate_advanced_pivot_table(uploaded_file):
    try:
        xls = pd.ExcelFile(uploaded_file, engine='openpyxl')
        if 'Summary' not in xls.sheet_names:
            return False, "The 'Summary' sheet was not found. Please ensure the sheet is named correctly."
        df_summary = pd.read_excel(uploaded_file, sheet_name='Summary', engine='openpyxl')
        if 'Region' not in df_summary.columns or len(df_summary.columns) < 2:
             return False, "The structure in the 'Summary' sheet doesn't look like the requested pivot table. It should have 'Region' as a row."
        if 'Average of Sales' not in df_summary.columns and 'Avg of Sales' not in df_summary.columns:
            return False, "The pivot table seems to be calculating something other than the Average of Sales. Please check the value field settings."
        return True, "File received. The pivot table structure in the 'Summary' sheet appears correct."
    except Exception as e:
        return False, f"An error occurred while reading your file. Please ensure it's a valid .xlsx format. Error: {e}"

# --- LLM Interaction ---
def get_llm_evaluation(prompt):
    try:
        response = model.generate_content(prompt, safety_settings={'HARM_CATEGORY_HARASSMENT': 'BLOCK_NONE'})
        return response.text
    except Exception as e:
        st.error(f"Could not connect to the AI model. Error: {e}", icon="📡")
        return "Error from model."

def generate_final_report(transcript):
    prompt = f"""
        As an expert hiring manager, analyze this mock Excel interview transcript. Provide a professional, constructive feedback report with 'Strengths' and 'Areas for Improvement'.
        The transcript includes scores and retries, use them to inform your feedback. For example, if a user got a question right after a retry, mention their persistence but also the need for initial accuracy.

        Transcript:
        ---
        {transcript}
        ---
        Generate the report.
    """
    return get_llm_evaluation(prompt)

# --- UI and State Management ---
if 'stage' not in st.session_state:
    st.session_state.stage = 'intro'
    st.session_state.messages = []
    st.session_state.transcript = ""
    st.session_state.sample_excel = create_sample_excel()
    st.session_state.question_ids = sorted(INTERVIEW_QUESTIONS.keys())
    st.session_state.q_index = 0
    st.session_state.retries_left = 0
    st.session_state.score = 0
    st.session_state.max_score = len(INTERVIEW_QUESTIONS) * 10

def restart_interview():
    st.session_state.stage = 'intro'
    st.session_state.messages = []
    st.session_state.transcript = ""
    st.session_state.q_index = 0
    st.session_state.retries_left = 0
    st.session_state.score = 0
    st.rerun()

# --- Main App Layout ---
st.title("🤖 Advanced AI Excel Interviewer")
col1, col2 = st.columns([2, 1])
with col2:
    st.markdown("### 📝 Interview Details")
    with st.container(border=True):
        progress_percent = (st.session_state.q_index / len(INTERVIEW_QUESTIONS))
        st.progress(progress_percent)
        st.markdown(f"**Question:** {st.session_state.q_index + 1 if st.session_state.stage != 'complete' else st.session_state.q_index} / {len(INTERVIEW_QUESTIONS)}")
        st.metric(label="Current Score", value=f"{st.session_state.score}/{st.session_state.max_score}")
        st.download_button(
           label="📥 Download Excel File",
           data=st.session_state.sample_excel,
           file_name="InterviewData.xlsx",
           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        if st.button("Restart Interview"):
            restart_interview()
with col1:
    chat_container = st.container(height=600)
    with chat_container:
        for message in st.session_state.messages:
            with st.chat_message(message["role"]):
                st.markdown(message["content"])
    if st.session_state.stage == 'intro':
        with chat_container:
            with st.chat_message("assistant"):
                st.markdown("### Welcome to the Advanced Excel Assessment!")
                st.markdown("I'm your AI interviewer. This session will test your conceptual knowledge and practical Excel skills through a series of adaptive questions.")
                st.markdown("Please download the Excel file from the panel on the right. When you're ready, click **Start Interview**.")
        if st.button("Start Interview"):
            st.session_state.stage = 'question'
            st.rerun()
    elif st.session_state.stage == 'question':
        q_id = st.session_state.question_ids[st.session_state.q_index]
        q = INTERVIEW_QUESTIONS[q_id]
        st.session_state.retries_left = q.get("retries", 0)
        with chat_container:
            with st.chat_message("assistant"):
                st.markdown(f"**Question {st.session_state.q_index + 1} ({q['difficulty'].capitalize()})**")
                st.markdown(q["text"])
        if q["type"] in ["conceptual", "practical_value"]:
            user_answer = st.text_input("Your answer:", key=f"q_{q_id}_text")
            if st.button("Submit Answer", key=f"q_{q_id}_submit"):
                if user_answer:
                    st.session_state.messages.append({"role": "user", "content": user_answer})
                    st.session_state.transcript += f"Q ({q['difficulty']}): {q['text']}\nA: {user_answer}\n"
                    st.session_state.user_answer_submitted = user_answer
                    st.session_state.stage = 'evaluation'
                    st.rerun()
                else:
                    st.warning("Please provide an answer.")
        elif q["type"] == "practical_file":
            uploaded_file = st.file_uploader("Upload your modified Excel file:", type=["xlsx"], key=f"q_{q_id}_file")
            if uploaded_file:
                st.session_state.messages.append({"role": "user", "content": f"(Uploaded file: {uploaded_file.name})"})
                st.session_state.transcript += f"Q ({q['difficulty']}): {q['text']}\nA: (User uploaded {uploaded_file.name})\n"
                st.session_state.user_answer_submitted = uploaded_file
                st.session_state.stage = 'evaluation'
                st.rerun()
    elif st.session_state.stage == 'evaluation':
        q_id = st.session_state.question_ids[st.session_state.q_index]
        q = INTERVIEW_QUESTIONS[q_id]
        user_answer = st.session_state.user_answer_submitted
        is_correct, feedback, current_score = False, "", 0
        with st.spinner("Analyzing your response..."):
            if q["type"] == "conceptual":
                prompt = q["evaluation_prompt"].format(user_answer=user_answer)
                evaluation = get_llm_evaluation(prompt)
                feedback = f"**AI Evaluation:** {evaluation}"
                try:
                    score_str = evaluation.split("Score:")[1].strip().split("/")[0]
                    current_score = int(score_str)
                    is_correct = current_score >= 7
                except (IndexError, ValueError): pass
            elif q["type"] == "practical_value":
                normalized = normalize_answer(user_answer)
                is_correct = normalized is not None and normalized == q["correct_answer"]
                feedback = "That is correct. Well done." if is_correct else "That's not the value I was expecting."
                if is_correct: current_score = 10
            elif q["type"] == "practical_file":
                is_correct, feedback = evaluate_advanced_pivot_table(user_answer)
                if is_correct: current_score = 10
        st.session_state.transcript += f"Feedback: {feedback}\n"
        st.session_state.messages.append({"role": "assistant", "content": feedback})
        if is_correct:
            st.session_state.score += current_score
            st.session_state.transcript += f"Result: Correct (Score: {current_score}/10)\n---\n"
            st.session_state.q_index += 1
            if st.session_state.q_index >= len(INTERVIEW_QUESTIONS): st.session_state.stage = 'report'
            else: st.session_state.stage = 'question'
        else:
            if st.session_state.retries_left > 0:
                st.session_state.retries_left -= 1
                st.session_state.transcript += f"Result: Incorrect. Retrying...\n"
                st.session_state.messages.append({"role": "assistant", "content": f"Please try that again. You have {st.session_state.retries_left + 1} attempt(s) left."})
                st.session_state.stage = 'question'
            else:
                st.session_state.transcript += f"Result: Incorrect (Score: 0/10)\n---\n"
                st.session_state.q_index += 1
                if st.session_state.q_index >= len(INTERVIEW_QUESTIONS): st.session_state.stage = 'report'
                else: st.session_state.stage = 'question'
        time.sleep(1)
        st.rerun()
    elif st.session_state.stage == 'report':
        with chat_container:
            with st.chat_message("assistant"):
                st.markdown("Thank you for completing the assessment. I'm now generating your detailed performance report.")
        with st.spinner("Compiling your feedback report..."):
            final_report = generate_final_report(st.session_state.transcript)
            st.session_state.messages.append({"role": "assistant", "content": f"### Interview Performance Report\n\n{final_report}"})
            st.session_state.stage = 'complete'
            st.rerun()
    elif st.session_state.stage == 'complete':
        st.success("Assessment Complete! Your final report is in the chat. You can restart the interview using the button on the right.", icon="✅")
        with st.expander("Show Full Interview Transcript"):
            st.text_area("Transcript", st.session_state.transcript, height=400)

