import streamlit as st
import pandas as pd
import google.generativeai as genai
import io
import time
import re

st.set_page_config(layout="wide", page_title="AI-Powered Excel Assessment", page_icon="📊")

try:
    genai.configure(api_key=st.secrets["GEMINI_API_KEY"])
    model = genai.GenerativeModel('gemini-1.5-flash')
except Exception as e:
    st.error("API Key is not configured correctly. Please add it to your Streamlit secrets.", icon="🚨")
    st.stop()

st.markdown("""
<style>
    /* Main app styling */
    .stApp {
        background-color: #0f172a; /* Slate 900 */
        color: #cbd5e1; /* Slate 300 */
    }
    /* Main content container */
    .st-emotion-cache-16txtl3 {
        background-color: #1e293b; /* Slate 800 */
        border: 1px solid #334155; /* Slate 700 */
        border-radius: 12px;
        padding: 2rem;
    }
    /* Chat message styling */
    .stChatMessage {
        background-color: #334155; /* Slate 700 */
        border-radius: 10px;
        box-shadow: 0 4px 6px rgba(0,0,0,0.2);
    }
    /* Button styling for a modern look */
    .stButton>button {
        border-radius: 8px;
        border: 1px solid #22d3ee; /* Cyan 400 */
        background-image: linear-gradient(to right, #06b6d4, #22d3ee);
        color: white;
        font-weight: bold;
        transition: transform 0.2s, box-shadow 0.2s;
    }
    .stButton>button:hover {
        transform: scale(1.03);
        box-shadow: 0 0 15px #22d3ee;
        color: white;
        border: 1px solid #67e8f9;
    }
    /* Title and header styling */
    h1, h2 {
        color: #f8fafc; /* Slate 50 */
        border-bottom: 2px solid #22d3ee;
        padding-bottom: 0.3rem;
    }
    /* Progress bar styling */
    .stProgress > div > div > div > div {
        background-image: linear-gradient(to right, #06b6d4, #22d3ee);
    }
    /* Metric label styling */
    .st-emotion-cache-1g8m2i4 {
        color: #94a3b8; /* Slate 400 */
    }
</style>
""", unsafe_allow_html=True)

@st.cache_data
def create_enhanced_excel():
    """Generates a more complex in-memory Excel file for advanced questions."""
    # Data for the first sheet: Employee Sales
    employee_sales = {
        'EmployeeID': ['E101', 'E102', 'E103', 'E101', 'E104', 'E102', 'E105', 'E103', 'E104', 'E105'],
        'SaleDate': pd.to_datetime(['2023-04-10', '2023-04-12', '2023-04-15', '2023-05-02', '2023-05-05', '2023-05-08', '2023-06-11', '2023-06-14', '2023-06-18', '2023-06-20']),
        'ProductID': ['P202', 'P301', 'P101', 'P203', 'P401', 'P302', 'P102', 'P202', 'P402', 'P101'],
        'UnitsSold': [5, 20, 2, 8, 3, 15, 4, 7, 2, 3],
        'SaleValue': [1250, 400, 4000, 2000, 6000, 300, 800, 1750, 4000, 6000]
    }
    # Data for the second sheet: Product Details
    product_details = {
        'ProductID': ['P101', 'P102', 'P202', 'P203', 'P301', 'P302', 'P401', 'P402'],
        'Category': ['Hardware', 'Hardware', 'Software', 'Software', 'Accessory', 'Accessory', 'Service', 'Service'],
        'ProductName': ['Laptop Pro', 'Monitor HD', 'OS License', 'Antivirus', 'Wireless Mouse', 'Keyboard', 'Support Plan', 'Cloud Storage'],
        'CostPerUnit': [1500, 180, 200, 40, 15, 25, 1800, 1500]
    }
    df_sales = pd.DataFrame(employee_sales)
    df_products = pd.DataFrame(product_details)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_sales.to_excel(writer, sheet_name='Sales', index=False)
        df_products.to_excel(writer, sheet_name='Products', index=False)
    output.seek(0)
    return output.read()

INTERVIEW_QUESTIONS = {
    "1": {
        "difficulty": "Easy",
        "type": "conceptual",
        "text": "Let's begin. What is the difference between a **relative** and an **absolute** cell reference in Excel, and when would you use an absolute reference?",
        "evaluation_prompt": """
            Evaluate the user's answer on relative vs. absolute references.
            - **Relative Reference (1 pt):** Mentions that it changes when a formula is copied.
            - **Absolute Reference (1 pt):** Mentions that it remains constant (using '$').
            - **Use Case (1 pt):** Provides a valid example like a fixed tax rate or a lookup value.
            Score the answer out of 10 based on clarity and correctness (e.g., all 3 points = 10/10).
            User's Answer: "{user_answer}"
            Format: Evaluation: [Brief evaluation] | Score: [Score]/10
        """
    },
    "2": {
        "difficulty": "Easy",
        "type": "practical_value",
        "text": "Using the provided Excel file, what is the **total `SaleValue`** from all sales recorded in the `Sales` sheet?",
        "correct_answer": 26500,
        "retries": 1
    },
    "3": {
        "difficulty": "Medium",
        "type": "practical_value",
        "text": "Using `VLOOKUP` or `XLOOKUP`, find the `Category` for `ProductID` **P401**. What is it?",
        "correct_answer": "Service", # Case-insensitive check
        "retries": 1
    },
    "4": {
        "difficulty": "Medium",
        "type": "practical_value",
        "text": "Calculate the total number of **unique** employees who made a sale. How many are there?",
        "correct_answer": 5,
        "retries": 1
    },
    "5": {
        "difficulty": "Hard",
        "type": "practical_value",
        "text": "Calculate the total `SaleValue` specifically for the **'Hardware'** category. This will require you to combine data from both sheets.",
        "correct_answer": 10800,
        "retries": 1
    },
    "6": {
        "difficulty": "Hard",
        "type": "practical_file",
        "text": "For the final task, please modify the Excel file. In the `Sales` sheet, add a new column named `Profit`. Calculate the profit for each sale (`SaleValue` - (`UnitsSold` * `CostPerUnit`)). Then, use **Conditional Formatting** to highlight all `Profit` values **greater than $2000** with a green fill. Upload the modified file.",
        "evaluation_logic": "evaluate_profit_and_formatting",
        "retries": 0
    }
}


def normalize_answer(answer, expected_type):
    """Cleans and converts answer for comparison."""
    try:
        clean_str = str(answer).strip().lower()
        if expected_type == 'numeric':
            return float(re.sub(r'[$,\s]', '', clean_str))
        elif expected_type == 'text':
            return clean_str
    except (ValueError, TypeError):
        return None

def evaluate_profit_and_formatting(uploaded_file):
    """Advanced evaluation for the final practical task."""
    try:
        xls = pd.ExcelFile(uploaded_file, engine='openpyxl')
        if 'Sales' not in xls.sheet_names:
            return False, "The 'Sales' sheet is missing from the uploaded file."
        df_sales = pd.read_excel(uploaded_file, sheet_name='Sales', engine='openpyxl')
        if 'Profit' not in df_sales.columns:
            return False, "The 'Profit' column was not found in the 'Sales' sheet."
        if not any(df_sales['Profit'].round() == 250):
             return False, "The profit calculation seems incorrect. Please double-check your formula."
        return True, "File uploaded. The 'Profit' column and its calculation appear correct. The conditional formatting will be reviewed manually."
    except Exception as e:
        return False, f"Could not process the uploaded file. Please ensure it's a valid .xlsx file. Error: {e}"

def get_llm_response(prompt):
    """Gets a response from the Gemini model with robust error handling."""
    try:
        response = model.generate_content(prompt, safety_settings={'HARM_CATEGORY_HARASSMENT': 'BLOCK_NONE'})
        return response.text
    except Exception as e:
        st.error(f"AI model communication error: {e}", icon="📡")
        return "Error: Could not get a response from the AI model."

def generate_final_report(transcript):
    """Generates a professional feedback report based on the interview transcript."""
    prompt = f"""
        Act as a Senior Technical Recruiter specializing in finance and data analytics roles.
        Analyze the following Excel mock interview transcript and generate a detailed, professional feedback report.
        The report should be structured with:
        1.  **Overall Summary:** A brief, encouraging paragraph summarizing the candidate's performance.
        2.  **Key Strengths:** 2-3 bullet points highlighting what the candidate did well. Use their scores and correct answers as evidence.
        3.  **Areas for Development:** 2-3 constructive bullet points on where they struggled. Mention retries or incorrect answers and suggest specific Excel functions or concepts to review.
        4.  **Final Recommendation:** A concluding sentence about their readiness for an Excel-intensive role.

        Transcript:
        ---
        {transcript}
        ---
    """
    return get_llm_response(prompt)


# --- Session State Management ---
def initialize_session():
    """Sets up the initial state for the interview session."""
    if 'stage' not in st.session_state:
        st.session_state.stage = 'intro'
        st.session_state.messages = []
        st.session_state.transcript = ""
        st.session_state.sample_excel = create_enhanced_excel()
        st.session_state.question_ids = sorted(INTERVIEW_QUESTIONS.keys())
        st.session_state.q_index = 0
        st.session_state.retries_left = 0
        st.session_state.score = 0
        st.session_state.max_score = len(INTERVIEW_QUESTIONS) * 10

def restart_interview():
    """Resets the session state to start a new interview."""
    for key in list(st.session_state.keys()):
        del st.session_state[key]
    st.rerun()

initialize_session()

st.title("📊 AI-Powered Excel Assessment")

with st.sidebar:
    st.header("Interview Control Panel")
    st.download_button(
       label="Download Assessment File",
       data=st.session_state.sample_excel,
       file_name="ProfessionalAssessment.xlsx",
       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
       use_container_width=True
    )
    if st.button("Restart Interview", use_container_width=True, type="secondary"):
        restart_interview()

    st.divider()
    st.header("Progress")
    progress_percent = (st.session_state.q_index / len(INTERVIEW_QUESTIONS))
    st.progress(progress_percent, text=f"Question {st.session_state.q_index + 1 if st.session_state.stage != 'complete' else st.session_state.q_index} of {len(INTERVIEW_QUESTIONS)}")
    st.metric(label="Current Score", value=f"{st.session_state.score} / {st.session_state.max_score}")

chat_container = st.container(height=600)
for message in st.session_state.messages:
    with chat_container.chat_message(message["role"]):
        st.markdown(message["content"])

if st.session_state.stage == 'intro':
    with chat_container.chat_message("assistant"):
        st.markdown("### Welcome to the Professional Excel Assessment!")
        st.markdown("This session is designed to evaluate your practical and conceptual Excel abilities. You will be presented with questions of varying difficulty.")
        st.markdown("Please download the assessment file from the sidebar. When you are ready to begin, click **Start Assessment**.")
    if st.button("Start Assessment"):
        st.session_state.stage = 'question'
        st.rerun()

elif st.session_state.stage == 'question':
    q_id = st.session_state.question_ids[st.session_state.q_index]
    q = INTERVIEW_QUESTIONS[q_id]
    st.session_state.retries_left = q.get("retries", 0)

    with chat_container.chat_message("assistant"):
        st.markdown(f"**Question {st.session_state.q_index + 1}: {q['difficulty']}**")
        st.markdown(q["text"])

    if q["type"] in ["conceptual", "practical_value"]:
        user_answer = st.chat_input("Enter your answer here...", key=f"q_{q_id}_text")
        if user_answer:
            st.session_state.messages.append({"role": "user", "content": user_answer})
            st.session_state.transcript += f"Q: {q['text']}\nA: {user_answer}\n"
            st.session_state.user_answer_submitted = user_answer
            st.session_state.stage = 'evaluation'
            st.rerun()
    elif q["type"] == "practical_file":
        with st.form(key=f"q_{q_id}_form", clear_on_submit=True):
            uploaded_file = st.file_uploader("Upload your completed Excel file:", type=["xlsx"])
            submitted = st.form_submit_button("Upload and Submit")
            if submitted and uploaded_file is not None:
                st.session_state.messages.append({"role": "user", "content": f"(File Uploaded: {uploaded_file.name})"})
                st.session_state.transcript += f"Q: {q['text']}\nA: (User uploaded {uploaded_file.name})\n"
                st.session_state.user_answer_submitted = uploaded_file
                st.session_state.stage = 'evaluation'
                st.rerun()

elif st.session_state.stage == 'evaluation':
    q_id = st.session_state.question_ids[st.session_state.q_index]
    q = INTERVIEW_QUESTIONS[q_id]
    user_answer = st.session_state.user_answer_submitted
    is_correct, feedback, score = False, "", 0

    with st.spinner("Analyzing response..."):
        if q["type"] == "conceptual":
            prompt = q["evaluation_prompt"].format(user_answer=user_answer)
            evaluation = get_llm_response(prompt)
            feedback = f"**AI Evaluation:** {evaluation}"
            try:
                score = int(re.search(r'Score:\s*(\d+)/10', evaluation).group(1))
                is_correct = score >= 7
            except (AttributeError, ValueError):
                score = 0
                is_correct = False
        elif q["type"] == "practical_value":
            expected_type = 'numeric' if isinstance(q['correct_answer'], (int, float)) else 'text'
            norm_user_ans = normalize_answer(user_answer, expected_type)
            norm_correct_ans = normalize_answer(q['correct_answer'], expected_type)
            is_correct = norm_user_ans == norm_correct_ans
            feedback = "That is correct. Excellent." if is_correct else "That is not the expected answer."
            if is_correct: score = 10
        elif q["type"] == "practical_file":
            is_correct, feedback = evaluate_profit_and_formatting(user_answer)
            if is_correct: score = 10

    st.session_state.transcript += f"Feedback: {feedback}\n"
    with chat_container.chat_message("assistant"):
        st.markdown(feedback)

    if is_correct:
        st.session_state.score += score
        st.session_state.transcript += f"Result: Correct (Score: {score}/10)\n---\n"
    else:
        if st.session_state.retries_left > 0:
            st.session_state.retries_left -= 1
            st.session_state.transcript += f"Result: Incorrect. Retrying...\n"
            with chat_container.chat_message("assistant"):
                st.warning(f"Please try that again. You have {st.session_state.retries_left + 1} attempt(s) remaining.")
            st.session_state.stage = 'question'
        else:
            st.session_state.transcript += f"Result: Incorrect (Score: 0/10)\n---\n"

    if is_correct or st.session_state.retries_left <= 0:
        st.session_state.q_index += 1
        if st.session_state.q_index >= len(INTERVIEW_QUESTIONS):
            st.session_state.stage = 'report'
        else:
            st.session_state.stage = 'question'

    time.sleep(1.5) 
    st.rerun()

elif st.session_state.stage == 'report':
    with chat_container.chat_message("assistant"):
        with st.spinner("Compiling your detailed performance report..."):
            final_report = generate_final_report(st.session_state.transcript)
            st.session_state.final_report = final_report
        st.markdown("### Your Assessment is Complete")
        st.markdown("Thank you for your time. Your detailed performance report is now available below.")
    st.session_state.stage = 'complete'
    st.rerun()

elif st.session_state.stage == 'complete':
    st.success("Assessment Finished! Review your report below.", icon="✅")
    with st.container(border=True):
        st.markdown(st.session_state.final_report)
    with st.expander("Show Full Interview Transcript"):
        st.text_area("Transcript", st.session_state.transcript, height=400)

