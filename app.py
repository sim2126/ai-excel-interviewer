import streamlit as st
import pandas as pd
import google.generativeai as genai
import io
import time
import re

# --- Page Configuration ---
st.set_page_config(layout="wide", page_title="AI Excel Assessment Suite", page_icon="💼")

# --- API Configuration ---
try:
    genai.configure(api_key=st.secrets["GEMINI_API_KEY"])
    model = genai.GenerativeModel('gemini-1.5-flash')
except Exception as e:
    st.error("API Key is not configured correctly. Please add it to your Streamlit secrets.", icon="🚨")
    st.stop()

st.markdown("""
<style>
    /* Main app styling with gradient background */
    .stApp {
        background-color: #0d1117; /* GitHub Dark BG */
        color: #c9d1d9; /* GitHub Dark Text */
    }
    /* Main content container styling */
    .st-emotion-cache-16txtl3 {
        background-color: #161b22; /* GitHub Dark Paper */
        border: 1px solid #30363d; /* GitHub Dark Border */
        border-radius: 12px;
        padding: 2rem;
    }
    /* Chat message styling */
    .stChatMessage {
        background-color: #21262d; /* GitHub Dark Component BG */
        border: 1px solid #30363d;
        border-radius: 10px;
        box-shadow: 0 4px 6px rgba(0,0,0,0.3);
    }
    /* Main action button styling */
    .stButton>button {
        border-radius: 8px;
        border: 1px solid #8b5cf6; /* Violet 500 */
        background-image: linear-gradient(to right, #7c3aed, #a78bfa); /* Violet 600 -> 400 */
        color: white;
        font-weight: bold;
        transition: transform 0.2s, box-shadow 0.2s;
    }
    .stButton>button:hover {
        transform: scale(1.03);
        box-shadow: 0 0 15px #a78bfa;
        color: white;
        border: 1px solid #c4b5fd;
    }
    /* Secondary button styling (e.g., Restart) */
    .st-emotion-cache-7ym5gk button {
        border-color: #4b5563; /* Gray 600 */
        background-color: #374151; /* Gray 700 */
    }
    /* Title and header styling */
    h1, h2 {
        color: #f0f6fc; /* GitHub Dark Heading */
        border-bottom: 2px solid #8b5cf6;
        padding-bottom: 0.3rem;
    }
    /* Progress bar styling */
    .stProgress > div > div > div > div {
        background-image: linear-gradient(to right, #7c3aed, #a78bfa);
    }
    /* Metric label styling */
    .st-emotion-cache-1g8m2i4 {
        color: #8b949e; /* GitHub Dark Secondary Text */
    }
</style>
""", unsafe_allow_html=True)

@st.cache_data
def create_enhanced_excel():
    employee_sales = {
        'EmployeeID': ['E101', 'E102', 'E103', 'E101', 'E104', 'E102', 'E105', 'E103', 'E104', 'E105'],
        'SaleDate': pd.to_datetime(['2023-04-10', '2023-04-12', '2023-04-15', '2023-05-02', '2023-05-05', '2023-05-08', '2023-06-11', '2023-06-14', '2023-06-18', '2023-06-20']),
        'ProductID': ['P202', 'P301', 'P101', 'P203', 'P401', 'P302', 'P102', 'P202', 'P402', 'P101'],
        'UnitsSold': [5, 20, 2, 8, 3, 15, 4, 7, 2, 3],
        'SaleValue': [1250, 400, 4000, 2000, 6000, 300, 800, 1750, 4000, 6000]
    }
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
        "hint": "Think about the '$' symbol and how it affects formulas when you drag or copy them across cells.",
        "evaluation_prompt": "..." # Prompt remains the same
    },
    "2": {
        "difficulty": "Easy",
        "type": "practical_value",
        "text": "Using the provided Excel file, what is the **total `SaleValue`** from all sales recorded in the `Sales` sheet?",
        "hint": "Look for a function that adds up all numbers in a range. The range you need is the entire `SaleValue` column.",
        "correct_answer": 26500,
        "retries": 1
    },
    "3": {
        "difficulty": "Medium",
        "type": "practical_value",
        "text": "Using `VLOOKUP` or `XLOOKUP`, find the `Category` for `ProductID` **P401**. What is it?",
        "hint": "This function needs a lookup value (P401), a table to search in (the Products sheet), and the column number to return the result from.",
        "correct_answer": "Service",
        "retries": 1
    },
    "4": {
        "difficulty": "Medium",
        "type": "practical_value",
        "text": "Calculate the total number of **unique** employees who made a sale. How many are there?",
        "hint": "Excel has functions to count unique values. You might need to combine `COUNT` with a function that identifies unique items, or use a PivotTable.",
        "correct_answer": 5,
        "retries": 1
    },
    "5": {
        "difficulty": "Hard",
        "type": "practical_value",
        "text": "Calculate the total `SaleValue` specifically for the **'Hardware'** category. This will require you to combine data from both sheets.",
        "hint": "Consider using `SUMIF` or `SUMIFS`. You'll need to define a criteria range (the product categories) and a sum range (the sale values).",
        "correct_answer": 10800,
        "retries": 1
    },
    "6": {
        "difficulty": "Hard",
        "type": "practical_file",
        "text": "For the final task, please modify the Excel file. In the `Sales` sheet, add a new column named `Profit`. Calculate the profit for each sale (`SaleValue` - (`UnitsSold` * `CostPerUnit`)). Then, use **Conditional Formatting** to highlight all `Profit` values **greater than $2000** with a green fill. Upload the modified file.",
        "hint": "To get the `CostPerUnit` for the profit formula, you'll need to use a lookup function within the `Sales` sheet that pulls data from the `Products` sheet.",
        "evaluation_logic": "evaluate_profit_and_formatting",
        "retries": 0
    }
}
INTERVIEW_QUESTIONS["1"]["evaluation_prompt"] = """
            Evaluate the user's answer on relative vs. absolute references.
            - **Relative Reference (1 pt):** Mentions that it changes when a formula is copied.
            - **Absolute Reference (1 pt):** Mentions that it remains constant (using '$').
            - **Use Case (1 pt):** Provides a valid example like a fixed tax rate or a lookup value.
            Score the answer out of 10 based on clarity and correctness (e.g., all 3 points = 10/10).
            User's Answer: "{user_answer}"
            Format: Evaluation: [Brief evaluation] | Score: [Score]/10
        """


def normalize_answer(answer, expected_type):
    """Cleans and converts answer for case-insensitive comparison."""
    try:
        clean_str = str(answer).strip().lower() # .lower() ensures case-insensitivity
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

# --- LLM Interaction ---
def get_llm_response(prompt):
    try:
        response = model.generate_content(prompt, safety_settings={'HARM_CATEGORY_HARASSMENT': 'BLOCK_NONE'})
        return response.text
    except Exception as e:
        st.error(f"AI model communication error: {e}", icon="📡")
        return "Error: Could not get a response from the AI model."

def generate_final_report(transcript):
    """Generates a professional feedback report, now considering hints and skips."""
    prompt = f"""
        Act as a Senior Technical Recruiter specializing in finance and data analytics roles.
        Analyze the following Excel mock interview transcript and generate a detailed, professional feedback report.
        Pay attention to whether the candidate used hints or skipped questions.
        The report should be structured with:
        1.  **Overall Summary:** A brief, encouraging paragraph summarizing the candidate's performance.
        2.  **Key Strengths:** 2-3 bullet points highlighting what the candidate did well.
        3.  **Areas for Development:** 2-3 constructive bullet points. Mention if hints were needed or questions were skipped as an area for improvement.
        4.  **Final Recommendation:** A concluding sentence about their readiness for an Excel-intensive role.

        Transcript:
        ---
        {transcript}
        ---
    """
    return get_llm_response(prompt)


# --- Session State Management ---
def initialize_session():
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
        st.session_state.hint_used = [] # Track which questions a hint was used for

def restart_interview():
    for key in list(st.session_state.keys()):
        del st.session_state[key]
    st.rerun()

initialize_session()

# --- Main App UI ---
st.title("💼 AI Excel Assessment Suite")

with st.sidebar:
    st.header("Control Panel")
    st.download_button(
       label="Download Assessment File",
       data=st.session_state.sample_excel,
       file_name="EnterpriseAssessment.xlsx",
       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
       use_container_width=True
    )
    if st.button("Restart Interview", use_container_width=True, type="secondary"):
        restart_interview()

    st.divider()
    st.header("Progress")
    progress_percent = (st.session_state.q_index / len(INTERVIEW_QUESTIONS))
    st.progress(progress_percent, text=f"Question {st.session_state.q_index + 1 if st.session_state.stage != 'complete' else len(INTERVIEW_QUESTIONS)} of {len(INTERVIEW_QUESTIONS)}")
    st.metric(label="Current Score", value=f"{st.session_state.score} / {st.session_state.max_score}")

chat_container = st.container(height=600)
for message in st.session_state.messages:
    with chat_container.chat_message(message["role"]):
        st.markdown(message["content"])

# --- Application Logic Flow ---
if st.session_state.stage == 'intro':
    with chat_container.chat_message("assistant"):
        st.markdown("### Welcome to the Enterprise Excel Assessment!")
        st.markdown("This session will evaluate your practical and conceptual Excel abilities. Download the assessment file from the sidebar. When ready, click **Start Assessment**.")
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

    # Action buttons (Hint/Skip)
    action_cols = st.columns([1, 1, 4]) 
    with action_cols[0]:
        if q_id not in st.session_state.hint_used:
            if st.button("Get a Hint💡"):
                st.session_state.hint_used.append(q_id)
                st.session_state.transcript += f"Hint Used for Q{q_id}.\n"
                with chat_container.chat_message("assistant"):
                    st.info(f"**Hint:** {q['hint']}")
                st.session_state.messages.append({"role": "assistant", "content": f"**Hint:** {q['hint']}"})
    with action_cols[1]:
        if st.button("Skip Question ➡️"):
            st.session_state.transcript += f"Q: {q['text']}\nA: (Question Skipped)\nFeedback: Skipped by user.\nResult: Incorrect (Score: 0/10)\n---\n"
            st.session_state.q_index += 1
            if st.session_state.q_index >= len(INTERVIEW_QUESTIONS):
                st.session_state.stage = 'report'
            st.rerun()

    # Answer input
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
                score, is_correct = 0, False
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
            st.session_state.transcript += "Result: Incorrect. Retrying...\n"
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

