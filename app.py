import streamlit as st
from docx import Document
import pandas as pd
import re
from io import BytesIO

st.set_page_config(page_title="Word → iSpring QuizMaker", layout="wide")
st.title("📄 Word → iSpring QuizMaker (Tự động bỏ A.B.C.D & nhận gạch dưới)")
st.markdown("**Upload file Word → Kiểm tra & sửa bảng → Tải Excel**")

# ====================== HÀM ĐỌC WORD (NÂNG CAO) ======================
def parse_word_file(docx_file):
    doc = Document(docx_file)
    questions = []
    current_q = None
    current_options = []

    for paragraph in doc.paragraphs:
        text = paragraph.text.strip()
        if not text:
            continue

        # Phát hiện câu hỏi mới
        if re.match(r'^Câu\s+\d+[:.]', text, re.IGNORECASE):
            if current_q and current_options:
                questions.append({"question": current_q, "options": current_options[:]})
            current_q = text
            current_options = []
            continue

        # Xử lý lựa chọn A. B. C. D. + kiểm tra gạch dưới
        if re.match(r'^[A-D]\.', text):
            # Kiểm tra gạch dưới (underline)
            is_correct = False
            for run in paragraph.runs:
                if run.text.strip().startswith(('A.', 'B.', 'C.', 'D.')):
                    if run.underline:           # True nếu bị gạch dưới
                        is_correct = True
                    break

            # Bỏ "A. ", "B. "...
            content = re.sub(r'^[A-D]\.\s*', '', text).strip()
            if is_correct:
                content = "*" + content

            current_options.append(content)
        # Nội dung bổ sung cho câu hỏi
        elif current_q and not current_options:
            current_q += "\n" + text

    # Thêm câu hỏi cuối cùng
    if current_q and current_options:
        questions.append({"question": current_q, "options": current_options[:]})

    # Tạo DataFrame theo format iSpring
    data = []
    for q in questions:
        row = {
            "Question Type": "MC",
            "Question Text": q["question"],
            "Image": "", "Video": "", "Audio": "",
            "Answer 1": q["options"][0] if len(q["options"]) > 0 else "",
            "Answer 2": q["options"][1] if len(q["options"]) > 1 else "",
            "Answer 3": q["options"][2] if len(q["options"]) > 2 else "",
            "Answer 4": q["options"][3] if len(q["options"]) > 3 else "",
            "Answer 5": "", "Answer 6": "", "Answer 7": "", "Answer 8": "", "Answer 9": "", "Answer 10": "",
            "Correct Feedback": "Chúc mừng bạn! Đáp án đúng.",
            "Incorrect Feedback": "Rất tiếc, đáp án chưa chính xác.",
            "Points": 1
        }
        data.append(row)

    return pd.DataFrame(data)


# ====================== GIAO DIỆN ======================
uploaded_file = st.file_uploader("📤 Chọn file Word (.docx)", type=["docx"])

if uploaded_file:
    st.success(f"✅ Đã tải: {uploaded_file.name}")
    
    df = parse_word_file(uploaded_file)
    
    st.subheader("📋 Bảng câu hỏi (click vào ô để sửa)")
    st.info("✅ Đáp án nào có * là đáp án đúng (đã tự động nhận từ gạch dưới trong Word). Bạn vẫn có thể sửa tay nếu cần.")

    edited_df = st.data_editor(
        df,
        use_container_width=True,
        hide_index=True,
        num_rows="dynamic",
        column_config={
            "Question Text": st.column_config.TextColumn("Câu hỏi", width="large"),
            "Answer 1": st.column_config.TextColumn("Answer 1", width="medium"),
            "Answer 2": st.column_config.TextColumn("Answer 2", width="medium"),
            "Answer 3": st.column_config.TextColumn("Answer 3", width="medium"),
            "Answer 4": st.column_config.TextColumn("Answer 4", width="medium"),
        }
    )

    if st.button("💾 Tải file Excel cho iSpring", type="primary", use_container_width=True):
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            edited_df.to_excel(writer, index=False, sheet_name="Sample")
        output.seek(0)

        st.download_button(
            label="📥 TẢI NGAY FILE iSpring_Questions_Import.xlsx",
            data=output,
            file_name="iSpring_Questions_Import.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
        st.success("✅ File Excel đã sẵn sàng import vào iSpring QuizMaker!")

st.caption("Ứng dụng chạy cục bộ • Streamlit + python-docx")