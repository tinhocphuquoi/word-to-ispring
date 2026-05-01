import streamlit as st
from docx import Document
import pandas as pd
import re
from io import BytesIO

st.set_page_config(page_title="Word → iSpring QuizMaker", layout="wide")
st.title("📄 Word → iSpring QuizMaker (Đã fix lỗi lưu thay đổi)")
st.markdown("**Upload file Word → Chỉnh sửa bảng → Tải Excel (thay đổi được lưu)**")

# ====================== HÀM ĐỌC WORD ======================
def parse_word_file(docx_file):
    doc = Document(docx_file)
    questions = []
    current_q = None
    current_options = []

    for paragraph in doc.paragraphs:
        text = paragraph.text.strip()
        if not text:
            continue

        if re.match(r'^Câu\s+\d+[:.]', text, re.IGNORECASE):
            if current_q and current_options:
                questions.append({"question": current_q, "options": current_options[:]})
            current_q = text
            current_options = []
            continue

        if re.match(r'^[A-D]\.', text):
            is_correct = any(run.underline for run in paragraph.runs if run.text.strip().startswith(('A.', 'B.', 'C.', 'D.')))
            content = re.sub(r'^[A-D]\.\s*', '', text).strip()
            if is_correct:
                content = "*" + content
            current_options.append(content)
        elif current_q and not current_options:
            current_q += "\n" + text

    if current_q and current_options:
        questions.append({"question": current_q, "options": current_options[:]})
    
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
    
    # Đọc file Word
    original_df = parse_word_file(uploaded_file)

    # Khởi tạo session_state để giữ thay đổi
    if "edited_df" not in st.session_state or st.session_state.get("last_uploaded") != uploaded_file.name:
        st.session_state.edited_df = original_df.copy()
        st.session_state.last_uploaded = uploaded_file.name

    # ==================== CHỈNH HÀNG LOẠT ====================
    st.subheader("⚙️ Cài đặt chung cho tất cả câu hỏi")
    col1, col2, col3 = st.columns(3)
    with col1:
        new_points = st.number_input("Điểm (Points)", value=1, min_value=0, step=1)
    with col2:
        new_correct = st.text_input("Correct Feedback", value="Chúc mừng bạn! Đáp án đúng.")
    with col3:
        new_incorrect = st.text_input("Incorrect Feedback", value="Rất tiếc, đáp án chưa chính xác.")

    if st.button("🚀 Áp dụng cho TẤT CẢ câu hỏi", type="primary", use_container_width=True):
        st.session_state.edited_df["Points"] = new_points
        st.session_state.edited_df["Correct Feedback"] = new_correct
        st.session_state.edited_df["Incorrect Feedback"] = new_incorrect
        st.success("✅ Đã áp dụng thay đổi hàng loạt!")

    # ==================== BẢNG CHỈNH SỬA (ĐÃ FIX) ====================
    st.subheader("📋 Bảng câu hỏi (click vào ô để sửa)")

    # Thêm cột STT
    display_df = st.session_state.edited_df.copy()
    display_df.insert(0, "STT", range(1, len(display_df) + 1))

    edited_df = st.data_editor(
        display_df,
        use_container_width=True,
        hide_index=True,
        height=700,
        num_rows="dynamic",
        key="quiz_editor",                     # ← Quan trọng: fix lỗi lưu thay đổi
        column_config={
            "STT": st.column_config.NumberColumn("STT", width="small", disabled=True),
            "Question Type": st.column_config.TextColumn("Loại", width="small", disabled=True),
            "Question Text": st.column_config.TextColumn("Câu hỏi", width="large"),
            "Answer 1": st.column_config.TextColumn("Đáp án 1", width="medium"),
            "Answer 2": st.column_config.TextColumn("Đáp án 2", width="medium"),
            "Answer 3": st.column_config.TextColumn("Đáp án 3", width="medium"),
            "Answer 4": st.column_config.TextColumn("Đáp án 4", width="medium"),
            "Points": st.column_config.NumberColumn("Điểm", width="small"),
            "Correct Feedback": st.column_config.TextColumn("Phản hồi đúng", width="medium"),
            "Incorrect Feedback": st.column_config.TextColumn("Phản hồi sai", width="medium"),
        },
        column_order=[
            "STT", "Question Type", "Question Text", "Answer 1", "Answer 2", "Answer 3", "Answer 4",
            "Points", "Correct Feedback", "Incorrect Feedback"
        ]
    )

    # Cập nhật lại session_state sau khi chỉnh sửa
    st.session_state.edited_df = edited_df.drop(columns=["STT"], errors="ignore")

    # Nút khôi phục
    if st.button("🔄 Khôi phục bảng gốc", use_container_width=True):
        st.session_state.edited_df = original_df.copy()
        st.rerun()

    # ==================== TẢI EXCEL ====================
    if st.button("💾 Tải file Excel cho iSpring", type="primary", use_container_width=True):
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            st.session_state.edited_df.to_excel(writer, index=False, sheet_name="Sample")
        output.seek(0)

        st.download_button(
            label="📥 TẢI NGAY FILE iSpring_Questions_Import.xlsx",
            data=output,
            file_name="iSpring_Questions_Import.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
        st.success("✅ File Excel đã được tạo với **tất cả thay đổi của bạn**!")

st.caption("Ứng dụng Streamlit • Đã fix lỗi lưu thay đổi trên bảng")
