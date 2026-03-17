import streamlit as st
import pypandoc
import os
import urllib.request
import tarfile

st.set_page_config(page_title="Word 轉換神器", page_icon="🖨️", layout="wide")
st.title("阿凱老師的 Word 轉換神器 🖨️")

# =======================================================
# 轉檔引擎初始化
# =======================================================
PANDOC_VERSION = "3.1.1"
PANDOC_DIR = "/tmp/pandoc_bin"
PANDOC_EXE = os.path.join(PANDOC_DIR, f"pandoc-{PANDOC_VERSION}/bin/pandoc")

if not os.path.exists(PANDOC_EXE):
    with st.spinner("系統初始化中：正在植入核心轉檔引擎，請稍候幾秒鐘..."):
        os.makedirs(PANDOC_DIR, exist_ok=True)
        tar_path = os.path.join(PANDOC_DIR, "pandoc.tar.gz")
        url = f"https://github.com/jgm/pandoc/releases/download/{PANDOC_VERSION}/pandoc-{PANDOC_VERSION}-linux-amd64.tar.gz"
        urllib.request.urlretrieve(url, tar_path)
        with tarfile.open(tar_path, "r:gz") as tar:
            tar.extractall(path=PANDOC_DIR, filter='data')
        os.remove(tar_path)

os.environ["PYPANDOC_PANDOC"] = PANDOC_EXE
# =======================================================

# =======================================================
# 狀態管理 (Session State) 與功能函式
# =======================================================
if "input_text" not in st.session_state:
    st.session_state["input_text"] = ""

def load_uploaded_file():
    """讀取上傳檔案並填入文字框"""
    if st.session_state.uploaded_file is not None:
        try:
            st.session_state["input_text"] = st.session_state.uploaded_file.getvalue().decode("utf-8")
        except Exception as e:
            st.error(f"檔案讀取失敗，請確認是否為純文字檔。錯誤訊息：{e}")

# =======================================================
# 網頁介面區塊
# =======================================================

# 區塊 1：使用說明與專屬 Prompt
with st.expander("💡 點我展開：使用說明與專屬 AI Prompt"):
    st.markdown("請在 ChatGPT、Claude 或 Gemini 等 AI 工具中，將以下指令加在問題的最後面，以確保產出的格式完美支援本系統：")
    prompt_text = """請將所有產出的內容放入單一個『程式碼區塊 (Code block)』中輸出，並嚴格遵守以下格式，以利後續程式解析：
1. 排版保留：請使用標準 Markdown 語法處理文字排版（如 ### 標題、** 粗體）。
2. 表格規範：若有表格，請使用標準 Markdown 表格（| 表格 |），並且【表格的正上方與正下方，務必各保留一個空白行】。
3. 數學公式：請一律使用標準 LaTeX 語法呈現（行內公式使用 $ $，獨立行公式使用 $$ $$）。"""
    st.code(prompt_text, language="markdown")

st.divider()

# 區塊 2：檔案上傳與文字輸入
col1, col2 = st.columns([3, 1])
with col1:
    st.file_uploader("📂 可選擇上傳 .txt 或 .md 檔案 (內容會自動帶入下方輸入框)：", 
                     type=["txt", "md"], 
                     key="uploaded_file", 
                     on_change=load_uploaded_file)
with col2:
    st.markdown("<br>", unsafe_allow_html=True) # 調整按鈕對齊高度
    # 【修正痛點】使用強制重整 (rerun) 來解決清空按鈕被系統忽略的問題
    if st.button("🗑️ 一鍵清空", use_container_width=True):
        st.session_state["input_text"] = ""
        st.rerun()

# 核心文字輸入框
text_input = st.text_area(
    "請在此貼上 AI 產生的內容 (支援 Markdown 與 LaTeX，請使用 Ctrl+V 貼上)：", 
    key="input_text",
    height=350
)

# 加入明顯的輔助按鈕，讓使用者點擊後觸發畫面更新
st.button("👀 貼上內容後，請點此載入預覽", use_container_width=True)

# 區塊 3：預覽與下載
if text_input:
    st.divider()
    
    # 預覽畫面 (預設展開)
    with st.expander("👀 畫面預覽 (點擊可展開/收合)", expanded=True):
        st.markdown(text_input)
    
    st.divider()
    
    st.subheader("📥 匯出 Word 講義")
    
    use_two_columns = st.checkbox("啟用雙排版（適合考卷）", help="需事先將 two_columns_template.docx 範本檔上傳至系統最外層目錄")
    
    output_filename = "converted_document.docx"
    
    if st.button("🚀 開始轉換並準備下載", type="primary"):
        try:
            extra_args = []
            if use_two_columns:
                if os.path.exists("two_columns_template.docx"):
                    extra_args.append('--reference-doc=two_columns_template.docx')
                else:
                    st.warning("⚠️ 系統找不到 'two_columns_template.docx' 範本檔，將自動降級為預設單欄排版。請確認檔案已上傳至最外層目錄！")
            
            with st.spinner("努力排版中，請稍候..."):
                pypandoc.convert_text(
                    text_input, 
                    'docx', 
                    format='markdown', 
                    outputfile=output_filename,
                    extra_args=extra_args
                )
            
            with open(output_filename, "rb") as file:
                st.download_button(
                    label="✅ 一鍵下載 Word 檔",
                    data=file,
                    file_name="數學講義.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
                
        except Exception as e:
            st.error(f"轉換過程中發生錯誤：{e}")
