import streamlit as st
import pandas as pd
import os
from ruby_processor import apply_ruby_to_document

# Page config
st.set_page_config(
    page_title="Word特殊ルビ振りツール | 同人誌のルビ振りツール",
    layout="wide",
)

# Load External Resources
def load_css(file_name):
    # Use absolute path to ensure file is found regardless of CWD
    base_dir = os.path.dirname(os.path.abspath(__file__))
    file_path = os.path.join(base_dir, file_name)
    
    if os.path.exists(file_path):
        with open(file_path, encoding='utf-8') as f:
            st.markdown(f'<style>{f.read()}</style>', unsafe_allow_html=True)
    else:
        st.warning(f"Style file not found: {file_name}")

def load_html(file_name):
    base_dir = os.path.dirname(os.path.abspath(__file__))
    file_path = os.path.join(base_dir, file_name)
    
    if os.path.exists(file_path):
        with open(file_path, encoding='utf-8') as f:
            st.markdown(f.read(), unsafe_allow_html=True)

load_css("style.css")
load_html("header.html")

# Initialize Session State
if 'ruby_list' not in st.session_state:
    st.session_state.ruby_list = []
if 'step' not in st.session_state:
    st.session_state.step = 1

# --- Layout: 3 Columns for Steps ---
col1, col2, col3 = st.columns(3)

# ========================
# Step 1: File Selection
# ========================
with col1:
    with st.container(border=True):
        st.markdown("### 1. ファイル選択")
        st.write("対象のWordファイル (.docx) を選択してください。")
        
        uploaded_file = st.file_uploader("Upload Word", type=['docx'], label_visibility="collapsed")
        
        if uploaded_file is not None:
            st.success("✅ 読み込み完了")
            # If step is still 1, show button to proceed
            if st.session_state.step == 1:
                if st.button("次へ進む (ルビ設定)", key="next_to_2", type="primary", use_container_width=True):
                    st.session_state.step = 2
                    st.rerun()
            else:
                # Assuming if step > 1, user might want to re-upload. 
                # Keeping it simple: Just show text that it's active.
                pass

# ========================
# Step 2: Ruby Settings
# ========================
with col2:
    # Only show content if step >= 2
    if st.session_state.step >= 2:
        with st.container(border=True):
            st.markdown("### 2. ルビ設定")
            
            # Input Form
            with st.form("add_ruby_form", clear_on_submit=True, border=False):
                noun = st.text_input("名詞 (漢字)", placeholder="例: 運命")
                ruby = st.text_input("ルビ (読み)", placeholder="例: さだめ")
                if st.form_submit_button("リストに追加", use_container_width=True):
                    if noun and ruby:
                        st.session_state.ruby_list.append({"noun": noun, "ruby": ruby})
                        st.toast(f"追加: {noun}({ruby})")
                    else:
                        st.warning("両方入力してください")

            # List View
            if st.session_state.ruby_list:
                st.markdown("---")
                st.caption(f"登録数: {len(st.session_state.ruby_list)}件")
                
                df = pd.DataFrame(st.session_state.ruby_list)
                edited_df = st.data_editor(
                    df,
                    column_config={
                        "noun": st.column_config.TextColumn("名詞"),
                        "ruby": st.column_config.TextColumn("ルビ"),
                    },
                    use_container_width=True,
                    num_rows="dynamic",
                    key="ruby_editor",
                    hide_index=True
                )
                
                # Sync edits
                current = edited_df.to_dict('records')
                # data_editor uses column names from dataframe for records
                # Our DF created from list of dicts {'noun':..., 'ruby':...} matches.
                if current != st.session_state.ruby_list:
                    st.session_state.ruby_list = current

            # Next Button
            st.write("")
            if st.session_state.ruby_list:
                if st.session_state.step == 2:
                    if st.button("次へ進む (適用)", key="next_to_3", type="primary", use_container_width=True):
                        st.session_state.step = 3
                        st.rerun()
            else:
                if st.session_state.step == 2:
                    st.info("リストを追加してください")

    else:
        # Placeholder or Empty
        with st.container(border=True):
            st.markdown("### 2. ルビ設定")
            st.caption("ファイルを選択すると編集できます")


# ========================
# Step 3: Execution Mode
# ========================
with col3:
    if st.session_state.step >= 3:
        with st.container(border=True):
            st.markdown("### 3. モード・実行")
            
            mode = st.radio(
                "適用モード",
                ('once', 'per_page', 'all'),
                format_func=lambda x: {
                    'once': '最初の一回のみ',
                    'per_page': 'ページ毎に一回',
                    'all': 'すべて'
                }[x]
            )
            
            st.markdown("---")
            if st.button("変換実行", type="primary", use_container_width=True):
                if not uploaded_file:
                    st.error("ファイルが未選択です")
                    st.session_state.step = 1
                    st.rerun()
                elif not st.session_state.ruby_list:
                    st.error("ルビ設定が空です")
                    st.session_state.step = 2
                    st.rerun()
                else:
                    # Execute
                    temp_dir = "temp"
                    os.makedirs(temp_dir, exist_ok=True)
                    input_path = os.path.join(temp_dir, uploaded_file.name)
                    base_name = os.path.splitext(uploaded_file.name)[0]
                    output_filename = f"{base_name}_ruby.docx"
                    output_path = os.path.join(temp_dir, output_filename)
                    
                    with open(input_path, "wb") as f:
                        f.write(uploaded_file.getbuffer())
                    
                    ruby_settings = [{'word': str(item['noun']), 'ruby': str(item['ruby'])} 
                                   for item in st.session_state.ruby_list]
                    
                    with st.spinner("変換中..."):
                        try:
                            result_path = apply_ruby_to_document(input_path, output_path, ruby_settings, mode=mode)
                            st.success("完了！")
                            st.balloons()
                            
                            with open(result_path, "rb") as f:
                                st.download_button(
                                    "ダウンロード",
                                    data=f.read(),
                                    file_name=output_filename,
                                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                    type="primary",
                                    use_container_width=True
                                )
                        except Exception as e:
                            st.error(f"Error: {e}")

    else:
        with st.container(border=True):
            st.markdown("### 3. モード・実行")
            st.caption("設定完了後に選択できます")
