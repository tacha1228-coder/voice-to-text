import streamlit as st
import os
import tempfile
import io

from faster_whisper import WhisperModel
import yt_dlp
from docx import Document
from docx.shared import Pt
import arabic_reshaper
from bidi.algorithm import get_display
from weasyprint import HTML

st.set_page_config(page_title="محول الصوت إلى نص", page_icon="🎙️", layout="wide")

st.title("🎙️ محول الصوت إلى Word / PDF")
st.write("حوّل أي صوت أو رابط يوتيوب إلى ملف Word أو PDF مجانًا")

with st.sidebar:
    st.header("⚙️ الإعدادات")
    lang_choice = st.selectbox("اللغة", ["تلقائي", "العربية", "الإنجليزية", "الفرنسية"])
    lang_map = {"تلقائي": None, "العربية": "ar", "الإنجليزية": "en", "الفرنسية": "fr"}

tab1, tab2 = st.tabs(["🔗 رابط يوتيوب", "📁 رفع ملف صوتي"])

audio_path = None
url = ""

with tab1:
    url = st.text_input("الصقي رابط يوتيوب هنا")
    if url:
        audio_path = "downloaded_audio"

with tab2:
    uploaded = st.file_uploader("اختاري ملفًا صوتيًا", type=["mp3", "wav", "m4a", "ogg", "mp4", "webm"])
    if uploaded:
        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".wav")
        tmp.write(uploaded.read())
        tmp.close()
        audio_path = tmp.name

def download_youtube(url):
    base_opts = {
        'format': 'bestaudio/best',
        'outtmpl': 'audio.%(ext)s',
        'postprocessors': [{'key': 'FFmpegExtractAudio', 'preferredcodec': 'wav'}],
        'quiet': True,
        'no_warnings': True,
        'force_ipv4': True,
        'extractor_args': {
            'youtube': {
                'player_client': ['android', 'web_safari', 'web'],
            }
        },
        'user_agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
    }
    try:
        with yt_dlp.YoutubeDL(base_opts) as ydl:
            ydl.download([url])
        return "audio.wav"
    except Exception as e1:
        try:
            st.info("🔄 الطريقة الأولى فشلت، جارٍ تجربة طريقة أخرى...")
            alt_opts = {
                'format': 'bestaudio/best',
                'outtmpl': 'audio.%(ext)s',
                'postprocessors': [{'key': 'FFmpegExtractAudio', 'preferredcodec': 'wav'}],
                'quiet': True,
                'no_warnings': True,
                'force_ipv4': True,
                'extractor_args': {
                    'youtube': {'player_client': ['ios', 'mweb']}
                },
            }
            with yt_dlp.YoutubeDL(alt_opts) as ydl:
                ydl.download([url])
            return "audio.wav"
        except Exception as e2:
            raise Exception(f"تعذّر التحميل. جرّبي فيديو آخر أو استخدمي رفع الملف.")

if audio_path and st.button("🚀 ابدأ التحويل", type="primary"):
    progress = st.progress(0)
    status = st.empty()
    try:
        if audio_path == "downloaded_audio":
            status.text("⏬ جارٍ تحميل الصوت من يوتيوب...")
            progress.progress(10)
            audio_path = download_youtube(url)
        
        status.text("🎧 جارٍ تفريغ النص... (قد يستغرق وقتًا)")
        progress.progress(30)
        
        model = WhisperModel("small", device="cpu", compute_type="int8")
        segments, info = model.transcribe(
            audio_path, 
            language=lang_map[lang_choice], 
            vad_filter=True,
            vad_parameters=dict(min_silence_duration_ms=500)
        )
        
        full_text = ""
        for seg in segments:
            full_text += seg.text + "\n"
        
        progress.progress(85)
        st.session_state['text'] = full_text.strip()
        status.text("✅ تم التفريغ!")
        progress.progress(100)
        st.success("🎉 اكتمل التفريغ!")
        
    except Exception as e:
        st.error(f"حدث خطأ: {e}")

if 'text' in st.session_state and st.session_state['text']:
    st.subheader("📝 النص المفرغ")
    st.text_area("يمكنك نسخ النص من هنا", st.session_state['text'], height=250)
    text = st.session_state['text']
    
    col1, col2 = st.columns(2)
    
    with col1:
        doc = Document()
        doc.styles['Normal'].font.name = 'Arial'
        doc.styles['Normal'].font.size = Pt(12)
        for line in text.split('\n'):
            if line.strip():
                doc.add_paragraph(line)
        buf = io.BytesIO()
        doc.save(buf)
        buf.seek(0)
        st.download_button("📄 تحميل Word", buf, "transcript.docx",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
    
    with col2:
        lines_html = ""
        for line in text.split('\n'):
            if line.strip():
                reshaped = arabic_reshaper.reshape(line)
                bidi = get_display(reshaped)
                lines_html += f"<p>{bidi}</p>"
        html = f"""<html dir="rtl"><head><meta charset="utf-8">
        <style>body{{font-family:Arial;direction:rtl;text-align:right;padding:20px}}
        p{{line-height:1.8}}</style></head><body>{lines_html}</body></html>"""
        pdf_buf = io.BytesIO()
        HTML(string=html).write_pdf(pdf_buf)
        pdf_buf.seek(0)
        st.download_button("📕 تحميل PDF", pdf_buf, "transcript.pdf", "application/pdf")
