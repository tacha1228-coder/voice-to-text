import streamlit as st
import io
import tempfile
import os
import time

from faster_whisper import WhisperModel
import yt_dlp
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import arabic_reshaper
from bidi.algorithm import get_display
from weasyprint import HTML

st.set_page_config(page_title="محول الصوت إلى نص", page_icon="🎙️", layout="wide", initial_sidebar_state="auto")

st.error("""
### ⚠️ تنبيه شرعي وأخلاقي هام

**لا أُحل استخدام هذا التطبيق في:**
- ❌ السرقات العلمية أو انتحال المحتوى
- ❌ تفريغ محتويات الكتب التي يُخاف فيها شرع الله
- ❌ أي استخدام يخالف الأمانة العلمية

**من يفعل ذلك فإني خصيمه يوم القيامة، ولا سماح بيننا في الدنيا ولا في الآخرة.**
""")

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
        file_size_mb = uploaded.size / (1024 * 1024)
        st.success(f"✅ تم رفع الملف: {uploaded.name} ({file_size_mb:.2f} MB)")
        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".wav")
        tmp.write(uploaded.read())
        tmp.close()
        audio_path = tmp.name

def format_time(seconds):
    if seconds < 60:
        return f"{int(seconds)} ثانية"
    elif seconds < 3600:
        m = int(seconds // 60)
        s = int(seconds % 60)
        return f"{m} د {s} ث"
    else:
        h = int(seconds // 3600)
        m = int((seconds % 3600) // 60)
        return f"{h} س {m} د"

def set_rtl(paragraph):
    pPr = paragraph._p.get_or_add_pPr()
    bidi = OxmlElement('w:bidi')
    pPr.append(bidi)
    paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT

def download_youtube(url):
    base_opts = {
        'format': 'bestaudio/best',
        'outtmpl': 'audio.%(ext)s',
        'postprocessors': [{'key': 'FFmpegExtractAudio', 'preferredcodec': 'wav'}],
        'quiet': True,
        'no_warnings': True,
        'force_ipv4': True,
        'extractor_args': {'youtube': {'player_client': ['android', 'web_safari', 'web']}},
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
                'extractor_args': {'youtube': {'player_client': ['ios', 'mweb']}},
            }
            with yt_dlp.YoutubeDL(alt_opts) as ydl:
                ydl.download([url])
            return "audio.wav"
        except Exception as e2:
            raise Exception("تعذّر التحميل. جرّبي فيديو آخر أو استخدمي رفع الملف.")

if audio_path and st.button("🚀 ابدأ التحويل", type="primary"):
    try:
        # ================= المرحلة 1 =================
        if audio_path == "downloaded_audio":
            st.markdown("### 📥 المرحلة 1 من 3: تحميل الصوت من يوتيوب")
            bar1 = st.progress(0, text="0%")
            bar1.progress(30, text="30% - الاتصال بيوتيوب...")
            time.sleep(0.2)
            bar1.progress(60, text="60% - جارٍ التحميل...")
            audio_path = download_youtube(url)
            bar1.progress(100, text="100% - ✅ تم التحميل")
        else:
            st.markdown("### 📤 المرحلة 1 من 3: تجهيز الملف الصوتي")
            bar1 = st.progress(0, text="0%")
            for i in range(1, 11):
                bar1.progress(i * 10, text=f"{i*10}% - تجهيز الملف")
                time.sleep(0.05)
            bar1.progress(100, text="100% - ✅ الملف جاهز")
        
        # ================= المرحلة 2: التفريغ =================
        st.markdown("### 🎧 المرحلة 2 من 3: تفريغ النص (تحميل الموديل)")
        bar2 = st.progress(0, text="0%")
        status2 = st.empty()
        
        status2.info("⏳ يتم تحميل موديل Whisper لأول مرة (قد يستغرق 1-2 دقيقة)...")
        bar2.progress(5, text="5% - تحميل الموديل...")
        
        model = WhisperModel("small", device="cpu", compute_type="int8")
        
        bar2.progress(15, text="15% - جارٍ التفريغ...")
        status2.info("🎙️ جارٍ تفريغ النص...")
        
        segments, info = model.transcribe(
            audio_path,
            language=lang_map[lang_choice],
            vad_filter=True,
            vad_parameters=dict(min_silence_duration_ms=500)
        )
        
        full_text = ""
        seg_list = list(segments)
        total_segs = len(seg_list) if seg_list else 1
        
        # محاولة الحصول على مدة الصوت الكلية
        try:
            total_duration = info.duration
        except:
            total_duration = 0
        
        start_time = time.time()
        
        for i, seg in enumerate(seg_list):
            full_text += seg.text + "\n"
            
            # النسبة بناءً على الطابع الزمني للمقطع
            if total_duration > 0:
                percent = min(15 + int((seg.end / total_duration) * 80), 95)
            else:
                percent = 15 + int((i / total_segs) * 80)
            
            # حساب الوقت المتبقي
            elapsed = time.time() - start_time
            if i > 0 and total_duration > 0:
                progress_ratio = seg.end / total_duration
                if progress_ratio > 0:
                    total_estimated = elapsed / progress_ratio
                    remaining = total_estimated - elapsed
                    time_str = f" | ⏳ الوقت المتبقي: {format_time(remaining)}"
                else:
                    time_str = ""
            else:
                time_str = " | ⏳ جاري الحساب..."
            
            current_time_str = format_time(seg.end) if total_duration > 0 else ""
            bar2.progress(percent / 100, text=f"{percent}% - عند الدقيقة {current_time_str}{time_str}")
        
        bar2.progress(100, text="100% - ✅ تم التفريغ")
        status2.success("✅ تم تفريغ النص بنجاح!")
        st.session_state['text'] = full_text.strip()
        
        # ================= المرحلة 3: إنشاء الملفات =================
        st.markdown("### 📄 المرحلة 3 من 3: إنشاء ملفات Word و PDF")
        bar3 = st.progress(0, text="0%")
        
        bar3.progress(30, text="30% - إنشاء Word...")
        text = st.session_state['text']
        
        doc = Document()
        style = doc.styles['Normal']
        style.font.name = 'Arial'
        style.font.size = Pt(12)
        style.element.rPr.rFonts.set(qn('w:cs'), 'Arial')
        
        for line in text.split('\n'):
            if line.strip():
                p = doc.add_paragraph(line)
                set_rtl(p)
        
        docx_buf = io.BytesIO()
        doc.save(docx_buf)
        docx_buf.seek(0)
        
        bar3.progress(70, text="70% - إنشاء PDF...")
        
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
        
        bar3.progress(100, text="100% - ✅ الملفات جاهزة")
        
        st.divider()
        st.success("🎉 اكتمل التحويل!")
        
        # ================= التحميل =================
        st.subheader("📝 النص المفرغ")
        st.text_area("يمكنك نسخ النص من هنا", text, height=250)
        
        col1, col2 = st.columns(2)
        with col1:
            st.download_button("📄 تحميل Word", docx_buf, "transcript.docx",
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
        with col2:
            st.download_button("📕 تحميل PDF", pdf_buf, "transcript.pdf", "application/pdf")
        
    except Exception as e:
        st.error(f"حدث خطأ: {e}")
