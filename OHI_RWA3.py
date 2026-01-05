import streamlit as st
import os
from groq import Groq
import base64
from io import BytesIO
from PIL import Image
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

# Konfigurasi halaman
st.set_page_config(
    page_title="OHI Rapport Writer Assistance",
    page_icon="📊",
    layout="wide"
)

# Custom CSS
st.markdown("""
    <style>
    .main-header {
        font-size: 2.5rem;
        font-weight: bold;
        color: #1E3A8A;
        text-align: center;
        margin-bottom: 1rem;
    }
    .sub-header {
        font-size: 1.2rem;
        color: #475569;
        text-align: center;
        margin-bottom: 2rem;
    }
    .section-header {
        font-size: 1.5rem;
        font-weight: bold;
        color: #1E40AF;
        margin-top: 2rem;
        margin-bottom: 1rem;
        border-bottom: 2px solid #3B82F6;
        padding-bottom: 0.5rem;
    }
    .info-box {
        background-color: #EFF6FF;
        padding: 1.5rem;
        border-radius: 0.5rem;
        border-left: 4px solid #3B82F6;
        margin: 1rem 0;
    }
    .success-box {
        background-color: #F0FDF4;
        padding: 1rem;
        border-radius: 0.5rem;
        border-left: 4px solid #10B981;
        margin: 1rem 0;
    }
    </style>
""", unsafe_allow_html=True)

# Fungsi untuk mendapatkan API Key dari secrets
@st.cache_resource
def get_api_key():
    """Ambil API Key dari Streamlit secrets"""
    try:
        api_key = st.secrets["GROQ_API_KEY"]
        return api_key
    except:
        api_key = os.getenv("GROQ_API_KEY")
        if not api_key:
            st.error("""
            ❌ **API Key tidak ditemukan!**
            
            **Untuk Streamlit Cloud:**
            1. Settings → Secrets
            2. Tambahkan: `GROQ_API_KEY = "gsk_..."`
            
            **Untuk Local:**
            1. Buat `.streamlit/secrets.toml`
            2. Tambahkan: `GROQ_API_KEY = "gsk_..."`
            """)
            st.stop()
        return api_key

def extract_table_data(table):
    """Ekstrak data dari tabel"""
    table_text = []
    for row in table.rows:
        row_data = [cell.text.strip() for cell in row.cells]
        table_text.append(" | ".join(row_data))
    return "\n".join(table_text)

def extract_images_from_pptx(pptx_file, max_images=13):
    """Ekstrak gambar dari PowerPoint"""
    images = []
    try:
        prs = Presentation(pptx_file)
        img_count = 0
        
        for slide_num, slide in enumerate(prs.slides, 1):
            if img_count >= max_images:
                break
                
            for shape in slide.shapes:
                if img_count >= max_images:
                    break
                    
                if hasattr(shape, "image"):
                    try:
                        pil_image = Image.open(BytesIO(shape.image.blob))
                        
                        if pil_image.mode == 'RGBA':
                            bg = Image.new('RGB', pil_image.size, (255, 255, 255))
                            bg.paste(pil_image, mask=pil_image.split()[3])
                            pil_image = bg
                        elif pil_image.mode != 'RGB':
                            pil_image = pil_image.convert('RGB')
                        
                        pil_image.thumbnail((1024, 1024), Image.Resampling.LANCZOS)
                        
                        buffered = BytesIO()
                        pil_image.save(buffered, format="JPEG", quality=80, optimize=True)
                        img_str = base64.b64encode(buffered.getvalue()).decode()
                        
                        size_kb = len(img_str) / 1024
                        if size_kb > 500:
                            continue
                        
                        images.append({
                            'data': img_str,
                            'slide': slide_num,
                            'type': 'image',
                            'size_kb': size_kb
                        })
                        img_count += 1
                    except Exception as e:
                        st.warning(f"Gagal ekstrak gambar dari slide {slide_num}")
                        
                elif shape.shape_type == MSO_SHAPE_TYPE.TABLE:
                    try:
                        table_data = extract_table_data(shape.table)
                        if len(table_data) > 2000:
                            table_data = table_data[:2000] + "..."
                        images.append({
                            'data': table_data,
                            'slide': slide_num,
                            'type': 'table'
                        })
                    except:
                        pass
        
        return images, img_count
    except Exception as e:
        st.error(f"Error membaca PowerPoint: {str(e)}")
        return [], 0

def encode_image(image_file):
    """Encode image dengan kompresi"""
    try:
        image = Image.open(image_file)
        
        if image.mode == 'RGBA':
            bg = Image.new('RGB', image.size, (255, 255, 255))
            bg.paste(image, mask=image.split()[3])
            image = bg
        elif image.mode != 'RGB':
            image = image.convert('RGB')
        
        image.thumbnail((1024, 1024), Image.Resampling.LANCZOS)
        
        buffered = BytesIO()
        image.save(buffered, format="JPEG", quality=85, optimize=True)
        return base64.b64encode(buffered.getvalue()).decode()
    except Exception as e:
        st.error(f"Error encoding image: {str(e)}")
        return None

def analyze_with_groq(api_key, images, tables_text, analysis_type="initial"):
    """Analisis dengan Groq API"""
    try:
        client = Groq(api_key=api_key)
        content = []
        
        if analysis_type == "initial":
            prompt = """Analisis data OHI dari gambar/tabel yang diberikan.

Ekstrak:
1. SEMUA skor numerik dengan dimensinya
2. TOP 5-7 skor tertinggi (Kekuatan)
3. BOTTOM 5-7 skor terendah (Perbaikan)
4. Skor rata-rata dan pola

Output dalam format terstruktur (minimal 600 kata)."""
        
        else:
            prompt = """Buatlah laporan OHI SANGAT KOMPREHENSIF (2000-2500 kata):

**BAGIAN 1: KEKUATAN ORGANISASI** (500-600 kata)
Untuk setiap dimensi kuat:
- Skor & kenapa ini kekuatan
- Dampak positif konkret
- Cara mempertahankan & leverage
- Contoh praktis

**BAGIAN 2: AREA PERBAIKAN** (700-800 kata)
Untuk setiap dimensi lemah:
- Root cause analysis detail
- Dampak jika tak diperbaiki
- Rekomendasi SANGAT spesifik dengan langkah-langkah
- Quick wins (1-2 minggu) dengan contoh detail
- Initiatives (1-3 bulan) dengan roadmap
- Long-term (3-6 bulan) dengan strategy
- KPIs & cara tracking

**BAGIAN 3: REKOMENDASI LEADERSHIP** (800-1100 kata)
Minimal 12-15 rekomendasi detail:
- Setiap rekomendasi dengan contoh SANGAT spesifik
- Quick wins, medium-term, long-term
- Leadership behaviors harian dengan contoh
- Communication strategy
- Timeline & resource requirements
- Monitoring approach

Contoh BAIK:
"Setiap Senin 9:00-9:15, stand-up meeting. Format: (1) Achievements minggu lalu dengan metrics (2) Top 3 priorities minggu ini (3) Blockers & support needed. Gunakan Miro board shared. Rotate facilitator. Record di Notion."

Gunakan Bahasa Indonesia, profesional, detail, actionable."""

        content.append({"type": "text", "text": prompt})
        
        if tables_text:
            if len(tables_text) > 3000:
                tables_text = tables_text[:3000] + "..."
            content.append({"type": "text", "text": f"\n=== DATA TABEL ===\n{tables_text}\n==="})
        
        max_imgs = 3 if analysis_type == "initial" else 5
        for idx, img in enumerate(images[:max_imgs]):
            if isinstance(img, dict) and img.get('type') == 'image':
                content.append({
                    "type": "image_url",
                    "image_url": {"url": f"data:image/jpeg;base64,{img['data']}"}
                })
            elif isinstance(img, str):
                content.append({
                    "type": "image_url",
                    "image_url": {"url": f"data:image/jpeg;base64,{img}"}
                })
        
        max_tokens = 4096 if analysis_type == "initial" else 16384
        
        response = client.chat.completions.create(
            messages=[{"role": "user", "content": content}],
            model="meta-llama/llama-4-maverick-17b-128e-instruct",
            temperature=0.7,
            max_tokens=max_tokens,
        )
        
        return response.choices[0].message.content
        
    except Exception as e:
        st.error(f"Error analisis: {str(e)}")
        return None

# Main App
api_key = get_api_key()

st.markdown('<div class="main-header">📊 OHI Rapport Writer Assistance</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">McKinsey Organizational Health Index Framework & AI</div>', unsafe_allow_html=True)

with st.expander("ℹ️ Tentang Aplikasi", expanded=False):
    st.markdown("""
    **Framework OHI McKinsey - 9 Outcomes:**
    Direction • Leadership • Culture & Climate • Accountability • 
    Coordination & Control • Capabilities • Motivation • 
    External Orientation • Innovation & Learning
    
    **Output:** Laporan komprehensif 2000+ kata dengan analisis detail dan rekomendasi actionable
    """)

st.markdown('<div class="section-header">📤 Upload File OHI</div>', unsafe_allow_html=True)

upload_type = st.radio("Pilih tipe file:", ["PowerPoint (.pptx)", "Gambar (PNG/JPG)"], horizontal=True)

if upload_type == "PowerPoint (.pptx)":
    pptx_file = st.file_uploader("Upload PowerPoint", type=["pptx"])
    
    if pptx_file:
        st.success(f"✅ File: {pptx_file.name}")
        
        if st.button("🚀 Analisis & Generate Rapport", type="primary", use_container_width=True):
            with st.spinner("Mengekstrak konten..."):
                pptx_file.seek(0)
                extracted, img_count = extract_images_from_pptx(pptx_file)
                
                if extracted:
                    images = [i for i in extracted if i.get('type') == 'image']
                    tables = [i for i in extracted if i.get('type') == 'table']
                    st.success(f"✅ {len(images)} gambar, {len(tables)} tabel")
                    
                    tables_text = "\n\n".join([t['data'] for t in tables])[:5000]
                    
                    st.info("📊 Tahap 1: Ekstraksi data...")
                    initial = analyze_with_groq(api_key, images[:5], tables_text, "initial")
                    
                    if initial:
                        with st.expander("📋 Data Terdeteksi", expanded=True):
                            st.markdown(initial)
                        
                        st.info("📝 Tahap 2: Menyusun rapport komprehensif...")
                        final = analyze_with_groq(api_key, images[:3], "", "final")
                        
                        if final:
                            st.markdown('<div class="section-header">📄 Rapport Lengkap</div>', unsafe_allow_html=True)
                            st.markdown(final)
                            
                            col1, col2, col3 = st.columns(3)
                            with col1:
                                st.download_button("⬇️ TXT", final, "OHI_Rapport.txt", "text/plain", use_container_width=True)
                            with col2:
                                st.download_button("⬇️ MD", final, "OHI_Rapport.md", "text/markdown", use_container_width=True)
                            with col3:
                                full = f"# DATA EKSTRAKSI\n\n{initial}\n\n---\n\n# RAPPORT\n\n{final}"
                                st.download_button("⬇️ Lengkap", full, "OHI_Complete.md", use_container_width=True)

else:
    files = st.file_uploader("Upload gambar", type=["png", "jpg", "jpeg"], accept_multiple_files=True)
    
    if files:
        st.success(f"✅ {len(files)} gambar")
        
        if st.button("🚀 Analisis", type="primary", use_container_width=True):
            with st.spinner("Menganalisis..."):
                encoded = [encode_image(f) for f in files if encode_image(f)]
                
                if encoded:
                    st.info("📊 Ekstraksi data...")
                    initial = analyze_with_groq(api_key, encoded, "", "initial")
                    
                    if initial:
                        with st.expander("📋 Data", expanded=True):
                            st.markdown(initial)
                        
                        st.info("📝 Menyusun rapport...")
                        final = analyze_with_groq(api_key, encoded, "", "final")
                        
                        if final:
                            st.markdown(final)
                            
                            col1, col2 = st.columns(2)
                            with col1:
                                st.download_button("⬇️ TXT", final, "OHI_Rapport.txt", use_container_width=True)
                            with col2:
                                st.download_button("⬇️ MD", final, "OHI_Rapport.md", use_container_width=True)

st.markdown("---")
st.markdown("""
<div style='text-align: center; color: #64748B; padding: 1rem;'>
    <p><strong>OHI Rapport Writer Assistance</strong></p>
    <p>Powered by McKinsey OHI Framework & Groq AI</p>
</div>
""", unsafe_allow_html=True)