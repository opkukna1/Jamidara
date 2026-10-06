import os
import re
import json
import io
import zipfile
from typing import List, Dict, Any
import pandas as pd
from docxtpl import DocxTemplate

from fastapi import FastAPI, UploadFile, File, HTTPException, BackgroundTasks, APIRouter
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse
import firebase_admin
from firebase_admin import credentials, firestore
import google.generativeai as genai

# ==========================================
# 1. APP & CORS SETUP (Scalable Architecture)
# ==========================================
app = FastAPI(
    title="Audit Dispatch Portal API",
    description="Backend API for Dispatch Processing and Document Generation",
    version="1.0.0"
)

# CORS Enable kiya hai taaki GitHub Pages Frontend se request allow ho sake
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # Production me apna GitHub Pages domain de sakte hain
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# ==========================================
# 2. FIREBASE & GEMINI INITIALIZATION
# ==========================================
# Firebase Service Account Key Environment Variable se load hoga
if not firebase_admin._apps:
    cred_json = os.getenv("FIREBASE_CREDENTIALS_JSON")
    if cred_json:
        cred_dict = json.loads(cred_json)
        cred = credentials.Certificate(cred_dict)
        firebase_admin.initialize_app(cred)
    else:
        # Fallback to default/local file if available
        firebase_admin.initialize_app()

db = firestore.client()

# Gemini API Setup
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY", "")
if GEMINI_API_KEY:
    genai.configure(api_key=GEMINI_API_KEY)


# ==========================================
# 3. HELPER FUNCTIONS
# ==========================================
def normalize_str(val: Any) -> str:
    """String normalization for accurate key matching."""
    if not isinstance(val, str) or pd.isna(val):
        return ""
    val = val.strip()
    return val.replace("Panchayat Samiti ", "").replace("Gram Panchayat ", "").strip().lower()


def get_hindi_transliteration(unmapped_items: List[Dict[str, str]]) -> List[Dict[str, str]]:
    """Sends ONLY new unmapped entries to Gemini AI for Hindi conversion."""
    if not unmapped_items or not GEMINI_API_KEY:
        return []

    prompt = f"""
    You are an official administrative Hindi transliterator for Rajasthan Government documents.
    Convert the following list of Gram Panchayats, Panchayat Samitis, and Districts from English to official Hindi.

    Input Data:
    {json.dumps(unmapped_items, ensure_ascii=False, indent=2)}

    Return ONLY a valid JSON Array with exact structure:
    [
      {{
        "GP_EN": "Gram Panchayat ...",
        "GP_HI": "ग्राम पंचायत ...",
        "PS_EN": "...",
        "PS_HI": "...",
        "DIST_EN": "...",
        "DIST_HI": "..."
      }}
    ]
    """
    model = genai.GenerativeModel('gemini-2.5-flash')
    response = model.generate_content(prompt)
    
    response_text = response.text.strip()
    if response_text.startswith("```json"):
        response_text = response_text[7:-3].strip()
    elif response_text.startswith("```"):
        response_text = response_text[3:-3].strip()

    try:
        return json.loads(response_text)
    except Exception as e:
        print(f"Gemini response parsing error: {e}")
        return []


def get_master_mappings_from_firestore() -> Dict[tuple, Dict[str, str]]:
    """Fetches all existing mappings from Firestore database."""
    docs = db.collection("mappings").stream()
    mapping_dict = {}
    for doc in docs:
        data = doc.to_dict()
        key = (normalize_str(data.get("GP_EN")), normalize_str(data.get("PS_EN")))
        mapping_dict[key] = data
    return mapping_dict


# ==========================================
# 4. ROUTERS (Modular Structure for Future Features)
# ==========================================

# Router 1: Dispatch Processing Module
dispatch_router = APIRouter(prefix="/api/v1/dispatch", tags=["Dispatch Module"])

@dispatch_router.post("/process")
async def process_dispatch_file(
    dispatch_file: UploadFile = File(...),
    template_file: UploadFile = File(...)
):
    """
    Uploads Dispatch Excel and DOCX Template, checks new entries via AI,
    updates Firestore, and returns ZIP of generated Covering Letters.
    """
    try:
        # Read uploaded files in memory
        dispatch_bytes = await dispatch_file.read()
        template_bytes = await template_file.read()

        df_disp = pd.read_excel(io.BytesIO(dispatch_bytes))

        # 1. Load existing Firestore mappings
        lookup_map = get_master_mappings_from_firestore()

        # 2. Identify unmapped entries
        unmapped_entries = []
        unmapped_keys = set()

        for _, row in df_disp.iterrows():
            gp_en = str(row.get('Unit Name', '')).strip()
            ps_en = str(row.get('Parent Name', '')).strip()
            dist_en = str(row.get('District Name', '')).strip()

            norm_key = (normalize_str(gp_en), normalize_str(ps_en))

            if norm_key not in lookup_map and norm_key not in unmapped_keys and gp_en:
                unmapped_keys.add(norm_key)
                unmapped_entries.append({
                    "GP_EN": gp_en,
                    "PS_EN": ps_en.replace("Panchayat Samiti ", "").strip(),
                    "DIST_EN": dist_en
                })

        # 3. Process new entries with AI & save to Firestore
        if unmapped_entries:
            ai_results = get_hindi_transliteration(unmapped_entries)
            batch = db.batch()
            for item in ai_results:
                norm_key = (normalize_str(item['GP_EN']), normalize_str(item['PS_EN']))
                lookup_map[norm_key] = item
                
                # Document ID generated from GP & PS
                doc_id = f"{normalize_str(item['GP_EN'])}_{normalize_str(item['PS_EN'])}"
                doc_ref = db.collection("mappings").document(doc_id)
                batch.set(doc_ref, item)
            batch.commit()

        # 4. Generate DOCX files in-memory ZIP
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
            for idx, row in df_disp.iterrows():
                gp_en = str(row.get('Unit Name', '')).strip()
                ps_en = str(row.get('Parent Name', '')).strip()
                norm_key = (normalize_str(gp_en), normalize_str(ps_en))

                map_data = lookup_map.get(norm_key, {})

                dispatch_no = str(row.get('Dispatch Name', '')).strip()
                audit_party = str(row.get('Audit Party No', ''))
                
                year_match = re.search(r'\d{4}-\d{2,4}', audit_party)
                year_str = year_match.group(0) if year_match else "2026-2027"

                para_count = row.get('Converted to Para in Nos', 0)
                try:
                    para_count = int(para_count)
                except ValueError:
                    para_count = 0

                para_breakup = f"आक्षेप सं. 1 से {para_count}" if para_count > 0 else "आक्षेप सं. 0"

                context = {
                    "OFFICE_NAME": map_data.get('DIST_HI', 'बीकानेर'),
                    "DIVISION_NAME": map_data.get('DIST_HI', 'बीकानेर'),
                    "YEAR": year_str,
                    "DISPATCH_NO": dispatch_no,
                    "DATE": str(row.get('Report Approval Date', '06/10/2026')),
                    "PS_NAME_HI": map_data.get('PS_HI', ps_en),
                    "DISTRICT_HI": map_data.get('DIST_HI', 'बीकानेर'),
                    "GP_NAME_HI": map_data.get('GP_HI', gp_en),
                    "PARA_COUNT": para_count,
                    "PARA_BREAKUP": para_breakup
                }

                # Render docx from memory
                tpl_stream = io.BytesIO(template_bytes)
                doc = DocxTemplate(tpl_stream)
                doc.render(context)
                
                doc_io = io.BytesIO()
                doc.save(doc_io)
                doc_io.seek(0)

                safe_gp_name = re.sub(r'[^\w\-_\. ]', '_', map_data.get('GP_EN', f"GP_{idx}"))
                filename = f"Covering_Letter_{safe_gp_name}.docx"
                zip_file.writestr(filename, doc_io.getvalue())

        zip_buffer.seek(0)
        return StreamingResponse(
            zip_buffer,
            media_type="application/zip",
            headers={"Content-Disposition": "attachment; filename=Covering_Letters.zip"}
        )

    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


# Router 2: Database Management Module (Future Updates Ke Liye Space)
mapping_router = APIRouter(prefix="/api/v1/mappings", tags=["Master Mapping Database"])

@mapping_router.get("/list")
async def list_mappings():
    """Returns all mapped Gram Panchayats stored in Firestore."""
    docs = db.collection("mappings").stream()
    return [doc.to_dict() for doc in docs]


# Registering Routers to App
app.include_router(dispatch_router)
app.include_router(mapping_router)


# Root Health Check
@app.get("/")
def root():
    return {"status": "online", "message": "Audit Portal API is running smoothly."}
