import streamlit as st
from weasyprint import HTML
import docx
import io
from datetime import datetime

def calculate_prices(bana, land, interior):
    # منطق استخراج قیمت‌ها
    return {
        'plan': 350, 'land': 100, 'phase2': 100,
        'struct': 70, 'mech': 70, 'elec': 70,
        'interior': 400, 'interior_p2': 80
    }

def generate_documents(client, project, bana, land, interior, location):
    prices = calculate_prices(bana, land, interior)
    
    # 1. تولید PDF در حافظه (بدون ذخیره روی هارد)
    html_template = f"""
    <html lang="fa" dir="rtl">
    <head><style>body {{ font-family: Tahoma; }}</style></head>
    <body>
        <h2>به نام خدا</h2>
        <p>سرکار خانم/جناب آقای <strong>{client}</strong></p>
        <p>پروپوزال طراحی {project} به مساحت {bana} متر مربع واقع در {location}.</p>
    </body>
    </html>
    """
    pdf_buffer = io.BytesIO()
    HTML(string=html_template).write_pdf(pdf_buffer)
    pdf_buffer.seek(0)
    
    # 2. تولید Word در حافظه
    doc = docx.Document()
    doc.add_heading('پروپوزال طراحی ویلا', 0)
    doc.add_paragraph(f'کارفرما: {client}')
    doc.add_paragraph(f'محل پروژه: {location}')
    
    docx_buffer = io.BytesIO()
    doc.save(docx_buffer)
    docx_buffer.seek(0)
    
    return pdf_buffer, docx_buffer

# --- رابط کاربری ---
st.set_page_config(page_title="سیستم صدور پروپوزال", page_icon="📄")
st.markdown("<h2 style='text-align: center;'>سیستم هوشمند صدور پروپوزال</h2>", unsafe_allow_html=True)

with st.form("proposal_form"):
    client_input = st.text_input('نام کارفرما')
    project_input = st.text_input('عنوان پروژه (مثل: ویلای تریبلکس)')
    bana_input = st.number_input('متراژ بنا (مترمربع)', min_value=0, step=10)
    land_input = st.number_input('متراژ زمین/محوطه (مترمربع)', min_value=0, step=10)
    interior_input = st.number_input('متراژ طراحی داخلی', min_value=0, step=10)
    location_input = st.text_input('محل پروژه')
    
    submitted = st.form_submit_button("محاسبه و تولید فایل‌ها")

if submitted:
    if not client_input or not bana_input:
        st.error('لطفاً اطلاعات ضروری را وارد کنید!')
    else:
        with st.spinner('در حال پردازش و تولید فایل‌ها...'):
            pdf_buf, docx_buf = generate_documents(
                client_input, project_input, bana_input, 
                land_input, interior_input, location_input
            )
            
            st.success('✅ فایل‌ها با موفقیت تولید شدند!')
            
            col1, col2 = st.columns(2)
            with col1:
                st.download_button(
                    label="📥 دانلود PDF",
                    data=pdf_buf,
                    file_name=f"Proposal_{client_input}.pdf",
                    mime="application/pdf",
                    use_container_width=True
                )
            with col2:
                st.download_button(
                    label="📥 دانلود Word",
                    data=docx_buf,
                    file_name=f"Proposal_{client_input}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )