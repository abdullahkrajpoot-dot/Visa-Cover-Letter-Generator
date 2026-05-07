import streamlit as st
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from io import BytesIO

def generate_cover_letter(data):
    doc = Document()
    
    # Global Font Settings
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Times New Roman'
    font.size = Pt(11)

    # --- Header ---
    header = doc.add_paragraph("To\nThe Visa Officer\nEmbassy of the Kingdom of Belgium\nIslamabad, Pakistan")
    header.paragraph_format.space_after = Pt(12)
    
    # --- Subject ---
    subject = doc.add_paragraph()
    subject_run = subject.add_run("Subject: Application for Belgium Tourist Visa")
    subject_run.bold = True
    subject_run.underline = True
    subject.paragraph_format.space_after = Pt(18)

    # --- Introduction ---
    intro = doc.add_paragraph(
        f"I, {data['full_name']}, son of {data['father_name']}, a Pakistani national, "
        f"holding passport number {data['passport']} and CNIC number {data['cnic']}, "
        f"am applying for a tourist visa for travel from {data['start_date']} to {data['end_date']}. "
        f"The purpose of my visit is tourism during this period."
    )
    intro.paragraph_format.space_after = Pt(12)

    # --- Section Styling Helper ---
    def add_section_heading(text):
        h = doc.add_heading(text, level=2)
        run = h.runs[0]
        run.font.color.rgb = RGBColor(79, 129, 189) # Blue color like screenshot
        run.font.size = Pt(12)
        h.paragraph_format.space_before = Pt(12)
        h.paragraph_format.space_after = Pt(6)

    # --- Purpose of Travel & Detailed Itinerary ---
    add_section_heading('Purpose of Travel')
    doc.add_paragraph(
        "The purpose of my visit to Belgium is tourism, during which I plan to explore the cultural, "
        "historical, and architectural highlights of the country through a well-structured travel itinerary. "
        f"My travel is scheduled from {data['start_date']} to {data['end_date']}, during which I will be primarily "
        "based in Brussels and will also undertake day trips to other major cities."
    )
    
    doc.add_paragraph("In Brussels, I plan to visit:")
    brussels_points = [
        "Grand Place – iconic central square",
        "Galeries Royales Saint-Hubert – historic luxury arcade",
        "Atomium – unique architectural monument",
        "Mini-Europe, Mont des Arts, and Magritte Museum",
        "Parc du Cinquantenaire – large public park with triumphal arch"
    ]
    for p in brussels_points:
        doc.add_paragraph(p, style='List Bullet')

    doc.add_paragraph("My itinerary also includes day trips to:")
    day_trips = [
        "Bruges – 'Venice of the North'",
        "Ghent – Gravensteen Castle",
        "Antwerp – Diamond district",
        "Dinant – Citadel of Dinant"
    ]
    for p in day_trips:
        doc.add_paragraph(p, style='List Bullet')

    doc.add_paragraph(f"After Belgium, I will travel to {data['next_dest']} as per my travel plans.")

    # --- Accommodation Details ---
    add_section_heading('Accommodation Details')
    doc.add_paragraph(
        f"Hotel: {data['hotel_name']}\n"
        f"Address: {data['hotel_address']}\n"
        f"Contact: {data['hotel_contact']}"
    )

    # --- Financial Arrangements ---
    add_section_heading('Financial Arrangements')
    doc.add_paragraph(
        "All travel expenses will be self-funded. Attached are my bank statements (reflecting financial "
        "stability) and tax returns for the recent years."
    )

    # --- Business & Professional Background ---
    add_section_heading('Business & Professional Background')
    doc.add_paragraph(
        f"I am the owner of {data['business_name']}, a construction enterprise registered with FBR "
        f"(NTN: {data['ntn']}). I am an active member of the Islamabad Chamber of Commerce & Industry."
    )

    # --- Family & Ties to Pakistan ---
    add_section_heading('Family & Ties to Pakistan')
    doc.add_paragraph(data['family_details'])

    # --- Travel History ---
    add_section_heading('Travel History')
    doc.add_paragraph(data['travel_history'])

    # --- Declaration & Request ---
    add_section_heading('Declaration & Request')
    doc.add_paragraph(
        "I affirm my commitment to strictly adhering to all Belgian laws and Schengen regulations. "
        "I kindly request the issuance of a short-stay tourist visa."
    )

    # --- Closing ---
    doc.add_paragraph("\nYour Sincerely,")
    final = doc.add_paragraph(f"{data['full_name']}\nContact: {data['contact']}")
    
    bio = BytesIO()
    doc.save(bio)
    return bio.getvalue()

# --- Streamlit UI ---
st.set_page_config(page_title="Pro Visa Gen", layout="wide")
st.title("📄 Ali Baba Travel Advisor - Professional Cover Letter")

with st.form("visa_form"):
    col1, col2 = st.columns(2)
    with col1:
        st.subheader("👤 Personal Details")
        name = st.text_input("Full Name", "Rana Muhammad Asif")
        fname = st.text_input("Father Name", "Rana Abdul Majeed")
        passp = st.text_input("Passport No", "BJ1880214")
        cnic_no = st.text_input("CNIC No", "34401-44150217")
        biz = st.text_input("Business Name", "RANA BUILDERS")
        ntn_no = st.text_input("NTN", "2815574-2")

    with col2:
        st.subheader("✈️ Travel Details")
        s_date = st.text_input("Start Date", "13 June 2026")
        e_date = st.text_input("End Date", "22 June 2026")
        next_d = st.text_input("Next Destination (if any)", "the USA from 22 June to 27 June 2026")
        h_name = st.text_input("Hotel", "Dolce by Wyndham La Hulpe Brussels")
        h_addr = st.text_area("Hotel Address", "135 Chaussée de Bruxelles, 1310 La Hulpe, Belgium")
        h_cont = st.text_input("Hotel Contact", "+32 2 290 9800")

    st.subheader("📋 Additional Information")
    fam = st.text_area("Family Ties", "I am a married individual and a father of two children, residing in Islamabad. My business and family commitments ensure my return.")
    hist = st.text_area("Detailed Travel History", "Saudi Arabia (3 visits), Malaysia (2), Turkey (2), UK (2), Egypt (2), and single visits to USA, Belgium, Spain, Greece, Japan.")
    
    submit = st.form_submit_button("Generate Long-Form Letter")

if submit:
    user_data = {
        "full_name": name, "father_name": fname, "passport": passp, "cnic": cnic_no,
        "start_date": s_date, "end_date": e_date, "business_name": biz, "ntn": ntn_no,
        "hotel_name": h_name, "hotel_address": h_addr, "hotel_contact": h_cont,
        "family_details": fam, "travel_history": hist, "next_dest": next_d,
        "contact": "+92 3007799500"
    }
    file = generate_cover_letter(user_data)
    st.success("Professional Letter Generated Successfully!")
    st.download_button("📥 Download Final Document", file, f"Visa_Letter_{name}.docx")
