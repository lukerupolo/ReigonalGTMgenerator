# app.py
import streamlit as st
import pandas as pd
from pptx import Presentation
from pptx.util import Inches
import io
import json

# --- Page Configuration ---
st.set_page_config(
    page_title="DreamAI Setups",
    page_icon="🤖",
    layout="wide",
    initial_sidebar_state="expanded",
)

# --- App Styling ---
def local_css(file_name):
    with open(file_name) as f:
        st.markdown(f'<style>{f.read()}</style>', unsafe_allow_html=True)

# This is a placeholder for custom CSS if needed.
# For now, we will use Streamlit's default styling which is clean and modern.
# local_css("style.css")


# --- MOCK API FUNCTIONS ---
# In a real application, these would be actual API calls.

def mock_deep_research_api(region):
    """
    Mocks a call to a deep research API.
    Returns structured data for a given region.
    """
    st.info(f"🤖 Calling Deep Research API for {region}...")
    if region == "Japan":
        return {
            "market_size": "¥5.5 Trillion",
            "key_trends": ["Miniaturization", "High-speed connectivity demand", "Aging population tech adoption"],
            "consumer_behavior": "Brand loyalty is high, preference for premium quality and service.",
            "competitor_landscape": ["Sony", "Panasonic", "Rakuten Mobile"]
        }
    elif region == "Australia":
        return {
            "market_size": "AUD $50 Billion",
            "key_trends": ["Growth in e-commerce", "High adoption of renewable energy tech", "Outdoor lifestyle influence"],
            "consumer_behavior": "Price-sensitive but value-driven, strong digital engagement.",
            "competitor_landscape": ["Telstra", "JB Hi-Fi", "Atlassian"]
        }
    else:
        return {
            "market_size": "Data not available",
            "key_trends": ["N/A"],
            "consumer_behavior": "N/A",
            "competitor_landscape": ["N/A"]
        }

def mock_ai_summarizer(text):
    """Mocks an AI call to summarize text."""
    st.info("🤖 AI is summarizing the key points...")
    if text:
        words = text.split()
        summary = " ".join(words[:25]) + "..." if len(words) > 25 else text
        return f"Key takeaway: {summary}"
    return "No text to summarize."

def mock_ppt_style_extractor(ppt_file):
    """Mocks the extraction of styles from a PPT file."""
    st.info("🤖 AI is analyzing the presentation's visual style...")
    # In a real app, this would parse master slides, fonts, colors, etc.
    return {
        "theme_name": "Office Theme (Mock)",
        "primary_color": "RGB(68, 114, 196)",
        "font_family": "Calibri (Mock)",
        "logo_placeholder": "path/to/logo.png"
    }
    
def mock_ppt_content_parser(ppt_file):
    """Mocks the extraction of content from a PPT file."""
    st.info("🤖 AI is parsing the content structure...")
    # In a real app, this would use NLP to find relevant slides.
    return {
        "global_objectives": "Increase market share by 15% globally.\nLaunch Product X in 10 new markets.\nAchieve a 20% growth in online sales.",
        "global_timeline_points": ["Q1: Global Kick-off", "Q2: Product Finalization", "Q3: Marketing Campaign Launch", "Q4: Sales Push"],
        "global_activation_style": "Three-column layout with icons."
    }


# --- POWERPOINT GENERATION ---
def create_gtm_presentation(data):
    """
    Creates the final PowerPoint presentation based on user input.
    """
    prs = Presentation()
    
    # Use a simple title slide layout
    title_slide_layout = prs.slide_layouts[0]
    
    # --- Title Slide ---
    slide = prs.slides.add_slide(title_slide_layout)
    title = slide.shapes.title
    subtitle = slide.placeholders[1]
    title.text = f"GTM Strategy: {data['project_name']}"
    subtitle.text = f"Prepared for {data['region']}"

    # Use a title and content layout for subsequent slides
    content_slide_layout = prs.slide_layouts[1]
    
    # --- Objectives Slide ---
    slide = prs.slides.add_slide(content_slide_layout)
    slide.shapes.title.text = "Regional Objectives"
    body_shape = slide.shapes.placeholders[1]
    tf = body_shape.text_frame
    tf.text = "Global Objectives Alignment:\n"
    p = tf.add_paragraph()
    p.text = data['global_objectives']
    p.level = 1
    tf.add_paragraph() # Spacer
    tf.add_paragraph().text = "Regional Alignments & Additions:"
    p = tf.add_paragraph()
    p.text = data['regional_objectives']
    p.level = 1
    
    # --- Market Insights Slide ---
    slide = prs.slides.add_slide(content_slide_layout)
    slide.shapes.title.text = f"Market Insights: {data['region']}"
    body_shape = slide.shapes.placeholders[1]
    tf = body_shape.text_frame
    tf.text = "Automated Deep Research Insights:"
    for key, value in data['api_insights'].items():
        p = tf.add_paragraph()
        p.text = f"- {key.replace('_', ' ').title()}: {value}"
        p.level = 1
    tf.add_paragraph() # Spacer
    p = tf.add_paragraph()
    p.text = "Custom Regional Insights:"
    p = tf.add_paragraph()
    p.text = data['custom_insights']
    p.level = 1
    
    # --- Timeline Slide ---
    slide = prs.slides.add_slide(content_slide_layout)
    slide.shapes.title.text = "Project Timeline"
    body_shape = slide.shapes.placeholders[1]
    tf = body_shape.text_frame
    tf.clear()
    for item in data['timeline']:
        p = tf.add_paragraph()
        p.text = item
    
    # --- Activation Slide ---
    slide = prs.slides.add_slide(prs.slide_layouts[5]) # Title Only Layout
    slide.shapes.title.text = "Regional Activation Plan"
    
    # Add three text boxes for the layout
    left = Inches(0.5)
    top = Inches(1.5)
    width = Inches(3.0)
    height = Inches(5.0)
    
    # Insights Box
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    p = tf.add_paragraph()
    p.text = "Insights Summary"
    p.font.bold = True
    p = tf.add_paragraph()
    p.text = data['activation_insights_summary']
    
    # Activations Box
    left = Inches(3.5)
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    p = tf.add_paragraph()
    p.text = "Activation Plan"
    p.font.bold = True
    p = tf.add_paragraph()
    p.text = data['activation_plan']

    # Measurement Box
    left = Inches(6.5)
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    p = tf.add_paragraph()
    p.text = "Measurement & KPIs"
    p.font.bold = True
    p = tf.add_paragraph()
    p.text = data['measurement_plan']
    
    # --- Investment Slide ---
    slide = prs.slides.add_slide(content_slide_layout)
    slide.shapes.title.text = "Investment Summary"
    body_shape = slide.shapes.placeholders[1]
    
    try:
        df = pd.DataFrame(data['investment_data'])
        # Add a table to the slide
        rows, cols = df.shape
        table = slide.shapes.add_table(rows+1, cols, Inches(1.5), Inches(2.0), Inches(7), Inches(1.0)).table
        
        # Set column names
        for col_idx, col_name in enumerate(df.columns):
            table.cell(0, col_idx).text = col_name

        # Add data
        for row_idx in range(rows):
            for col_idx in range(cols):
                table.cell(row_idx+1, col_idx).text = str(df.iloc[row_idx, col_idx])
    except Exception as e:
        body_shape.text_frame.text = f"Could not generate investment table. Please check format.\nError: {e}"


    # --- Overview Slide ---
    slide = prs.slides.add_slide(content_slide_layout)
    slide.shapes.title.text = "AI-Generated Overview"
    body_shape = slide.shapes.placeholders[1]
    tf = body_shape.text_frame
    tf.text = data['overview_summary']
    
    # Save to a byte stream
    ppt_io = io.BytesIO()
    prs.save(ppt_io)
    ppt_io.seek(0)
    return ppt_io


# --- SESSION STATE MANAGEMENT ---
if 'step' not in st.session_state:
    st.session_state.step = 0
    st.session_state.project_data = {}

def next_step():
    st.session_state.step += 1
    
def prev_step():
    st.session_state.step -= 1

# --- UI LAYOUT ---

st.sidebar.title("DreamAI Setups 🤖")
st.sidebar.markdown("---")

# --- Step 0: Welcome & Upload ---
if st.session_state.step == 0:
    st.title("Welcome to DreamAI Setups")
    st.markdown("Automate the creation of regional GTM slide decks. Start by uploading your global GTM presentation.")
    
    project_name = st.text_input("Enter a Project Name for this regional deck:", key="project_name_input")
    region = st.selectbox("Select Target Region:", ["Australia", "Japan", "Korea", "China"], key="region_input")
    
    uploaded_file = st.file_uploader("Upload Global GTM Deck (.pptx)", type="pptx", key="uploader")

    if st.button("Start Analysis & Build", type="primary"):
        if uploaded_file and project_name and region:
            with st.spinner('Analyzing Presentation... This may take a moment.'):
                # Store initial info
                st.session_state.project_data['project_name'] = project_name
                st.session_state.project_data['region'] = region
                
                # Mock analysis
                st.session_state.project_data['style'] = mock_ppt_style_extractor(uploaded_file)
                parsed_content = mock_ppt_content_parser(uploaded_file)
                st.session_state.project_data['global_objectives'] = parsed_content['global_objectives']
                st.session_state.project_data['global_timeline_points'] = parsed_content['global_timeline_points']
                
                next_step()
                st.rerun()
        else:
            st.error("Please provide a project name, select a region, and upload a .pptx file.")

# --- Multi-Step Form ---
if st.session_state.step > 0:
    st.sidebar.header(f"Project: {st.session_state.project_data.get('project_name', 'New Project')}")
    st.sidebar.markdown(f"**Region:** {st.session_state.project_data.get('region', 'N/A')}")
    st.sidebar.markdown("---")
    
    # Progress Bar
    progress_value = st.session_state.step / 6
    st.progress(progress_value)

    st.sidebar.page_link("app.py", label="↩️ Start Over", icon="🏠")
    st.sidebar.markdown("---")


# --- Step 1: Objectives Alignment ---
if st.session_state.step == 1:
    st.header("Step 1: Objectives Alignment")
    
    st.subheader("Global Objectives Extracted")
    st.markdown("The following objectives were extracted from the global deck.")
    st.text_area("Global Objectives", value=st.session_state.project_data['global_objectives'], height=150, disabled=True)
    
    st.subheader("Regional Objectives Questionnaire")
    st.markdown("Confirm and adapt these objectives for your region. Add any region-specific goals.")
    
    q1 = st.radio("Do the global objectives fully align with your regional strategy?", ["Yes, completely", "Mostly, with minor additions", "No, significant changes are needed"])
    regional_objectives = st.text_area("Enter your specific regional objectives, priorities, and adaptations here:", height=200, key="regional_obj_input")

    col1, col2 = st.columns([1,1])
    with col2:
        if st.button("Save & Next: Insights →", type="primary", use_container_width=True):
            st.session_state.project_data['regional_objectives'] = regional_objectives
            next_step()
            st.rerun()


# --- Step 2: Regional Insight Generation ---
if st.session_state.step == 2:
    st.header("Step 2: Regional Insight Generation")
    
    st.subheader("Automated Deep Research")
    st.markdown("Based on your selected region, the AI has pulled the following market data.")
    
    api_insights = mock_deep_research_api(st.session_state.project_data['region'])
    st.json(api_insights)
    
    st.subheader("Custom Regional Insights")
    st.markdown("Add your own qualitative findings, customer feedback, or specific market knowledge.")
    custom_insights = st.text_area("Add your custom insights here:", height=200, key="custom_insight_input")

    col1, col2, col3 = st.columns([1,1,1])
    with col1:
        st.button("← Back: Objectives", on_click=prev_step, use_container_width=True)
    with col3:
        if st.button("Save & Next: Timeline →", type="primary", use_container_width=True):
            st.session_state.project_data['api_insights'] = api_insights
            st.session_state.project_data['custom_insights'] = custom_insights
            next_step()
            st.rerun()

# --- Step 3: Timeline Customization ---
if st.session_state.step == 3:
    st.header("Step 3: Timeline Customization")
    st.markdown("The global timeline has been imported. Add your region-specific activations and milestones.")
    
    # Combine global and allow adding regional
    timeline_items = st.session_state.project_data['global_timeline_points'].copy()
    
    st.subheader("Current Timeline")
    
    if 'regional_events' not in st.session_state.project_data:
        st.session_state.project_data['regional_events'] = []

    # Display current items
    for i, item in enumerate(st.session_state.project_data['regional_events']):
        st.text(f"- {item}")

    # Add new item
    new_event = st.text_input("Add a new regional timeline event (e.g., 'Q3: Local Influencer Campaign')")
    if st.button("Add Event"):
        if new_event:
            st.session_state.project_data['regional_events'].append(new_event)
            st.rerun()

    final_timeline = timeline_items + st.session_state.project_data['regional_events']

    st.subheader("Preview of Final Timeline")
    st.expander("Click to see the combined timeline").write(final_timeline)

    col1, col2, col3 = st.columns([1,1,1])
    with col1:
        st.button("← Back: Insights", on_click=prev_step, use_container_width=True)
    with col3:
        if st.button("Save & Next: Activations →", type="primary", use_container_width=True):
            st.session_state.project_data['timeline'] = final_timeline
            next_step()
            st.rerun()


# --- Step 4: Activation Planning ---
if st.session_state.step == 4:
    st.header("Step 4: Activation Planning")
    st.markdown("Plan your regional activation on a single slide. The AI will summarize the insights for you.")

    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.subheader("Insights Summary")
        insights_summary = mock_ai_summarizer(st.session_state.project_data.get('custom_insights', ''))
        st.info(insights_summary)

    with col2:
        st.subheader("Activation Plan")
        activation_plan = st.text_area("Detail your specific marketing activations.", height=300, key="activation_plan_input")

    with col3:
        st.subheader("Measurement (KPIs)")
        measurement_plan = st.text_area("Define how you will measure success.", height=300, key="measurement_plan_input")
        
    col_back, col_mid, col_next = st.columns([1,1,1])
    with col_back:
        st.button("← Back: Timeline", on_click=prev_step, use_container_width=True)
    with col_next:
        if st.button("Save & Next: Investment →", type="primary", use_container_width=True):
            st.session_state.project_data['activation_insights_summary'] = insights_summary
            st.session_state.project_data['activation_plan'] = activation_plan
            st.session_state.project_data['measurement_plan'] = measurement_plan
            next_step()
            st.rerun()

# --- Step 5: Investment Summary ---
if st.session_state.step == 5:
    st.header("Step 5: Investment Summary")
    st.markdown("Provide the regional investment details. You can use a structured format like CSV.")
    
    st.info("Paste your budget data below. The first line should be the headers. Example:\nCategory,Q1 Budget,Q2 Budget\nMedia Spend,$50000,$75000\nEvents,$20000,$10000")
    
    investment_data_str = st.text_area("Paste budget data here (CSV format):", height=200, key="investment_input")
    
    # Basic validation and parsing to dict for PPT function
    investment_data = []
    if investment_data_str:
        lines = investment_data_str.strip().split('\n')
        if len(lines) > 1:
            headers = lines[0].split(',')
            for line in lines[1:]:
                values = line.split(',')
                if len(values) == len(headers):
                    investment_data.append(dict(zip(headers, values)))

    st.subheader("Investment Data Preview")
    if investment_data:
        st.dataframe(pd.DataFrame(investment_data))
    else:
        st.warning("No valid data entered yet.")
        
    col_back, col_mid, col_next = st.columns([1,1,1])
    with col_back:
        st.button("← Back: Activations", on_click=prev_step, use_container_width=True)
    with col_next:
        if st.button("Save & Next: Overview →", type="primary", use_container_width=True):
            st.session_state.project_data['investment_data'] = investment_data
            next_step()
            st.rerun()

# --- Step 6: AI-Powered Overview & Export ---
if st.session_state.step == 6:
    st.header("Step 6: Final Review & Export")
    st.balloons()
    
    st.subheader("AI-Generated Overview")
    st.markdown("The AI has created a summary based on all the information you provided.")
    
    # Create a comprehensive string for the AI to summarize
    full_text = f"Project: {st.session_state.project_data.get('project_name')}\n"
    full_text += f"Region: {st.session_state.project_data.get('region')}\n"
    full_text += f"Objectives: {st.session_state.project_data.get('regional_objectives')}\n"
    full_text += f"Activations: {st.session_state.project_data.get('activation_plan')}\n"
    full_text += f"Measurement: {st.session_state.project_data.get('measurement_plan')}"

    overview_summary = mock_ai_summarizer(full_text)
    st.text_area("Overview Summary", value=overview_summary, height=200, disabled=True)
    st.session_state.project_data['overview_summary'] = overview_summary
    
    st.subheader("Export Your Presentation")
    st.markdown("Your regional GTM deck is ready. Click the button below to download the .pptx file.")
    
    try:
        ppt_file = create_gtm_presentation(st.session_state.project_data)
        st.download_button(
            label="📥 Download PowerPoint (.pptx)",
            data=ppt_file,
            file_name=f"{st.session_state.project_data['project_name']}_Regional_GTM.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            use_container_width=True
        )
    except Exception as e:
        st.error(f"An error occurred while generating the presentation: {e}")
        st.error("Please go back and check your inputs, especially the investment data format.")


    col_back, col_mid = st.columns([1,1])
    with col_back:
        st.button("← Back: Investment", on_click=prev_step, use_container_width=True)

