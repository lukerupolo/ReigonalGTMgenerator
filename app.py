# app.py
import streamlit as st
import pandas as pd
from pptx import Presentation
from pptx.util import Inches
import io
import json
import requests # Used for making API calls
from xml.etree.ElementTree import fromstring, tostring

# --- Page Configuration ---
st.set_page_config(
    page_title="DreamAI Setups",
    page_icon="🤖",
    layout="wide",
    initial_sidebar_state="expanded",
)

# --- HELPER FUNCTION FOR SLIDE COPYING ---

def copy_slide_from_source(dest_pres, src_slide):
    """
    Copies a slide from a source presentation to the destination presentation.
    This function performs a deep copy of the slide's underlying XML.
    """
    # Create a new blank slide in the destination presentation using the same layout
    # as the source slide. This sets up the basic structure.
    try:
        slide_layout = dest_pres.slide_layouts.get_by_name(src_slide.slide_layout.name)
    except KeyError:
        # Fallback to a standard 'Title and Content' layout if not found
        slide_layout = dest_pres.slide_layouts[1]
        
    new_slide = dest_pres.slides.add_slide(slide_layout)

    # The core of the copy operation: duplicate shapes from source to destination
    for shape in src_slide.shapes:
        # Create a new element from the source shape's XML
        el = shape.element
        new_el = fromstring(tostring(el))
        
        # Add the duplicated element to the new slide's shape tree
        new_slide.shapes._spTree.insert_element_before(new_el, 'p:extLst')

# --- LIVE API FUNCTIONS ---

def call_openai_api(payload, api_key):
    """
    Generic function to call the OpenAI Chat Completions API.
    """
    if not api_key:
        st.error("API Key not found. Please enter your OpenAI API key in the sidebar.")
        return None

    api_url = "https://api.openai.com/v1/chat/completions"
    headers = {
        'Content-Type': 'application/json',
        'Authorization': f'Bearer {api_key}'
    }
    
    try:
        response = requests.post(api_url, headers=headers, json=payload)
        response.raise_for_status() # Raises an exception for bad status codes (4xx or 5xx)
        return response.json()
    except requests.exceptions.RequestException as e:
        st.error(f"API Request Failed: {e}")
        try:
            error_details = e.response.json()
            st.error(f"API Error Details: {error_details.get('error', {}).get('message', 'No details')}")
        except:
            pass
        return None

def get_deep_research(region, api_key):
    """
    Calls the OpenAI API to get structured market research data.
    """
    st.info(f"🤖 Calling OpenAI API for Deep Research on {region}...")
    
    prompt = f"You are a market research analyst. Provide a market analysis for the tech industry in {region}. Return ONLY a valid JSON object with the following keys: 'market_size', 'key_trends' (as a list of strings), 'consumer_behavior', and 'competitor_landscape' (as a list of strings)."
    
    payload = {
        "model": "gpt-3.5-turbo",
        "messages": [{"role": "user", "content": prompt}],
        "response_format": {"type": "json_object"}
    }
    
    result = call_openai_api(payload, api_key)
    
    if result and 'choices' in result:
        try:
            json_text = result['choices'][0]['message']['content']
            return json.loads(json_text)
        except (KeyError, IndexError, json.JSONDecodeError) as e:
            st.error(f"Failed to parse research data from API response: {e}")
            return None
    return None

def get_ai_summary(text_to_summarize, api_key):
    """
    Calls the OpenAI API to get a text summary.
    """
    st.info("🤖 Calling OpenAI API for summarization...")
    
    prompt = f"Summarize the following text in one or two sentences, capturing the key takeaway: '{text_to_summarize}'"
    
    payload = {
        "model": "gpt-3.5-turbo",
        "messages": [{"role": "user", "content": prompt}]
    }
    result = call_openai_api(payload, api_key)
    
    if result and 'choices' in result:
        try:
            return result['choices'][0]['message']['content']
        except (KeyError, IndexError) as e:
            st.error(f"Failed to parse summary from API response: {e}")
            return "Summary could not be generated."
    return "Summary could not be generated."

# --- POWERPOINT GENERATION ---
def create_gtm_presentation(data):
    """
    Creates the final PowerPoint presentation based on user input.
    """
    prs = Presentation()
    source_pres = data.get('source_presentation')
    objective_indices = data.get('objective_slide_indices', [])
    
    # --- Title Slide ---
    title_slide_layout = prs.slide_layouts[0]
    slide = prs.slides.add_slide(title_slide_layout)
    title = slide.shapes.title
    subtitle = slide.placeholders[1]
    title.text = f"GTM Strategy: {data['project_name']}"
    subtitle.text = f"Prepared for {data['region']}"

    # --- Copy Objective Slides ---
    if not objective_indices:
        st.warning("Could not automatically identify an 'Objective' slide. A placeholder slide will be added.")
        slide = prs.slides.add_slide(prs.slide_layouts[1])
        slide.shapes.title.text = "Objectives"
        slide.placeholders[1].text_frame.text = "Objective slide from global deck could not be identified."
    else:
        st.info(f"Copying {len(objective_indices)} objective slide(s)...")
        for idx in objective_indices:
            src_slide = source_pres.slides[idx]
            copy_slide_from_source(prs, src_slide)

    # Use a title and content layout for subsequent slides
    content_slide_layout = prs.slide_layouts[1]
    
    # --- Market Insights Slide ---
    slide = prs.slides.add_slide(content_slide_layout)
    slide.shapes.title.text = f"Market Insights: {data['region']}"
    tf = slide.placeholders[1].text_frame
    tf.text = "Automated Deep Research Insights:"
    api_insights_data = data.get('api_insights', {})
    if api_insights_data:
        for key, value in api_insights_data.items():
            p = tf.add_paragraph(); p.text = f"- {key.replace('_', ' ').title()}: {value}"; p.level = 1
    else:
        p = tf.add_paragraph(); p.text = "No research data generated."; p.level = 1
    tf.add_paragraph() # Add a space
    p = tf.add_paragraph(); p.text = "Custom Regional Insights:"; p.level = 0
    p = tf.add_paragraph(); p.text = data.get('custom_insights', 'Not provided.'); p.level = 1
    
    # --- Timeline Slide ---
    slide = prs.slides.add_slide(content_slide_layout)
    slide.shapes.title.text = "Project Timeline"
    tf = slide.placeholders[1].text_frame
    tf.clear()
    for item in data.get('timeline', []):
        p = tf.add_paragraph(); p.text = item
    
    # --- Activation Slide ---
    slide = prs.slides.add_slide(prs.slide_layouts[5]) # Title Only Layout
    slide.shapes.title.text = "Regional Activation Plan"
    
    left, top, width, height = Inches(0.5), Inches(1.5), Inches(3.0), Inches(5.0)
    txBox = slide.shapes.add_textbox(left, top, width, height); tf = txBox.text_frame
    p = tf.add_paragraph(); p.text = "Insights Summary"; p.font.bold = True
    p = tf.add_paragraph(); p.text = data.get('activation_insights_summary', 'Not generated.')
    
    left = Inches(3.5)
    txBox = slide.shapes.add_textbox(left, top, width, height); tf = txBox.text_frame
    p = tf.add_paragraph(); p.text = "Activation Plan"; p.font.bold = True
    p = tf.add_paragraph(); p.text = data.get('activation_plan', 'Not provided.')

    left = Inches(6.5)
    txBox = slide.shapes.add_textbox(left, top, width, height); tf = txBox.text_frame
    p = tf.add_paragraph(); p.text = "Measurement & KPIs"; p.font.bold = True
    p = tf.add_paragraph(); p.text = data.get('measurement_plan', 'Not provided.')
    
    # --- Investment Slide ---
    slide = prs.slides.add_slide(content_slide_layout)
    slide.shapes.title.text = "Investment Summary"
    body_shape = slide.shapes.placeholders[1]
    try:
        investment_data = data.get('investment_data', [])
        if investment_data:
            df = pd.DataFrame(investment_data)
            rows, cols = df.shape
            table = slide.shapes.add_table(rows+1, cols, Inches(1.5), Inches(2.0), Inches(7), Inches(1.0)).table
            for col_idx, col_name in enumerate(df.columns): table.cell(0, col_idx).text = col_name
            for row_idx in range(rows):
                for col_idx in range(cols): table.cell(row_idx+1, col_idx).text = str(df.iloc[row_idx, col_idx])
        else:
            body_shape.text_frame.text = "No investment data provided."
    except Exception as e:
        body_shape.text_frame.text = f"Could not generate investment table. Error: {e}"

    # --- Overview Slide ---
    slide = prs.slides.add_slide(content_slide_layout)
    slide.shapes.title.text = "AI-Generated Overview"
    slide.placeholders[1].text_frame.text = data.get('overview_summary', 'Not generated.')
    
    ppt_io = io.BytesIO()
    prs.save(ppt_io)
    ppt_io.seek(0)
    return ppt_io

# --- SESSION STATE MANAGEMENT ---
if 'step' not in st.session_state:
    st.session_state.step = 0
    st.session_state.project_data = {}
    st.session_state.api_key = ""

def next_step(): st.session_state.step += 1
def prev_step(): st.session_state.step -= 1

# --- UI LAYOUT ---
st.sidebar.title("DreamAI Setups 🤖")
st.sidebar.markdown("---")
st.sidebar.header("Configuration")
api_key_input = st.sidebar.text_input("Enter Your OpenAI API Key", type="password", key="api_key_input_widget")
if api_key_input:
    st.session_state.api_key = api_key_input
st.sidebar.markdown("---")


# --- Step 0: Welcome & Upload ---
if st.session_state.step == 0:
    st.title("Welcome to DreamAI Setups")
    st.markdown("Automate the creation of regional GTM slide decks. Start by uploading your global GTM presentation.")
    
    if not st.session_state.api_key:
        st.warning("Please enter your OpenAI API Key in the sidebar to begin.")

    project_name = st.text_input("Enter a Project Name:", key="project_name_input")
    region = st.selectbox("Select Target Region:", ["Australia", "Japan", "Korea", "China"], key="region_input")
    uploaded_file = st.file_uploader("Upload Global GTM Deck (.pptx)", type="pptx", key="uploader", disabled=not st.session_state.api_key)

    if st.button("Start Analysis & Build", type="primary", disabled=not st.session_state.api_key):
        if uploaded_file and project_name and region:
            with st.spinner('Analyzing Presentation... This may take a moment.'):
                source_pres = Presentation(uploaded_file)
                
                # Identify objective slides and extract their text content and index
                objective_slide_indices = []
                objective_texts = []
                for i, slide in enumerate(source_pres.slides):
                    if slide.shapes.title and "objective" in slide.shapes.title.text.lower():
                        objective_slide_indices.append(i)
                        full_text = "\n".join(shape.text for shape in slide.shapes if shape.has_text_frame)
                        objective_texts.append(full_text)
                
                st.session_state.project_data = {
                    'project_name': project_name,
                    'region': region,
                    'source_presentation': source_pres,
                    'objective_slide_indices': objective_slide_indices,
                    'objective_texts': "\n---\n".join(objective_texts)
                }
                
                next_step()
                st.rerun()
        else:
            st.error("Please provide a project name, select a region, and upload a .pptx file.")

# --- Multi-Step Form ---
if st.session_state.step > 0:
    st.sidebar.header(f"Project: {st.session_state.project_data.get('project_name', 'New Project')}")
    st.sidebar.markdown(f"**Region:** {st.session_state.project_data.get('region', 'N/A')}")
    st.sidebar.markdown("---")
    
    progress_value = st.session_state.step / 5 # Updated step count
    st.progress(progress_value)

    if st.sidebar.button("↩️ Start Over"):
        st.session_state.step = 0
        st.session_state.project_data = {}
        st.rerun()
    st.sidebar.markdown("---")

# --- Step 1: Objectives Review ---
if st.session_state.step == 1:
    st.header("Step 1: Objectives Review")
    st.subheader("Global Objectives Extracted")
    st.markdown("The following objectives slide(s) were found in the global deck and will be copied exactly into your new presentation.")
    
    extracted_text = st.session_state.project_data.get('objective_texts', 'No objective slide was found.')
    st.text_area("Extracted Text", value=extracted_text, height=250, disabled=True)
    
    st.button("Next: Insights →", type="primary", use_container_width=True, on_click=next_step)

# --- Step 2: Regional Insight Generation ---
if st.session_state.step == 2:
    st.header("Step 2: Regional Insight Generation")
    with st.spinner("AI is conducting research..."):
        api_insights = get_deep_research(st.session_state.project_data['region'], st.session_state.api_key)
    if api_insights:
        st.subheader("Automated Deep Research")
        st.json(api_insights)
    else:
        st.error("Could not fetch AI-powered insights. Please check your API key and try again.")
    st.subheader("Custom Regional Insights")
    custom_insights = st.text_area("Add your own qualitative findings:", height=200, key="custom_insight_input")
    col1, col2, col3 = st.columns([1,1,1])
    with col1: st.button("← Back: Objectives", on_click=prev_step, use_container_width=True)
    with col3:
        if st.button("Save & Next: Timeline →", type="primary", use_container_width=True):
            st.session_state.project_data['api_insights'] = api_insights
            st.session_state.project_data['custom_insights'] = custom_insights
            next_step()
            st.rerun()

# --- Step 3: Timeline Customization ---
if st.session_state.step == 3:
    st.header("Step 3: Timeline Customization")
    st.markdown("Add your region-specific activations and milestones.")
    if 'regional_events' not in st.session_state.project_data:
        st.session_state.project_data['regional_events'] = []
    st.subheader("Add Regional Timeline Events")
    for item in st.session_state.project_data['regional_events']: st.text(f"- {item}")
    new_event = st.text_input("Add a new timeline event (e.g., 'Q3: Local Influencer Campaign')")
    if st.button("Add Event"):
        if new_event:
            st.session_state.project_data['regional_events'].append(new_event)
            st.rerun()

    col1, col2, col3 = st.columns([1,1,1])
    with col1: st.button("← Back: Insights", on_click=prev_step, use_container_width=True)
    with col3:
        if st.button("Save & Next: Activations →", type="primary", use_container_width=True):
            st.session_state.project_data['timeline'] = st.session_state.project_data['regional_events']
            next_step()
            st.rerun()

# --- Step 4: Activation Planning ---
if st.session_state.step == 4:
    st.header("Step 4: Activation Planning")
    st.markdown("Plan your regional activation on a single slide.")
    col1, col2, col3 = st.columns(3)
    with col1:
        st.subheader("Insights Summary")
        with st.spinner("AI is summarizing..."):
            insights_summary = get_ai_summary(st.session_state.project_data.get('custom_insights', ''), st.session_state.api_key)
        st.info(insights_summary)
    with col2:
        st.subheader("Activation Plan")
        activation_plan = st.text_area("Detail your specific marketing activations.", height=300, key="activation_plan_input")
    with col3:
        st.subheader("Measurement (KPIs)")
        measurement_plan = st.text_area("Define how you will measure success.", height=300, key="measurement_plan_input")
    col_back, col_mid, col_next = st.columns([1,1,1])
    with col_back: st.button("← Back: Timeline", on_click=prev_step, use_container_width=True)
    with col_next:
        if st.button("Save & Next: Investment & Export →", type="primary", use_container_width=True):
            st.session_state.project_data['activation_insights_summary'] = insights_summary
            st.session_state.project_data['activation_plan'] = activation_plan
            st.session_state.project_data['measurement_plan'] = measurement_plan
            next_step()
            st.rerun()

# --- Step 5: Final Review & Export ---
if st.session_state.step == 5:
    st.header("Step 5: Final Review & Export")
    st.balloons()
    st.subheader("Investment Summary")
    st.markdown("Provide the regional investment details in CSV format.")
    st.info("Paste budget data below. First line must be headers.\nExample:\nCategory,Q1 Budget,Q2 Budget\nMedia Spend,$50000,$75000")
    investment_data_str = st.text_area("Paste budget data here (CSV format):", height=150, key="investment_input")
    investment_data = []
    if investment_data_str:
        lines = investment_data_str.strip().split('\n')
        if len(lines) > 1:
            try:
                headers = [h.strip() for h in lines[0].split(',')]
                for line in lines[1:]:
                    values = [v.strip() for v in line.split(',')]
                    if len(values) == len(headers): investment_data.append(dict(zip(headers, values)))
            except Exception: st.warning("Could not parse data. Check CSV format.")
    if investment_data: st.dataframe(pd.DataFrame(investment_data))
    
    st.subheader("AI-Generated Overview")
    full_text = f"Project: {st.session_state.project_data.get('project_name')}\nRegion: {st.session_state.project_data.get('region')}\nActivations: {st.session_state.project_data.get('activation_plan')}\nMeasurement: {st.session_state.project_data.get('measurement_plan')}"
    with st.spinner("AI is generating the final overview..."):
        overview_summary = get_ai_summary(full_text, st.session_state.api_key)
    st.text_area("Overview Summary", value=overview_summary, height=150, disabled=True)
    
    if st.button("Generate & Download Presentation", type="primary", use_container_width=True):
        st.session_state.project_data['investment_data'] = investment_data
        st.session_state.project_data['overview_summary'] = overview_summary
        try:
            with st.spinner("Building your PowerPoint file..."):
                ppt_file = create_gtm_presentation(st.session_state.project_data)
            st.download_button(label="✅ Download PowerPoint (.pptx)", data=ppt_file, file_name=f"{st.session_state.project_data['project_name']}_Regional_GTM.pptx", mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")
        except Exception as e:
            st.error(f"An error occurred while generating the presentation: {e}")

    col_back, col_mid = st.columns([1,1])
    with col_back: st.button("← Back: Activations", on_click=prev_step, use_container_width=True)
