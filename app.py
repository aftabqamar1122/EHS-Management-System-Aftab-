import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from datetime import date, datetime
import os
import time
import uuid
from io import BytesIO
from PIL import Image
import json
import warnings
warnings.filterwarnings('ignore')

# ============================================
# CONFIGURATION
# ============================================

# Page config
st.set_page_config(
    layout="wide",
    page_title="EHS Management System",
    page_icon="🛡️",
    initial_sidebar_state="expanded"
)

# Create directories for storing files
BASE_UPLOAD_DIR = "uploads"
DIRECTORIES = {
    "injury_photos": os.path.join(BASE_UPLOAD_DIR, "injury_photos"),
    "fir_documents": os.path.join(BASE_UPLOAD_DIR, "fir_documents"),
    "meeting_minutes": os.path.join(BASE_UPLOAD_DIR, "meeting_minutes"),
    "cost_bills": os.path.join(BASE_UPLOAD_DIR, "cost_bills"),
    "aspect_images": os.path.join(BASE_UPLOAD_DIR, "aspect_images")
}

for directory in DIRECTORIES.values():
    os.makedirs(directory, exist_ok=True)

# File paths
INCIDENT_FILE = "incident_sheet.xlsx"
ASPECT_FILE = "aspect_impact_sheet.xlsx"
LOGIN_FILE = "login_data.xlsx"

# ============================================
# AUTHENTICATION - UPDATED USERS
# ============================================

# User credentials with your custom users
USERS = {
    "Aftab": {"password": "Aftab1122", "role": "admin"},
    "Mukesh": {"password": "Mukesh1122", "role": "user"},
    "Himanshu": {"password": "Himanshu1122", "role": "user"},
    "Sunil": {"password": "Sunil1122", "role": "admin"},
    "Zeidan": {"password": "Zeidan1122", "role": "admin"}
}

EDITOR_ROLES = ["admin"]

def initialize_login_file():
    """Initialize login tracking file"""
    if not os.path.exists(LOGIN_FILE):
        df = pd.DataFrame(columns=['Username', 'Login Time', 'Action'])
        df.to_excel(LOGIN_FILE, index=False)

def log_activity(username, action):
    """Log user activity"""
    try:
        df = pd.read_excel(LOGIN_FILE)
    except:
        df = pd.DataFrame(columns=['Username', 'Login Time', 'Action'])
    
    new_entry = pd.DataFrame({
        'Username': [username],
        'Login Time': [datetime.now().strftime("%Y-%m-%d %H:%M:%S")],
        'Action': [action]
    })
    
    df = pd.concat([df, new_entry], ignore_index=True)
    df.to_excel(LOGIN_FILE, index=False)

def verify_login(username, password):
    """Verify user credentials"""
    return username in USERS and USERS[username]["password"] == password

def is_editor(username):
    """Check if user has edit permissions"""
    return username in USERS and USERS[username]["role"] in EDITOR_ROLES

# ============================================
# DATABASE FUNCTIONS
# ============================================

def initialize_excel_files():
    """Initialize Excel files with proper sheets"""
    files = {
        INCIDENT_FILE: ['Create', 'Approver', 'Injury', 'Investigation', 
                       'RootCause', 'CA_PA', 'Costing', 'Closure'],
        ASPECT_FILE: ['Activity', 'AspectImpact', 'ScoreCard', 'CAPA', 'Report']
    }
    
    for file_path, sheets in files.items():
        if not os.path.exists(file_path):
            with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                for sheet in sheets:
                    pd.DataFrame().to_excel(writer, sheet_name=sheet, index=False)

def load_sheet(file_path, sheet_name):
    """Load a specific sheet from Excel file"""
    try:
        return pd.read_excel(file_path, sheet_name=sheet_name)
    except:
        return pd.DataFrame()

def save_sheet(file_path, sheet_name, df):
    """Save a DataFrame to a specific sheet in Excel file"""
    # Load all existing sheets
    sheets_data = {}
    try:
        with pd.ExcelFile(file_path) as xls:
            for sheet in xls.sheet_names:
                if sheet != sheet_name:
                    sheets_data[sheet] = pd.read_excel(xls, sheet_name=sheet)
    except:
        pass
    
    # Add/update the target sheet
    sheets_data[sheet_name] = df
    
    # Write all sheets back
    with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
        for s_name, s_df in sheets_data.items():
            s_df.to_excel(writer, sheet_name=s_name, index=False)

def save_uploaded_file(file, directory, prefix=""):
    """Save an uploaded file and return the path"""
    if file is None:
        return ""
    
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    unique_id = str(uuid.uuid4())[:8]
    file_ext = os.path.splitext(file.name)[1]
    filename = f"{prefix}_{timestamp}_{unique_id}{file_ext}" if prefix else f"{timestamp}_{unique_id}{file_ext}"
    filepath = os.path.join(directory, filename)
    
    with open(filepath, "wb") as f:
        f.write(file.getbuffer())
    
    return filepath

def save_multiple_files(files, directory, prefix=""):
    """Save multiple uploaded files and return list of paths"""
    saved_paths = []
    if files:
        for i, file in enumerate(files):
            path = save_uploaded_file(file, directory, f"{prefix}_{i+1}")
            if path:
                saved_paths.append(path)
    return saved_paths

# ============================================
# FISHBONE DIAGRAM
# ============================================

def generate_fishbone_diagram(cause_type, causes):
    """Generate an interactive fishbone diagram using Plotly"""
    fig = go.Figure()
    
    # Main spine
    fig.add_shape(
        type="line",
        x0=0, y0=0,
        x1=1, y1=0,
        line=dict(color="black", width=3)
    )
    
    # Incident box
    fig.add_annotation(
        x=1.02, y=0,
        text="Incident",
        showarrow=False,
        font=dict(size=16, color="black"),
        bgcolor="lightblue",
        bordercolor="black",
        borderwidth=2,
        borderpad=6
    )
    
    categories = list(causes.keys())
    num_categories = len(categories)
    
    if num_categories > 0:
        spacing = 1.0 / (num_categories + 1)
        
        for i, category in enumerate(categories):
            pos = (i + 1) * spacing
            
            # Category spine
            fig.add_shape(
                type="line",
                x0=pos, y0=-0.4,
                x1=pos, y1=0,
                line=dict(color="black", width=2)
            )
            
            # Category label
            fig.add_annotation(
                x=pos, y=-0.45,
                text=category,
                showarrow=False,
                font=dict(size=12, color="black")
            )
            
            # Causes
            cause_list = causes[category]
            if cause_list:
                num_causes = len(cause_list)
                for j, cause in enumerate(cause_list):
                    cause_y = -0.4 * (j + 1) / (num_causes + 1)
                    fig.add_annotation(
                        x=pos + 0.05, y=cause_y,
                        text=f"• {cause}",
                        showarrow=False,
                        xanchor="left",
                        font=dict(size=10)
                    )
    
    fig.update_layout(
        title=f"{cause_type} Root Cause Analysis",
        showlegend=False,
        xaxis=dict(
            showticklabels=False,
            showgrid=False,
            zeroline=False,
            range=[-0.1, 1.2]
        ),
        yaxis=dict(
            showticklabels=False,
            showgrid=False,
            zeroline=False,
            range=[-0.6, 0.6]
        ),
        plot_bgcolor="white",
        height=500,
        margin=dict(l=20, r=20, t=40, b=40)
    )
    
    return fig

# ============================================
# VISUALIZATIONS
# ============================================

def create_incident_pie_chart(df):
    """Create pie chart for incident types"""
    if df.empty or 'Incident_type' not in df.columns:
        return None
    
    incident_counts = df['Incident_type'].value_counts()
    if incident_counts.empty:
        return None
    
    fig = go.Figure(data=[go.Pie(
        labels=incident_counts.index,
        values=incident_counts.values,
        hole=.3,
        marker_colors=['#1f77b4', '#ff7f0e', '#2ca02c', '#d62728', 
                       '#9467bd', '#8c564b', '#e377c2', '#7f7f7f']
    )])
    
    fig.update_layout(
        title="Incident Distribution by Type",
        height=400,
        margin=dict(l=10, r=10, t=50, b=10)
    )
    
    return fig

def create_severity_bar_chart(df):
    """Create bar chart for severity distribution"""
    if df.empty or 'Injury_status' not in df.columns:
        return None
    
    severity_order = ["None", "Minor", "Major", "Fatal"]
    severity_counts = df['Injury_status'].value_counts().reindex(severity_order, fill_value=0)
    
    if severity_counts.sum() == 0:
        return None
    
    colors = {
        'None': '#2ca02c',
        'Minor': '#ffbb78',
        'Major': '#ff7f0e',
        'Fatal': '#d62728'
    }
    bar_colors = [colors.get(s, '#1f77b4') for s in severity_counts.index]
    
    fig = go.Figure(data=[go.Bar(
        x=severity_counts.index,
        y=severity_counts.values,
        marker_color=bar_colors
    )])
    
    fig.update_layout(
        title="Incident Severity Distribution",
        xaxis_title="Severity Level",
        yaxis_title="Number of Incidents",
        height=400,
        margin=dict(l=10, r=10, t=50, b=10)
    )
    
    return fig

def create_trend_chart(df):
    """Create line chart for incident trends over time"""
    if df.empty or 'Incident Start Date' not in df.columns:
        return None
    
    try:
        df['Incident Start Date'] = pd.to_datetime(df['Incident Start Date'], errors='coerce')
        df = df.dropna(subset=['Incident Start Date'])
        
        if df.empty:
            return None
        
        df['Month'] = df['Incident Start Date'].dt.to_period('M')
        monthly = df.groupby('Month').size().reset_index(name='Count')
        monthly['Month'] = monthly['Month'].astype(str)
        monthly = monthly.sort_values('Month')
        
        fig = go.Figure()
        fig.add_trace(go.Scatter(
            x=monthly['Month'],
            y=monthly['Count'],
            mode='lines+markers',
            name='Incidents',
            line=dict(color='#1f77b4', width=3),
            marker=dict(size=8)
        ))
        
        fig.update_layout(
            title="Monthly Incident Trends",
            xaxis_title="Month",
            yaxis_title="Number of Incidents",
            height=400,
            margin=dict(l=10, r=10, t=50, b=10)
        )
        
        return fig
    except:
        return None

def create_cost_chart(df):
    """Create horizontal bar chart for cost breakdown"""
    if df.empty:
        return None
    
    numeric_cols = df.select_dtypes(include=['int64', 'float64']).columns
    cost_columns = [col for col in numeric_cols 
                   if col not in ['Total Cost', 'submission_id'] 
                   and not col.endswith('_docs')]
    
    if not cost_columns:
        return None
    
    cost_totals = df[cost_columns].sum().sort_values(ascending=True)
    
    if cost_totals.sum() == 0:
        return None
    
    fig = go.Figure()
    fig.add_trace(go.Bar(
        y=cost_totals.index,
        x=cost_totals.values,
        orientation='h',
        marker=dict(
            color='rgba(50, 171, 96, 0.6)',
            line=dict(color='rgba(50, 171, 96, 1.0)', width=1)
        )
    ))
    
    fig.update_layout(
        title="Cost Breakdown by Category",
        xaxis_title="Amount (₹)",
        yaxis_title="Cost Category",
        height=500,
        margin=dict(l=10, r=10, t=50, b=10)
    )
    
    fig.update_xaxes(tickprefix="₹")
    
    return fig

# ============================================
# MAIN APP
# ============================================

def main():
    """Main application entry point"""
    
    # Initialize
    initialize_login_file()
    initialize_excel_files()
    
    # Session state
    if 'logged_in' not in st.session_state:
        st.session_state.logged_in = False
    if 'current_page' not in st.session_state:
        st.session_state.current_page = None
    if 'username' not in st.session_state:
        st.session_state.username = None
    if 'incidents' not in st.session_state:
        st.session_state.incidents = {}
    if 'investigation_data' not in st.session_state:
        st.session_state.investigation_data = {}
    if 'root_cause_data' not in st.session_state:
        st.session_state.root_cause_data = {}
    if 'current_assessment_id' not in st.session_state:
        st.session_state.current_assessment_id = None
    if 'edit_mode' not in st.session_state:
        st.session_state.edit_mode = False
    
    # ============================================
    # LOGIN PAGE
    # ============================================
    
    if not st.session_state.logged_in:
        st.title("🛡️ EHS Management System")
        st.subheader("Environment, Health & Safety")
        
        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            with st.container():
                # Try to load logo if exists
                try:
                    st.image("logo1.jpg", width=200) if os.path.exists("logo1.jpg") else None
                except:
                    pass
                st.markdown("---")
                st.subheader("🔐 Please Login")
                
                with st.form("login_form"):
                    username = st.text_input("Username", placeholder="Enter your username")
                    password = st.text_input("Password", type="password", placeholder="Enter your password")
                    
                    if st.form_submit_button("Login", use_container_width=True):
                        if verify_login(username, password):
                            st.session_state.logged_in = True
                            st.session_state.username = username
                            log_activity(username, "Login")
                            st.success("✅ Login successful!")
                            st.rerun()
                        else:
                            st.error("❌ Invalid username or password!")
        
        return
    
    # ============================================
    # MAIN APP - Logged In
    # ============================================
    
    # Sidebar
    with st.sidebar:
        st.title("🛡️ EHS System")
        st.markdown(f"**Welcome, {st.session_state.username}**")
        st.markdown(f"Role: {USERS[st.session_state.username]['role'].title()}")
        st.markdown("---")
        
        # Navigation
        nav_options = [
            ("📋 Incident Management", 1),
            ("📊 Aspect/Impact", 2),
            ("🗑️ Waste", 3),
            ("📝 Audit", 4),
            ("🔬 HAZOP", 5),
            ("🎓 Training", 6),
            ("📄 Permit", 7),
            ("⚠️ Risk", 8),
            ("🏥 Occupational Health", 9),
            ("🌍 Carbon Accounting", 10),
            ("⚡ Consumption", 11),
            ("🌿 Conservation", 12)
        ]
        
        for label, page_id in nav_options:
            if st.button(label, use_container_width=True):
                st.session_state.current_page = page_id
                st.rerun()
        
        st.markdown("---")
        if st.button("🚪 Logout", use_container_width=True):
            log_activity(st.session_state.username, "Logout")
            st.session_state.logged_in = False
            st.session_state.username = None
            st.rerun()
    
    # ============================================
    # PAGE ROUTING
    # ============================================
    
    page = st.session_state.current_page
    
    if page == 1:
        incident_management_page()
    elif page == 2:
        aspect_impact_page()
    elif page == 3:
        placeholder_page("🗑️ Waste Management", 
                        "Track and manage waste generation, disposal, and recycling.")
    elif page == 4:
        placeholder_page("📝 Audit Management",
                        "Schedule and track EHS audits, findings, and corrective actions.")
    elif page == 5:
        placeholder_page("🔬 HAZOP Studies",
                        "Conduct Hazard and Operability studies for process safety.")
    elif page == 6:
        placeholder_page("🎓 Training Management",
                        "Track employee EHS training, certifications, and compliance.")
    elif page == 7:
        placeholder_page("📄 Permit Management",
                        "Manage work permits, hot work permits, and confined space entries.")
    elif page == 8:
        placeholder_page("⚠️ Risk Assessment",
                        "Identify and evaluate workplace risks and hazards.")
    elif page == 9:
        placeholder_page("🏥 Occupational Health",
                        "Manage employee health records, medical exams, and wellness.")
    elif page == 10:
        placeholder_page("🌍 Carbon Accounting",
                        "Track carbon emissions, offsets, and sustainability metrics.")
    elif page == 11:
        placeholder_page("⚡ Consumption in Facility",
                        "Monitor resource consumption: energy, water, materials.")
    elif page == 12:
        placeholder_page("🌿 Conservation in Facility",
                        "Track conservation initiatives and sustainability efforts.")
    else:
        # Dashboard
        dashboard_page()

# ============================================
# INCIDENT MANAGEMENT PAGE
# ============================================

def incident_management_page():
    """Incident Management module"""
    st.title("📋 Incident Management")
    
    tabs = st.tabs([
        "Create", "Approver", "Injury", "Investigation", 
        "Root Cause", "CA/PA", "Costing", "Closure", "Report"
    ])
    
    # Tab 1: Create
    with tabs[0]:
        st.header("Create Incident")
        with st.form("create_form"):
            col1, col2 = st.columns(2)
            
            with col1:
                company_name = st.text_input("Company Name")
                department_name = st.text_input("Department Name")
                incident_number = st.text_input("Incident Number")
                incident_location = st.text_input("Incident Location")
                
            with col2:
                incident_type = st.selectbox("Incident Type", 
                    ["Injury", "Near Miss", "Property Damage", "Environmental", 
                     "Fire", "Security", "Other"])
                involved_persons = st.number_input("No. of Involved Persons", min_value=0)
                shift = st.selectbox("Shift (Duty Timings)", 
                    ["", "Morning (6:00 AM - 2:00 PM)", 
                     "Afternoon (2:00 PM - 10:00 PM)", 
                     "Night (10:00 PM - 6:00 AM)"])
            
            description = st.text_area("Description", height=100)
            incident_date = st.date_input("Incident Start Date", value=date.today())
            
            if st.form_submit_button("Submit Incident", use_container_width=True):
                data = {
                    "Company Name": company_name,
                    "Department Name": department_name,
                    "Incident Number": incident_number,
                    "Incident Location": incident_location,
                    "Incident_type": incident_type,
                    "No. of Involved Persons": involved_persons,
                    "Shift (Duty Timings)": shift,
                    "Description": description,
                    "Incident Start Date": incident_date
                }
                
                df = load_sheet(INCIDENT_FILE, 'Create')
                df = pd.concat([df, pd.DataFrame([data])], ignore_index=True)
                save_sheet(INCIDENT_FILE, 'Create', df)
                
                st.session_state.incidents[incident_number] = {"status": "Pending"}
                st.success("✅ Incident Created Successfully!")
                log_activity(st.session_state.username, f"Created incident: {incident_number}")
    
    # Tab 2: Approver
    with tabs[1]:
        st.header("Approver")
        
        df = load_sheet(INCIDENT_FILE, 'Create')
        if df.empty:
            st.info("No incidents available for approval.")
        else:
            incident_to_approve = st.selectbox(
                "Select Incident to Approve",
                df['Incident Number'].tolist() if 'Incident Number' in df.columns else []
            )
            
            if incident_to_approve:
                incident_data = df[df['Incident Number'] == incident_to_approve].iloc[0]
                st.dataframe(pd.DataFrame([incident_data]))
                
                if st.button("✅ Check & Approve", use_container_width=True):
                    st.session_state.incidents[incident_to_approve] = {"status": "Approved"}
                    log_activity(st.session_state.username, f"Approved incident: {incident_to_approve}")
                    st.success(f"Incident {incident_to_approve} approved successfully!")
    
    # Tab 3: Injury
    with tabs[2]:
        st.header("Injury Documentation")
        
        with st.form("injury_form"):
            col1, col2 = st.columns(2)
            
            with col1:
                emp_code = st.text_input("Employee Code")
                emp_name = st.text_input("Employee Name")
                incident_number = st.text_input("Incident Number")
                injury = st.selectbox("Injury Occurred?", ["Yes", "No"])
                
            with col2:
                injury_status = st.selectbox("Injury Status", ["None", "Minor", "Major", "Fatal"])
                investigation_status = st.text_input("Investigation Status")
                root_cause = st.text_input("Root Cause")
            
            injury_info = st.text_area("Injury Information", height=80)
            
            # Body parts
            body_parts_options = [
                "Head", "Face", "Eye", "Ear", "Nose", "Mouth", "Neck", "Shoulder",
                "Arm", "Elbow", "Wrist", "Hand", "Fingers", "Chest", "Back", "Abdomen",
                "Hip", "Leg", "Knee", "Ankle", "Foot", "Toes", "Internal Organs", "Multiple"
            ]
            selected_body_parts = st.multiselect("Injured Body Parts", body_parts_options)
            
            # FIR Details
            st.subheader("FIR Details")
            col1, col2 = st.columns(2)
            with col1:
                fir_number = st.text_input("FIR Number")
                fir_date = st.date_input("FIR Date")
            with col2:
                fir_police_station = st.text_input("Police Station")
            fir_details = st.text_area("FIR Brief Description", height=60)
            
            # File uploads
            st.subheader("Upload Documents")
            injury_photos = st.file_uploader("Injury Photos", 
                                            type=["jpg", "jpeg", "png"], 
                                            accept_multiple_files=True)
            fir_doc = st.file_uploader("FIR Document", 
                                      type=["jpg", "jpeg", "png", "pdf"])
            
            if st.form_submit_button("Submit Injury Info", use_container_width=True):
                # Save files
                photo_paths = save_multiple_files(injury_photos, DIRECTORIES["injury_photos"], incident_number)
                fir_path = save_uploaded_file(fir_doc, DIRECTORIES["fir_documents"], incident_number)
                
                data = {
                    "Emp. Code": emp_code,
                    "Emp. Name": emp_name,
                    "Incident Number": incident_number,
                    "Injury": injury,
                    "Injury_status": injury_status,
                    "Injury Information": injury_info,
                    "body_parts": ", ".join(selected_body_parts),
                    "Investigation": investigation_status,
                    "Root Cause": root_cause,
                    "fir_number": fir_number,
                    "fir_date": fir_date,
                    "fir_police_station": fir_police_station,
                    "fir_details": fir_details,
                    "injury_photos": ",".join(photo_paths),
                    "fir_doc_path": fir_path
                }
                
                df = load_sheet(INCIDENT_FILE, 'Injury')
                df = pd.concat([df, pd.DataFrame([data])], ignore_index=True)
                save_sheet(INCIDENT_FILE, 'Injury', df)
                
                st.success("✅ Injury information saved successfully!")
                log_activity(st.session_state.username, f"Added injury for incident: {incident_number}")
    
    # Tab 4: Investigation
    with tabs[3]:
        st.header("Investigation")
        
        approved_incidents = {k: v for k, v in st.session_state.incidents.items() 
                             if v.get("status") == "Approved"}
        
        if not approved_incidents:
            st.info("No approved incidents available for investigation.")
        else:
            selected = st.selectbox("Select Incident", list(approved_incidents.keys()))
            
            if selected:
                with st.form("investigation_form"):
                    col1, col2 = st.columns(2)
                    
                    with col1:
                        investigation_id = st.text_input("Investigation ID", 
                                                         value=f"INV-{selected}")
                        investigation_subject = st.text_input("Investigation Subject")
                    with col2:
                        investigation_date = st.date_input("Start Date", value=date.today())
                    
                    investigation_desc = st.text_area("Investigation Description", height=100)
                    not_applicable = st.checkbox("Not Applicable")
                    
                    if st.form_submit_button("Save Investigation", use_container_width=True):
                        data = {
                            "Incident ID": selected,
                            "Investigation ID": investigation_id,
                            "Subject": investigation_subject,
                            "Description": investigation_desc,
                            "Start Date": investigation_date,
                            "Not Applicable": not_applicable
                        }
                        
                        df = load_sheet(INCIDENT_FILE, 'Investigation')
                        if not df.empty and "Incident ID" in df.columns:
                            df = df[df["Incident ID"] != selected]
                        df = pd.concat([df, pd.DataFrame([data])], ignore_index=True)
                        save_sheet(INCIDENT_FILE, 'Investigation', df)
                        
                        st.session_state.investigation_data[selected] = data
                        st.success("✅ Investigation details saved!")
                        log_activity(st.session_state.username, f"Added investigation for: {selected}")
    
    # Tab 5: Root Cause
    with tabs[4]:
        st.header("Root Cause Analysis")
        
        investigated = [k for k in st.session_state.investigation_data.keys()]
        
        if not investigated:
            st.info("No incidents with investigation data available.")
        else:
            selected = st.selectbox("Select Incident", investigated, key="rc_select")
            
            if selected:
                cause_type = st.radio("Root Cause Type", ["Internal", "External"])
                
                if cause_type == "Internal":
                    categories = ["People", "Process", "Equipment", "Environment"]
                else:
                    categories = ["External Factors", "Regulatory", "Third-party", "Natural"]
                
                root_data = st.session_state.root_cause_data.get(selected, {"causes": {}})
                if "causes" not in root_data or set(root_data["causes"].keys()) != set(categories):
                    root_data["causes"] = {cat: [] for cat in categories}
                
                # Display and manage causes
                for category in categories:
                    st.subheader(f"{category} Factors")
                    
                    # Existing causes
                    if category in root_data["causes"]:
                        for i, cause in enumerate(root_data["causes"][category]):
                            col1, col2 = st.columns([5, 1])
                            with col1:
                                st.text_input(f"{category} Factor {i+1}", value=cause, 
                                             key=f"rc_{category}_{i}_{selected}", disabled=True)
                            with col2:
                                if st.button("🗑️", key=f"del_{category}_{i}_{selected}"):
                                    root_data["causes"][category].pop(i)
                                    st.session_state.root_cause_data[selected] = root_data
                                    st.rerun()
                    
                    # Add new cause
                    new_cause = st.text_input(f"Add {category} Factor", 
                                             key=f"new_{category}_{selected}")
                    if st.button(f"Add to {category}", key=f"add_{category}_{selected}"):
                        if new_cause:
                            root_data["causes"][category].append(new_cause)
                            st.session_state.root_cause_data[selected] = root_data
                            st.rerun()
                
                # Fishbone diagram
                has_causes = any(causes for causes in root_data["causes"].values() if causes)
                if has_causes:
                    st.subheader("Fishbone Analysis Diagram")
                    fig = generate_fishbone_diagram(cause_type, root_data["causes"])
                    st.plotly_chart(fig, use_container_width=True)
                
                if st.button("💾 Save Root Cause Analysis", use_container_width=True):
                    root_data["type"] = cause_type
                    st.session_state.root_cause_data[selected] = root_data
                    
                    excel_data = {"Incident ID": selected, "Type": cause_type}
                    for category, causes in root_data["causes"].items():
                        for i, cause in enumerate(causes):
                            excel_data[f"{category}_Cause_{i+1}"] = cause
                    
                    df = load_sheet(INCIDENT_FILE, 'RootCause')
                    if not df.empty and "Incident ID" in df.columns:
                        df = df[df["Incident ID"] != selected]
                    df = pd.concat([df, pd.DataFrame([excel_data])], ignore_index=True)
                    save_sheet(INCIDENT_FILE, 'RootCause', df)
                    
                    st.success("✅ Root cause analysis saved!")
                    log_activity(st.session_state.username, f"Added root cause for: {selected}")
    
    # Tab 6: CA/PA
    with tabs[5]:
        st.header("Corrective / Preventive Action")
        
        with st.form("capa_form"):
            col1, col2 = st.columns(2)
            
            with col1:
                incident_number = st.text_input("Incident Number")
                ca_number = st.text_input("CA Number", value="CA001")
                priority = st.selectbox("Priority", ["High", "Medium", "Low"])
                owner_id = st.text_input("Owner ID")
                
            with col2:
                start_date = st.date_input("Start Date")
                end_date = st.date_input("End Date")
                status = st.selectbox("Status", ["Open", "In Progress", "Closed"])
                mail_ids = st.text_input("Mail IDs (Owner, Implementer)")
            
            proposed_ca = st.text_area("Proposed CA")
            implemented_ca = st.text_area("Implemented CA")
            implementer_info = st.text_input("Implementer Info")
            owner_comments = st.text_area("Owner Comments")
            implementer_comments = st.text_area("Implementer Comments")
            
            st.subheader("Meeting Documentation")
            mom_docs = st.file_uploader("Upload Minutes of Meeting", 
                                       type=["jpg", "jpeg", "png", "pdf"],
                                       accept_multiple_files=True)
            
            if st.form_submit_button("Submit CA/PA", use_container_width=True):
                mom_paths = save_multiple_files(mom_docs, DIRECTORIES["meeting_minutes"], incident_number)
                
                data = {
                    "Incident Number": incident_number,
                    "CA Number": ca_number,
                    "Priority": priority,
                    "Owner ID": owner_id,
                    "Start Date": start_date,
                    "End Date": end_date,
                    "Status": status,
                    "Mail IDs": mail_ids,
                    "Proposed CA": proposed_ca,
                    "Implemented CA": implemented_ca,
                    "Implementer Info": implementer_info,
                    "Owner Comments": owner_comments,
                    "Implementer Comments": implementer_comments,
                    "mom_doc_paths": ",".join(mom_paths)
                }
                
                df = load_sheet(INCIDENT_FILE, 'CA_PA')
                df = pd.concat([df, pd.DataFrame([data])], ignore_index=True)
                save_sheet(INCIDENT_FILE, 'CA_PA', df)
                
                st.success("✅ CA/PA submitted successfully!")
                log_activity(st.session_state.username, f"Added CA/PA for incident: {incident_number}")
    
    # Tab 7: Costing
    with tabs[6]:
        st.header("Costing")
        
        with st.form("costing_form"):
            cost_fields = [
                "Medical Bills", "Injury Pay", "Transport", "Compensation", 
                "Man Hours Lost", "Machine Downtime", "Damages", 
                "Implementation Costs", "Miscellaneous"
            ]
            
            costs = {}
            cols = st.columns(3)
            for i, field in enumerate(cost_fields):
                with cols[i % 3]:
                    costs[field] = st.number_input(field, min_value=0.0, 
                                                   value=0.0, step=100.0)
            
            total = sum(costs.values())
            st.metric("Total Cost", f"₹{total:,.2f}")
            
            st.subheader("Bills Documentation")
            bill_categories = [
                "Medical Bills", "Transport Receipts", "Compensation Documents",
                "Machine Repair Bills", "Damage Assessment", 
                "Implementation Cost Documents", "Miscellaneous"
            ]
            
            bill_uploads = {}
            cols = st.columns(2)
            for i, category in enumerate(bill_categories):
                with cols[i % 2]:
                    bill_uploads[category] = st.file_uploader(
                        f"Upload {category}",
                        type=["jpg", "jpeg", "png", "pdf"],
                        accept_multiple_files=True,
                        key=f"bill_{i}"
                    )
            
            if st.form_submit_button("Save Costing", use_container_width=True):
                data = {}
                for field in cost_fields:
                    data[field] = costs[field]
                data["Total Cost"] = total
                
                for category, files in bill_uploads.items():
                    if files:
                        paths = save_multiple_files(files, DIRECTORIES["cost_bills"], 
                                                   category.lower().replace(" ", "_"))
                        data[f"{category}_docs"] = ",".join(paths)
                
                df = load_sheet(INCIDENT_FILE, 'Costing')
                df = pd.concat([df, pd.DataFrame([data])], ignore_index=True)
                save_sheet(INCIDENT_FILE, 'Costing', df)
                
                st.success("✅ Costing saved successfully!")
                log_activity(st.session_state.username, "Added costing data")
    
    # Tab 8: Closure
    with tabs[7]:
        st.header("Closure")
        
        with st.form("closure_form"):
            col1, col2 = st.columns(2)
            with col1:
                incident_number = st.text_input("Incident Number for Closure")
                closure_date = st.date_input("Closure Date", value=date.today())
            with col2:
                closure_status = st.selectbox("Status", ["Closed", "Cancelled", "Pending"])
            
            closure_comments = st.text_area("Comments", height=100)
            
            if st.form_submit_button("Submit Closure", use_container_width=True):
                data = {
                    "Incident Number": incident_number,
                    "Closure Date": closure_date,
                    "Status": closure_status,
                    "Comments": closure_comments
                }
                
                df = load_sheet(INCIDENT_FILE, 'Closure')
                df = pd.concat([df, pd.DataFrame([data])], ignore_index=True)
                save_sheet(INCIDENT_FILE, 'Closure', df)
                
                st.success("✅ Closure submitted successfully!")
                log_activity(st.session_state.username, f"Closed incident: {incident_number}")
    
    # Tab 9: Report
    with tabs[8]:
        st.header("Incident Report")
        
        # Individual report
        st.subheader("Search Incident")
        incident_id = st.text_input("Enter Incident Number")
        
        if incident_id:
            sheets = ['Create', 'Injury', 'Investigation', 'RootCause', 'CA_PA', 'Costing', 'Closure']
            for sheet in sheets:
                df = load_sheet(INCIDENT_FILE, sheet)
                if df.empty:
                    continue
                
                # Filter based on appropriate column
                if sheet == 'Create':
                    filtered = df[df['Incident Number'] == incident_id]
                elif sheet == 'Injury':
                    filtered = df[df['Incident Number'] == incident_id]
                elif sheet == 'Investigation':
                    filtered = df[(df['Incident ID'] == incident_id) | 
                                 (df['Investigation ID'].astype(str).str.contains(incident_id, na=False))]
                elif sheet == 'RootCause':
                    filtered = df[df['Incident ID'] == incident_id]
                elif sheet == 'CA_PA':
                    filtered = df[df['Incident Number'] == incident_id]
                else:
                    filtered = df
                
                if not filtered.empty:
                    st.subheader(sheet)
                    st.dataframe(filtered, use_container_width=True)
        
        # Visualizations
        st.markdown("---")
        st.header("📊 Analytics Dashboard")
        
        # Load data for charts
        create_df = load_sheet(INCIDENT_FILE, 'Create')
        injury_df = load_sheet(INCIDENT_FILE, 'Injury')
        cost_df = load_sheet(INCIDENT_FILE, 'Costing')
        
        col1, col2 = st.columns(2)
        
        with col1:
            fig = create_incident_pie_chart(create_df)
            if fig:
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.info("No incident type data available")
            
            fig = create_trend_chart(create_df)
            if fig:
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.info("No trend data available")
        
        with col2:
            fig = create_severity_bar_chart(injury_df)
            if fig:
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.info("No severity data available")
            
            fig = create_cost_chart(cost_df)
            if fig:
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.info("No cost data available")
        
        # Editing section
        if is_editor(st.session_state.username):
            st.markdown("---")
            st.subheader("✏️ Edit Data (Admin Only)")
            
            edit_sheet = st.selectbox("Select sheet to edit", 
                                     ["Create", "Injury", "Investigation", "RootCause", "CA_PA", "Costing", "Closure"])
            
            if edit_sheet:
                df = load_sheet(INCIDENT_FILE, edit_sheet)
                if not df.empty:
                    edited_df = st.data_editor(df, use_container_width=True, key=f"edit_{edit_sheet}")
                    
                    if st.button(f"💾 Save {edit_sheet} Changes", use_container_width=True):
                        save_sheet(INCIDENT_FILE, edit_sheet, edited_df)
                        st.success(f"✅ {edit_sheet} updated successfully!")
                        log_activity(st.session_state.username, f"Edited {edit_sheet} data")
                else:
                    st.info("No data in this sheet")

# ============================================
# ASPECT/IMPACT PAGE
# ============================================

def aspect_impact_page():
    """Aspect/Impact Assessment module"""
    st.title("📊 Aspect/Impact Management")
    
    # Assessment selection
    col1, col2 = st.columns([2, 1])
    with col1:
        df = load_sheet(ASPECT_FILE, 'Activity')
        if not df.empty and 'Aspect/Impact Assessment Number' in df.columns:
            assessments = ["New Assessment"] + df['Aspect/Impact Assessment Number'].dropna().unique().tolist()
            selected = st.selectbox("Select Assessment", assessments)
            
            if selected != "New Assessment":
                st.session_state.current_assessment_id = selected
                st.session_state.edit_mode = True
            else:
                st.session_state.current_assessment_id = None
                st.session_state.edit_mode = False
    
    tabs = st.tabs(["Activity", "Aspect/Impact", "Score Card", "CA/PA", "Report"])
    
    # Tab 1: Activity
    with tabs[0]:
        st.header("Activity Information")
        
        # Load existing data
        activity_data = {}
        if st.session_state.edit_mode and st.session_state.current_assessment_id:
            df = load_sheet(ASPECT_FILE, 'Activity')
            filtered = df[df['Aspect/Impact Assessment Number'] == st.session_state.current_assessment_id]
            if not filtered.empty:
                activity_data = filtered.iloc[0].to_dict()
        
        with st.form("activity_form"):
            col1, col2 = st.columns(2)
            
            with col1:
                company_name = st.text_input("Company Name", value=activity_data.get('Company Name', ''))
                incident_number = st.text_input("Incident Number", value=activity_data.get('Incident Number', ''))
                incident_type = st.selectbox("Incident Type", 
                    ["Injury", "Near Miss", "Property Damage", "Environmental", "Fire", "Security", "Other"],
                    index=["Injury", "Near Miss", "Property Damage", "Environmental", "Fire", "Security", "Other"].index(
                        activity_data.get('Incident Type', 'Injury')) if activity_data.get('Incident Type', 'Injury') in ["Injury", "Near Miss", "Property Damage", "Environmental", "Fire", "Security", "Other"] else 0)
                incident_date = st.date_input("Incident Date", 
                    value=pd.to_datetime(activity_data.get('Incident Date', date.today())).date() if activity_data.get('Incident Date') else date.today())
            
            with col2:
                incident_location = st.text_input("Incident Location", value=activity_data.get('Incident Location', ''))
                reported_by_name = st.text_input("Reported By (Name)", value=activity_data.get('Reported By Name', ''))
                reported_by_id = st.text_input("Reported By (ID)", value=activity_data.get('Reported By ID', ''))
                reported_by_dept = st.text_input("Reported By (Department)", value=activity_data.get('Reported By Department', ''))
            
            incident_description = st.text_area("Incident Description", 
                                               value=activity_data.get('Incident Description', ''),
                                               height=100)
            
            # Assessment number
            if not st.session_state.edit_mode:
                # Generate new assessment number
                if not df.empty and 'Aspect/Impact Assessment Number' in df.columns:
                    existing = df['Aspect/Impact Assessment Number'].astype(str).dropna()
                    highest = 0
                    for num in existing:
                        if num.startswith('AI-'):
                            try:
                                highest = max(highest, int(num.split('-')[1]))
                            except:
                                pass
                    assessment_number = f"AI-{highest + 1:03d}"
                else:
                    assessment_number = "AI-001"
            else:
                assessment_number = st.session_state.current_assessment_id
            
            st.info(f"**Assessment Number:** {assessment_number}")
            
            # Evidence photos
            evidence_photos = st.file_uploader("Evidence Photos", 
                                              type=["jpg", "jpeg", "png"],
                                              accept_multiple_files=True)
            
            if st.form_submit_button("💾 Save Activity Information", use_container_width=True):
                photo_paths = save_multiple_files(evidence_photos, DIRECTORIES["aspect_images"], assessment_number)
                
                data = {
                    "Company Name": company_name,
                    "Incident Number": incident_number,
                    "Incident Type": incident_type,
                    "Incident Date": incident_date,
                    "Incident Description": incident_description,
                    "Incident Location": incident_location,
                    "Reported By Name": reported_by_name,
                    "Reported By ID": reported_by_id,
                    "Reported By Department": reported_by_dept,
                    "Aspect/Impact Assessment Number": assessment_number,
                    "Evidence Photos": len(photo_paths),
                    "Evidence Photo Paths": ",".join(photo_paths)
                }
                
                df = load_sheet(ASPECT_FILE, 'Activity')
                if st.session_state.edit_mode:
                    df = df[df['Aspect/Impact Assessment Number'] != assessment_number]
                df = pd.concat([df, pd.DataFrame([data])], ignore_index=True)
                save_sheet(ASPECT_FILE, 'Activity', df)
                
                st.session_state.current_assessment_id = assessment_number
                st.session_state.edit_mode = True
                st.success(f"✅ Activity saved with Assessment: {assessment_number}")
                log_activity(st.session_state.username, f"Created/updated aspect assessment: {assessment_number}")
                st.rerun()
    
    # Tab 2: Aspect/Impact
    with tabs[1]:
        st.header("Aspect/Impact Analysis")
        
        if not st.session_state.current_assessment_id:
            st.warning("⚠️ Please save Activity information first.")
        else:
            aspect_data = {}
            df = load_sheet(ASPECT_FILE, 'AspectImpact')
            filtered = df[df['Assessment Number'] == st.session_state.current_assessment_id]
            if not filtered.empty:
                aspect_data = filtered.iloc[0].to_dict()
            
            with st.form("aspect_form"):
                st.subheader("Aspects (Contributing Factors)")
                
                col1, col2 = st.columns(2)
                
                with col1:
                    human_factors = st.text_area("Human Factors", 
                        value=aspect_data.get('Human Factors', ''),
                        placeholder="E.g., lapse in attention, fatigue", height=80)
                    equipment_factors = st.text_area("Equipment/Material Factors",
                        value=aspect_data.get('Equipment Factors', ''),
                        placeholder="E.g., mechanical breakdown, faulty components", height=80)
                    procedural_factors = st.text_area("Procedural Factors",
                        value=aspect_data.get('Procedural Factors', ''),
                        placeholder="E.g., deviation from SOPs", height=80)
                
                with col2:
                    environmental_factors = st.text_area("Environmental Factors",
                        value=aspect_data.get('Environmental Factors', ''),
                        placeholder="E.g., weather, physical environment", height=80)
                    organizational_factors = st.text_area("Organizational Factors",
                        value=aspect_data.get('Organizational Factors', ''),
                        placeholder="E.g., inadequate supervision", height=80)
                    external_factors = st.text_area("External Factors",
                        value=aspect_data.get('External Factors', ''),
                        placeholder="E.g., contractor failures", height=80)
                
                st.subheader("Impacts")
                
                col1, col2 = st.columns(2)
                
                with col1:
                    safety_impact = st.text_area("Safety Impact",
                        value=aspect_data.get('Safety Impact', ''),
                        placeholder="E.g., Injuries, near misses", height=80)
                    environment_impact = st.text_area("Environmental Impact",
                        value=aspect_data.get('Environment Impact', ''),
                        placeholder="E.g., Pollution, hazardous release", height=80)
                    operations_impact = st.text_area("Operations Impact",
                        value=aspect_data.get('Operations Impact', ''),
                        placeholder="E.g., Process disruptions, downtime", height=80)
                
                with col2:
                    financial_impact = st.text_area("Financial Impact",
                        value=aspect_data.get('Financial Impact', ''),
                        placeholder="E.g., Direct and indirect costs", height=80)
                    reputation_impact = st.text_area("Reputation Impact",
                        value=aspect_data.get('Reputation Impact', ''),
                        placeholder="E.g., Public image, stakeholder trust", height=80)
                    legal_impact = st.text_area("Legal/Compliance Impact",
                        value=aspect_data.get('Legal Impact', ''),
                        placeholder="E.g., Regulatory breaches, penalties", height=80)
                
                if st.form_submit_button("💾 Save Aspect/Impact Analysis", use_container_width=True):
                    data = {
                        "Assessment Number": st.session_state.current_assessment_id,
                        "Human Factors": human_factors,
                        "Equipment Factors": equipment_factors,
                        "Procedural Factors": procedural_factors,
                        "Environmental Factors": environmental_factors,
                        "Organizational Factors": organizational_factors,
                        "External Factors": external_factors,
                        "Safety Impact": safety_impact,
                        "Environment Impact": environment_impact,
                        "Operations Impact": operations_impact,
                        "Financial Impact": financial_impact,
                        "Reputation Impact": reputation_impact,
                        "Legal Impact": legal_impact
                    }
                    
                    df = load_sheet(ASPECT_FILE, 'AspectImpact')
                    if not df.empty and 'Assessment Number' in df.columns:
                        df = df[df['Assessment Number'] != st.session_state.current_assessment_id]
                    df = pd.concat([df, pd.DataFrame([data])], ignore_index=True)
                    save_sheet(ASPECT_FILE, 'AspectImpact', df)
                    
                    st.success("✅ Aspect/Impact analysis saved!")
                    log_activity(st.session_state.username, f"Added aspect/impact for assessment: {st.session_state.current_assessment_id}")
    
    # Tab 3: Score Card
    with tabs[2]:
        st.header("Risk Score Card")
        
        if not st.session_state.current_assessment_id:
            st.warning("⚠️ Please save Activity information first.")
        else:
            score_data = {}
            df = load_sheet(ASPECT_FILE, 'ScoreCard')
            filtered = df[df['Assessment Number'] == st.session_state.current_assessment_id]
            if not filtered.empty:
                score_data = filtered.iloc[0].to_dict()
            
            with st.form("score_form"):
                col1, col2, col3 = st.columns(3)
                
                with col1:
                    severity = st.number_input("Severity (S)", 
                        min_value=1, max_value=10, 
                        value=int(score_data.get('Severity', 1)))
                    st.caption("1-2: Minimal | 3-4: Minor | 5-6: Moderate | 7-8: Major | 9-10: Catastrophic")
                
                with col2:
                    occurrence = st.number_input("Occurrence (O)",
                        min_value=1, max_value=10,
                        value=int(score_data.get('Occurrence', 1)))
                    st.caption("1-2: Very unlikely | 3-4: Low | 5-6: Moderate | 7-8: High | 9-10: Almost certain")
                
                with col3:
                    detection = st.number_input("Detection (D)",
                        min_value=1, max_value=10,
                        value=int(score_data.get('Detection', 1)))
                    st.caption("1-2: Very high | 3-4: High | 5-6: Moderate | 7-8: Low | 9-10: Very low")
                
                rpn = severity * occurrence * detection
                
                # Risk level
                if rpn <= 50:
                    risk_level = "Low Risk"
                    risk_color = "green"
                elif rpn <= 100:
                    risk_level = "Moderate Risk"
                    risk_color = "orange"
                elif rpn <= 200:
                    risk_level = "High Risk"
                    risk_color = "red"
                else:
                    risk_level = "Critical Risk"
                    risk_color = "darkred"
                
                st.metric("Risk Priority Number (RPN)", rpn)
                st.markdown(f"<h3 style='color: {risk_color};'>Risk Level: {risk_level}</h3>", 
                           unsafe_allow_html=True)
                
                risk_notes = st.text_area("Risk Notes", value=score_data.get('Risk Notes', ''))
                
                if st.form_submit_button("💾 Save Score Card", use_container_width=True):
                    data = {
                        "Assessment Number": st.session_state.current_assessment_id,
                        "Severity": severity,
                        "Occurrence": occurrence,
                        "Detection": detection,
                        "RPN": rpn,
                        "Risk Level": risk_level,
                        "Risk Notes": risk_notes
                    }
                    
                    df = load_sheet(ASPECT_FILE, 'ScoreCard')
                    if not df.empty and 'Assessment Number' in df.columns:
                        df = df[df['Assessment Number'] != st.session_state.current_assessment_id]
                    df = pd.concat([df, pd.DataFrame([data])], ignore_index=True)
                    save_sheet(ASPECT_FILE, 'ScoreCard', df)
                    
                    st.success("✅ Score Card saved!")
                    log_activity(st.session_state.username, f"Added score card for assessment: {st.session_state.current_assessment_id}")
    
    # Tab 4: CA/PA
    with tabs[3]:
        st.header("Corrective/Preventive Actions")
        
        if not st.session_state.current_assessment_id:
            st.warning("⚠️ Please save Activity information first.")
        else:
            # Existing actions
            df = load_sheet(ASPECT_FILE, 'CAPA')
            filtered = df[df['Assessment Number'] == st.session_state.current_assessment_id]
            
            if not filtered.empty:
                st.subheader("Existing Actions")
                st.dataframe(filtered.drop(columns=['Assessment Number']), use_container_width=True)
            
            # Add new action
            st.subheader("Add New Action")
            with st.form("capa_form_ai"):
                col1, col2 = st.columns(2)
                
                with col1:
                    action_desc = st.text_area("Action Description", height=80)
                    assigned_to = st.text_input("Assigned To")
                
                with col2:
                    due_date = st.date_input("Due Date", value=date.today())
                    verification = st.selectbox("Verification Approach",
                        ["Audit", "Inspection", "Monitoring", "Testing", "Review", "Other"])
                
                lessons_learned = st.text_area("Lessons Learned", height=60)
                
                if st.form_submit_button("Add Action", use_container_width=True):
                    data = {
                        "Assessment Number": st.session_state.current_assessment_id,
                        "Action Description": action_desc,
                        "Assigned To": assigned_to,
                        "Due Date": due_date,
                        "Verification Approach": verification,
                        "Lessons Learned": lessons_learned,
                        "Status": "Open",
                        "Date Added": date.today()
                    }
                    
                    df = load_sheet(ASPECT_FILE, 'CAPA')
                    df = pd.concat([df, pd.DataFrame([data])], ignore_index=True)
                    save_sheet(ASPECT_FILE, 'CAPA', df)
                    
                    st.success("✅ Action added successfully!")
                    log_activity(st.session_state.username, f"Added CA/PA for assessment: {st.session_state.current_assessment_id}")
                    st.rerun()
    
    # Tab 5: Report
    with tabs[4]:
        st.header("Assessment Report")
        
        if not st.session_state.current_assessment_id:
            st.warning("⚠️ No assessment selected.")
        else:
            # Load all data
            activity_df = load_sheet(ASPECT_FILE, 'Activity')
            aspect_df = load_sheet(ASPECT_FILE, 'AspectImpact')
            score_df = load_sheet(ASPECT_FILE, 'ScoreCard')
            capa_df = load_sheet(ASPECT_FILE, 'CAPA')
            
            # Filter for current assessment
            current_activity = activity_df[activity_df['Aspect/Impact Assessment Number'] == st.session_state.current_assessment_id] if not activity_df.empty else pd.DataFrame()
            current_aspect = aspect_df[aspect_df['Assessment Number'] == st.session_state.current_assessment_id] if not aspect_df.empty else pd.DataFrame()
            current_score = score_df[score_df['Assessment Number'] == st.session_state.current_assessment_id] if not score_df.empty else pd.DataFrame()
            current_capa = capa_df[capa_df['Assessment Number'] == st.session_state.current_assessment_id] if not capa_df.empty else pd.DataFrame()
            
            # Display report
            if not current_activity.empty:
                st.subheader("Activity Summary")
                col1, col2 = st.columns(2)
                with col1:
                    st.write(f"**Assessment Number:** {st.session_state.current_assessment_id}")
                    st.write(f"**Company:** {current_activity.iloc[0].get('Company Name', 'N/A')}")
                    st.write(f"**Incident Type:** {current_activity.iloc[0].get('Incident Type', 'N/A')}")
                with col2:
                    st.write(f"**Date:** {current_activity.iloc[0].get('Incident Date', 'N/A')}")
                    st.write(f"**Location:** {current_activity.iloc[0].get('Incident Location', 'N/A')}")
                    st.write(f"**Reported By:** {current_activity.iloc[0].get('Reported By Name', 'N/A')}")
                
                st.write(f"**Description:** {current_activity.iloc[0].get('Incident Description', 'N/A')}")
            
            if not current_score.empty:
                st.subheader("Risk Assessment")
                col1, col2 = st.columns(2)
                with col1:
                    st.write(f"**Severity:** {int(current_score.iloc[0].get('Severity', 0))}")
                    st.write(f"**Occurrence:** {int(current_score.iloc[0].get('Occurrence', 0))}")
                    st.write(f"**Detection:** {int(current_score.iloc[0].get('Detection', 0))}")
                with col2:
                    rpn = current_score.iloc[0].get('RPN', 0)
                    risk_level = current_score.iloc[0].get('Risk Level', 'N/A')
                    st.metric("RPN", int(rpn))
                    st.write(f"**Risk Level:** {risk_level}")
            
            if not current_capa.empty:
                st.subheader("Actions")
                st.dataframe(current_capa.drop(columns=['Assessment Number']), use_container_width=True)

# ============================================
# DASHBOARD PAGE
# ============================================

def dashboard_page():
    """Main dashboard"""
    st.title("🛡️ EHS Management System")
    st.caption(f"Welcome back, {st.session_state.username}! Here's your EHS overview.")
    
    # Load data
    create_df = load_sheet(INCIDENT_FILE, 'Create')
    injury_df = load_sheet(INCIDENT_FILE, 'Injury')
    cost_df = load_sheet(INCIDENT_FILE, 'Costing')
    
    # Quick stats
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.metric("Total Incidents", len(create_df) if not create_df.empty else 0)
    
    with col2:
        if not create_df.empty and 'Incident_type' in create_df.columns:
            injury_count = len(create_df[create_df['Incident_type'] == 'Injury'])
        else:
            injury_count = 0
        st.metric("Injury Incidents", injury_count)
    
    with col3:
        if not injury_df.empty and 'Injury_status' in injury_df.columns:
            fatal_count = len(injury_df[injury_df['Injury_status'] == 'Fatal'])
        else:
            fatal_count = 0
        st.metric("Fatal Incidents", fatal_count, delta_color="inverse")
    
    with col4:
        if not cost_df.empty and 'Total Cost' in cost_df.columns:
            total_cost = cost_df['Total Cost'].sum()
        else:
            total_cost = 0
        st.metric("Total Cost", f"₹{total_cost:,.2f}")
    
    st.markdown("---")
    
    # Charts
    col1, col2 = st.columns(2)
    
    with col1:
        fig = create_incident_pie_chart(create_df)
        if fig:
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("No incident data available")
    
    with col2:
        fig = create_severity_bar_chart(injury_df)
        if fig:
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("No severity data available")
    
    col1, col2 = st.columns(2)
    
    with col1:
        fig = create_trend_chart(create_df)
        if fig:
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("No trend data available")
    
    with col2:
        fig = create_cost_chart(cost_df)
        if fig:
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("No cost data available")
    
    # Recent activity
    st.markdown("---")
    st.subheader("Recent Activity")
    
    try:
        login_df = pd.read_excel(LOGIN_FILE)
        if not login_df.empty:
            recent = login_df.tail(10)
            st.dataframe(recent, use_container_width=True)
    except:
        st.info("No activity logs available")

# ============================================
# PLACEHOLDER PAGES
# ============================================

def placeholder_page(title, description):
    """Generic placeholder page for future modules"""
    st.title(title)
    st.info(f"🔄 {description}")
    st.markdown("This module is under development. Check back soon for updates!")

# ============================================
# RUN APP
# ============================================

if __name__ == "__main__":
    main()
