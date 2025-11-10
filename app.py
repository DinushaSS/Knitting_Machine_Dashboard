# Load required libraries
import streamlit as st
import pandas as pd
import os
import plotly.express as px

# Set page configuration
st.set_page_config(
    page_title="Knitting Machine Dashboard",
    page_icon="🔎",
    layout="wide"
)

# Add CSS for button styling
st.markdown("""
<style>

    /* Reduce white space at the top */
    .main .block-container {
        padding-top: 1rem;
        padding-bottom: 2rem;
    }

    /* Target only the main content area, not sidebar */
    [data-testid="stMain"] {
        margin-top: -70px;
    }

    /* Ensure content has enough height */
    .main .block-container {
        min-height: 100vh;
    }

    /* Keep the sidebar intact */
    [data-testid="stSidebar"] {
        margin-top: 0px;
    }

    .stButton button {
        width: 150px;
        height: 50px;
        font-size: 16px;
        font-weight: bold;
    }

</style>
""", unsafe_allow_html=True)

# Dashboard Title
#st.title("🔎 Knitting Machine Dashboard")
st.markdown("""
    <h1 style = "text-align:center; font-family:'Lato', sans-serif; color: #e8e8e8;">
    KNITTING MACHINE DASHBOARD
    </h1>
    """,unsafe_allow_html=True)

# Sidebar
st.sidebar.title("Menu")

# Initialize session state for selected window
if 'selected_window' not in st.session_state:
    st.session_state.selected_window = "Overview"

# Menu options
menu_options = ["Overview", "Running", "Parking", "Advantis","Data Table"]

# Create menu buttons
for option in menu_options:
    is_active = st.session_state.selected_window == option
    
    button_type = "primary" if is_active else "secondary"

    if st.sidebar.button(option,key=option,type=button_type,use_container_width=True):
        st.session_state.selected_window = option
        st.rerun()

# Excel file path
file_path = "Knitting Machine Dashboard.xlsx"

try:
    # Display content based on selected window
    if st.session_state.selected_window == "Overview":
        # Load data from Excel file
        df = pd.read_excel(file_path, sheet_name='Machines')
        
        # Count each machine type
        machine_counts = df["Type"].value_counts()

        # Display total machine card
        st.markdown("<h3 style='text-align:center;font-size:20px;'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Total Machines Count:</h3>",unsafe_allow_html=True)

        # Total machine card with custom color
        col1, col2, col3, col4, col5 = st.columns(5)
        
        with col3:
            st.markdown("""
            <div style="background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); 
                        color: white; border-radius: 10px; padding: 10px; text-align: center;">
                <div style="font-size: 14px;">Total Machines</div>
                <div style="font-size: 24px; font-weight: bold;">{}</div>
            </div>
            """.format(len(df)), unsafe_allow_html=True)

        st.markdown('<div style="margin-top:20px;"></div>',unsafe_allow_html=True)
        
        # Create cards for each machine type
        st.markdown("<h3 style='text-align:left;font-size:20px;'>Count by Machine Types:</h3>",unsafe_allow_html=True)

        # Color palette for machine types
        colors = [
            "linear-gradient(135deg, #667eea 0%, #764ba2 100%)",
            "linear-gradient(135deg, #f093fb 0%, #f5576c 100%)",
            "linear-gradient(135deg, #4facfe 0%, #00f2fe 100%)",
            "linear-gradient(135deg, #43e97b 0%, #38f9d7 100%)",
            "linear-gradient(135deg, #fa709a 0%, #fee140 100%)",
            "linear-gradient(135deg, #a8edea 0%, #fed6e3 100%)"
        ]

        # Calculate how many columns per row (5 columns layout)
        cols_per_row = 5
        machine_types = list(machine_counts.items())

        # Display machine type cards in rows
        for i in range(0, len(machine_types), cols_per_row):
            cols = st.columns(cols_per_row)
            for j, (machine_type, count) in enumerate(machine_types[i:i+cols_per_row]):
                with cols[j]:
                    color = colors[j % len(colors)]
                    st.markdown("""
                    <div style="background: {}; 
                                color: white; border-radius: 10px; padding: 10px; text-align: center;">
                        <div style="font-size: 14px;">{} Machines</div>
                        <div style="font-size: 24px; font-weight: bold;">{}</div>
                    </div>
                    """.format(color, machine_type, count), unsafe_allow_html=True)
        
        #Add bar chart visualization

        #Add a space before title
        st.markdown('<div style="margin-top:20px;"></div>',unsafe_allow_html=True)

        st.markdown("<h3 style='text-align:center;font-size:20px;'>📊 Knitting Machines Distribution by Diameter:</h3>",unsafe_allow_html=True)

        #Create filter for machine types
        selected_types = st.multiselect("Select Machine Types to Display:",
        options=df["Type"].unique(),default=df["Type"].unique()[:3])

        #Filter data based on selection
        if selected_types:
            filtered_df = df[df["Type"].isin(selected_types)]

            #Create proper count by diameter and type
            count_data = filtered_df.groupby(["Diameter","Type"]).size().reset_index(name="Count")

            #Create bar chart
            fig = px.bar(count_data,x="Diameter",y="Count",color="Type",barmode="group",
            title="Machine Count by Diameter and Type",
            labels={"Diameter":"Diameter","Count":"Machine Count"},
            category_orders={"Diameter":sorted(filtered_df["Diameter"].unique())}
            )

            #Center the chart title
            fig.update_layout(title_x=0.37)

            #Customized the chart
            fig.update_layout(xaxis_title="Diameter",yaxis_title="Number of Machines",
            showlegend = True,bargap=0.2,bargroupgap=0.1,plot_bgcolor='#262730',paper_bgcolor='#262730')

            #Add rounded corners to chart aontainer
            st.markdown("""
            <style>
                [data-testid="stPlotlyChart"] > div {
                    border-radius: 10px;
                    overflow: hidden;
                    }
            </style>
            """,unsafe_allow_html=True)

            #Add data labels on bars
            fig.update_traces(texttemplate="%{y}",
            textposition="outside")

            #Display the chart
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("Please select at least one machine type to display the chart")           
    
    elif st.session_state.selected_window == "Running":
        
        #Load data from Excel file
        df = pd.read_excel(file_path,sheet_name='Machines')

        #Remove any rows where status is "Status" (header rows)
        df = df[df["Status"]!="Status"]

        #Filter only Active machines
        active_df = df[df["Status"] == "Active"]

        #Centered heading
        st.markdown("<h3 style='text-align:center;font-size:20px;'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Total Running Machines Count:</h3>",unsafe_allow_html=True)

        #Center the Total Active Machines Card
        col1, col2, col3, col4, col5 = st.columns(5)

        with col3:
            st.markdown("""
            <div style="background:linear-gradient(135deg, #43e97b 0%, #38f9d7 100%); 
                    color: white; border-radius: 10px; padding: 10px; text-align: center;">
            <div style="font-size: 14px;">Total Running Machines</div>
            <div style="font-size: 24px; font-weight: bold;">{}</div>
            </div>
            """.format(len(active_df)), unsafe_allow_html=True)
        
        #Count active machine by Type
        active_machine_counts = active_df["Type"].value_counts()

        st.markdown('<div style="margin-top:20px;"></div>',unsafe_allow_html=True)

        #Heading for machine types
        st.markdown("<h3 style='text-align:left;font-size:20px;'>Running Count by Machine Types:</h3>",unsafe_allow_html=True)

        # Color palette
        colors = [
            "linear-gradient(135deg, #667eea 0%, #764ba2 100%)",
            "linear-gradient(135deg, #f093fb 0%, #f5576c 100%)",
            "linear-gradient(135deg, #4facfe 0%, #00f2fe 100%)",
            "linear-gradient(135deg, #43e97b 0%, #38f9d7 100%)",
            "linear-gradient(135deg, #fa709a 0%, #fee140 100%)",
            "linear-gradient(135deg, #a8edea 0%, #fed6e3 100%)"
        ]

        #Display card
        if len(active_machine_counts) > 0:
            cols_per_row = 5
            machine_types = list(active_machine_counts.items())

            for i in range(0,len(machine_types),cols_per_row):
                cols = st.columns(cols_per_row)
                for j, (machine_type,count) in enumerate(machine_types[i:i+cols_per_row]):
                    with cols[j]:
                        color_index = (i+j) % len (colors)
                        color = colors[color_index]
                        st.markdown("""
                        <div style="background: {}; 
                            color: white; border-radius: 10px; padding: 10px; text-align: center;">
                            <div style="font-size: 14px;">{} Machines</div>
                            <div style="font-size: 24px; font-weight: bold;">{}</div>
                        </div>
                        """.format(color, machine_type, count), unsafe_allow_html=True)

        else:
            st.info("No active machines found")
        
        #Add some space before chart
        st.markdown('<div style = "Margin-top:20px;"></div>',unsafe_allow_html=True)

        #Chart title
        st.markdown("<h3 style='text-align:center;font-size:20px;'>📊 Running Knitting Machines Distribution by Diameter:</h3>",unsafe_allow_html=True)

        #Create filter for machine types
        selected_types = st.multiselect("Select Machine Types to Display:",
        options=active_df["Type"].unique(),default=active_df["Type"].unique()[:3] if len(active_df["Type"].unique()) >= 3 else active_df["Type"].unique(),key="running_machine_types")

        #Filter data based on selection
        if selected_types:
            filtered_active_df = active_df[active_df["Type"].isin(selected_types)]

            #Create count by diameter and type
            count_data = filtered_active_df.groupby(["Diameter","Type"]).size().reset_index(name="Count")

            #Create bar chart
            fig = px.bar(count_data,x="Diameter",y="Count",color="Type",barmode="group",
                title="Active Machine Count by Diameter and Type",
                labels={"Diameter":"Diameter","Count":"Machine Count"},
                category_orders={"Diameter":sorted(filtered_active_df["Diameter"].unique())})
            
            #Center the chart title
            fig.update_layout(title_x=0.33)

            #Customized the chart
            fig.update_layout(xaxis_title="Diameter",yaxis_title="Number of Machines",
                showlegend=True,bargap=0.2,bargroupgap=0.1,plot_bgcolor='#262730',paper_bgcolor='#262730')
            
            #Add data labels on bars
            fig.update_traces(texttemplate="%{y}",textposition="outside")

            #Add rounded corners to chart
            st.markdown("""
            <style>
                [data-testid="stPlotlyChart"] > div {
                    border-radius:10px;
                    overflow:hidden;
                    }
            </style>
            """,unsafe_allow_html=True)

            st.plotly_chart(fig,use_container_width=True)
        else:
            st.info("Please select at least one machine type to display the chart")

    elif st.session_state.selected_window == "Parking":
        
        # Load data from Excel file
        df = pd.read_excel(file_path,sheet_name='Machines')

        # Remove any rows where Status is "Status" (header rows)
        df = df[df["Status"] != "Status"]

        # Filter only Idle Machines
        idle_df = df[df["Status"] == "Idle"]

        # Centered heading 
        st.markdown("<h3 style = 'text-align:center; font-size: 20px;'>&nbsp&nbsp&nbsp&nbsp&nbspTotal Parking Machines Count:</h3>",unsafe_allow_html=True)

        # Center the Total Idle Machine card
        col1, col2, col3, col4, col5 = st.columns(5)

        with col3:
            st.markdown("""
            <div style="background: linear-gradient(135deg, #fa709a 0%, #fee140 100%); 
                    color: white; border-radius: 10px; padding: 10px; text-align: center;">
            <div style="font-size: 14px;">Total Parking Machines</div>
            <div style="font-size: 24px; font-weight: bold;">{}</div>
            </div>
            """.format(len(idle_df)), unsafe_allow_html=True)

        # Count idle machines by Type
        idle_machine_counts = idle_df["Type"].value_counts()

        st.markdown('<div style = "Margin-top:20px;"></div>',unsafe_allow_html=True)

        # Heading for machine types
        st.markdown("<h3 style='text-align: left;font-size:20px;'>Parking Count by Machine Types:</h3>",unsafe_allow_html=True)

        #Color pallette
        colors = [
        "linear-gradient(135deg, #667eea 0%, #764ba2 100%)",
        "linear-gradient(135deg, #f093fb 0%, #f5576c 100%)",
        "linear-gradient(135deg, #4facfe 0%, #00f2fe 100%)",
        "linear-gradient(135deg, #43e97b 0%, #38f9d7 100%)",
        "linear-gradient(135deg, #fa709a 0%, #fee140 100%)",
        "linear-gradient(135deg, #a8edea 0%, #fed6e3 100%)"
        ]

        # Display cards
        if len(idle_machine_counts) > 0:
            cols_per_row = 5
            machine_types = list(idle_machine_counts.items())
        
            for i in range(0, len(machine_types), cols_per_row):
                cols = st.columns(cols_per_row)
                for j, (machine_type, count) in enumerate(machine_types[i:i+cols_per_row]):
                    with cols[j]:
                        color_index = (i + j) % len(colors)
                        color = colors[color_index]
                        st.markdown("""
                        <div style="background: {}; 
                        color: white; border-radius: 10px; padding: 10px; text-align: center;">
                        <div style="font-size: 14px;">{} Machines</div>
                        <div style="font-size: 24px; font-weight: bold;">{}</div>
                        </div>
                        """.format(color, machine_type, count), unsafe_allow_html=True)
        else:
            st.info("No parking machines found")
        
        # Add space before chart
        st.markdown('<div style="margin-top: 20px;"></div>', unsafe_allow_html=True)
    
        # Chart title
        st.markdown("<h3 style='text-align: center; font-size: 20px;'>📊 Parking Knitting Machine Distribution by Diameter:</h3>", unsafe_allow_html=True)

        # Create two filters side by side
        filter_col1, filter_col2 = st.columns(2)

        with filter_col1:
            type_options = ["All"] + list(idle_df["Type"].unique())
            selected_type = st.selectbox("Select Machine Type:",
                                      options=type_options,
                                      key="parking_machine_type")

        with filter_col2:
            location_options = ["All", "Pathway Parking", "Batch Parking", "Training (M/C)"]
            selected_location = st.selectbox("Select Location Group:",
                                         options=location_options,
                                         key="parking_location_group")

        # Filter data based on both selections
        filtered_idle_df = idle_df.copy()

        if selected_type != "All":
            filtered_idle_df = filtered_idle_df[filtered_idle_df["Type"] == selected_type]

        if selected_location != "All":
            filtered_idle_df = filtered_idle_df[filtered_idle_df["Location Group"] == selected_location]

        # Check if there's data to display
        if len(filtered_idle_df) > 0:
            # Create count by diameter and type
            count_data = filtered_idle_df.groupby(["Diameter","Type"]).size().reset_index(name="Count")
            # Create bar chart
            fig = px.bar(count_data, x="Diameter", y="Count", color="Type", barmode="group",
                     title="Parking Machine Count by Diameter and Type",
                     labels={"Diameter":"Diameter","Count":"Machine Count"},
                     category_orders={"Diameter":sorted(filtered_idle_df["Diameter"].unique())})

            # Center the chart title
            fig.update_layout(title_x=0.35)

            # Customize the chart
            fig.update_layout(xaxis_title="Diameter",
                            yaxis_title="Number of Machines",
                            showlegend=True,
                            bargap=0.2,
                            bargroupgap=0.1,
                            plot_bgcolor='#262730',
                            paper_bgcolor='#262730')
            
            # Add data labels on bars
            fig.update_traces(texttemplate="%{y}", textposition="outside")

            # Add rounded corners to chart
            st.markdown("""
            <style>
                [data-testid="stPlotlyChart"] > div {
                    border-radius: 15px;
                    overflow: hidden;
                    }
            </style>
            """, unsafe_allow_html=True)
        
            # Display the chart
            st.plotly_chart(fig, use_container_width=True)
        
        else:
            st.info("No data available for the selected filters")

    elif st.session_state.selected_window == "Advantis":
        
        # Load data from Advantis machines sheet
        advantis_df = pd.read_excel(file_path, sheet_name='Advantis Machines')
    
        # Centered heading
        st.markdown("<h3 style='text-align: center; font-size: 20px;'>&nbsp&nbsp&nbsp&nbspAdvantis Machines Count:</h3>", unsafe_allow_html=True)
    
        # Center the Total Advantis Machines card
        col1, col2, col3, clo4, col5 = st.columns(5)
    
        with col3:
            st.markdown("""
            <div style="background: linear-gradient(135deg, #4facfe 0%, #00f2fe 100%); 
                    color: white; border-radius: 10px; padding: 10px; text-align: center;">
            <div style="font-size: 14px;">Total Advantis Machines</div>
            <div style="font-size: 24px; font-weight: bold;">{}</div>
            </div>
            """.format(len(advantis_df)), unsafe_allow_html=True)
    
        # Count advantis machines by Type
        advantis_machine_counts = advantis_df["Type"].value_counts()

        st.markdown('<div style = "Margin-top:20px;"></div>',unsafe_allow_html=True)
    
        # Heading for machine types
        st.markdown("<h3 style='text-align: left; font-size: 20px;'>Advantis Count by Machine Types:</h3>", unsafe_allow_html=True)
    
        # Color palette
        colors = [
            "linear-gradient(135deg, #667eea 0%, #764ba2 100%)",
            "linear-gradient(135deg, #f093fb 0%, #f5576c 100%)",
            "linear-gradient(135deg, #4facfe 0%, #00f2fe 100%)",
            "linear-gradient(135deg, #43e97b 0%, #38f9d7 100%)",
            "linear-gradient(135deg, #fa709a 0%, #fee140 100%)",
            "linear-gradient(135deg, #a8edea 0%, #fed6e3 100%)"
        ]
    
        # Display cards
        if len(advantis_machine_counts) > 0:
            cols_per_row = 5
            machine_types = list(advantis_machine_counts.items())
        
            for i in range(0, len(machine_types), cols_per_row):
                cols = st.columns(cols_per_row)
                for j, (machine_type, count) in enumerate(machine_types[i:i+cols_per_row]):
                    with cols[j]:
                        color_index = (i + j) % len(colors)
                        color = colors[color_index]
                        st.markdown("""
                        <div style="background: {}; 
                                color: white; border-radius: 10px; padding: 10px; text-align: center;">
                        <div style="font-size: 14px;">{} Machines</div>
                        <div style="font-size: 24px; font-weight: bold;">{}</div>
                        </div>
                        """.format(color, machine_type, count), unsafe_allow_html=True)
        else:
            st.info("No advantis machines found")
    
        # Add space before chart
        st.markdown('<div style="margin-top: 20px;"></div>', unsafe_allow_html=True)
    
        # Chart title
        st.markdown("<h3 style='text-align: center; font-size: 20px;'>📊 Advantis Machine Distribution by Diameter</h3>", unsafe_allow_html=True)
    
        # Create two dropdown filters side by side
        filter_col1, filter_col2 = st.columns(2)
    
        with filter_col1:
            type_options = ["All"] + list(advantis_df["Type"].unique())
            selected_type = st.selectbox("Select Machine Type:",
                                      options=type_options,
                                      key="advantis_machine_type")
    
        with filter_col2:
            location_options = ["All"] + list(advantis_df["Current Location"].unique())
            selected_location = st.selectbox("Select Current Location:",
                                         options=location_options,
                                         key="advantis_current_location")
    
        # Filter data based on both selections
        filtered_advantis_df = advantis_df.copy()
    
        if selected_type != "All":
            filtered_advantis_df = filtered_advantis_df[filtered_advantis_df["Type"] == selected_type]
    
        if selected_location != "All":
            filtered_advantis_df = filtered_advantis_df[filtered_advantis_df["Current Location"] == selected_location]
    
        # Check if there's data to display
        if len(filtered_advantis_df) > 0:
            # Create count by diameter and type
            count_data = filtered_advantis_df.groupby(["Diameter","Type"]).size().reset_index(name="Count")
        
            # Create bar chart
            fig = px.bar(count_data, x="Diameter", y="Count", color="Type", barmode="group",
                         title="Advantis Machine Count by Diameter and Type",
                         labels={"Diameter":"Diameter","Count":"Machine Count"},
                         category_orders={"Diameter":sorted(filtered_advantis_df["Diameter"].unique())})
        
            # Center the chart title
            fig.update_layout(
                title_x=0.5,
                title_xanchor='center'
            )
        
            # Customize the chart
            fig.update_layout(xaxis_title="Diameter",
                             yaxis_title="Number of Machines",
                             showlegend=True,
                             bargap=0.2,
                             bargroupgap=0.1,
                            plot_bgcolor='#262730',
                            paper_bgcolor='#262730')
        
            # Add data labels on bars
            fig.update_traces(texttemplate="%{y}", textposition="outside")
        
            # Add rounded corners to chart
            st.markdown("""
            <style>
                [data-testid="stPlotlyChart"] > div {
                    border-radius: 15px;
                    overflow: hidden;
                }
            </style>
            """, unsafe_allow_html=True)
        
            # Display the chart
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("No data available for the selected filters")
    
    elif st.session_state.selected_window == "Data Table":
        
        #Title
        st.markdown("<h3 style = 'text-align:center;font-size:20px;'>📅Data Tables</h3>",unsafe_allow_html=True)

        #Create tabs for different sheets
        tab1, tab2, tab3 = st.tabs(["MFI Machines","Advantis Machines","MFI - OUT"])

        with tab1:
            # Load machines sheet
            machines_df = pd.read_excel(file_path,sheet_name="Machines",usecols="A:H")

            # Convert service date to date only.(Remove Time)
            if 'Service Date' in machines_df.columns:
                machines_df['Service Date'] = pd.to_datetime(machines_df['Service Date']).dt.date

            st.markdown("<h3 style = 'text-align:center; font-size:20px;'>MFI Existing Machines Data</h3>",unsafe_allow_html=True)

            #Create filters in columns
            col1, col2, col3 = st.columns(3)

            with col1:
                status_options = ["All"] + list(machines_df["Status"].unique())
                status_filter = st.selectbox("Filter by Status:",options=status_options,
                key="machines_status_filter")
            
            with col2:
                type_options = ["All"] + list(machines_df["Type"].unique())
                type_filter = st.selectbox("Filter by Type:",options=type_options
                ,key="machines_type_filter")
            
            with col3:
                diameter_options = ["All"] + sorted(list(machines_df["Diameter"].unique()))
                diameter_filter = st.selectbox("Filter by Diameter:",options=diameter_options,key ="machines_diameter_filter")

            filtered_machines = machines_df.copy()

            if status_filter != "All":
                filtered_machines = filtered_machines[filtered_machines["Status"] == status_filter]

            if type_filter != "All":
                filtered_machines = filtered_machines[filtered_machines["Type"] == type_filter]

            if diameter_filter != "All":
                filtered_machines = filtered_machines[filtered_machines["Diameter"] == diameter_filter]

            # Show filtered data
            st.dataframe(filtered_machines,use_container_width=True)

            #Show count
            st.info(f"Showing {len(filtered_machines)} of {len(machines_df)} machines")
        
        with tab2:
            # Load advantis machines sheet
            advantis_df = pd.read_excel(file_path,sheet_name="Advantis Machines",usecols="A:G")
            st.markdown("<h3 style = 'text-align:center; font-size:20px;'>Advantis Existing Machines Data</h3>",unsafe_allow_html=True)
            
            #Crate filters with columns
            col1,col2,col3 = st.columns(3)

            with col1:
                type_options = ["All"] + list(advantis_df["Type"].unique())
                type_filter = st.selectbox("Filter by type:",options=type_options,
                key = "advantis_type_filter")

            with col2:
                diameter_options = ["All"] + sorted(list(advantis_df["Diameter"].unique()))
                diameter_filter = st.selectbox("Filter by diameter:",options=diameter_options,
                key = "advantis_diameter_filter")
            
            filtered_advantis = advantis_df.copy()

            if type_filter != "All":
                filtered_advantis = filtered_advantis[filtered_advantis["Type"] == type_filter]
            
            if diameter_filter != "All":
                filtered_advantis = filtered_advantis[filtered_advantis["Diameter"] == diameter_filter]

            # Show filtered data
            st.dataframe(filtered_advantis,use_container_width=True)

            #Show count
            st.info(f"Showing {len(filtered_advantis)} of {len(advantis_df)} machines")

        with tab3:
            # Load OUT machines sheet
            OUT_df = pd.read_excel(file_path,sheet_name="OUT",usecols="A:F")
            st.markdown("<h3 style = 'text-align:center; font-size:20px;'>Machines out from MFI Data</h3>",unsafe_allow_html=True)
            
            #Crate filters with columns
            col1,col2,col3 = st.columns(3)

            with col1:
                type_options = ["All"] + list(OUT_df["Type"].unique())
                type_filter = st.selectbox("Filter by type:",options=type_options,
                key = "OUT_type_filter")

            with col2:
                diameter_options = ["All"] + sorted(list(OUT_df["Diameter"].unique()))
                diameter_filter = st.selectbox("Filter by diameter:",options=diameter_options,
                key = "OUT_diameter_filter")
            
            filtered_OUT = OUT_df.copy()

            if type_filter != "All":
                filtered_OUT = filtered_OUT[filtered_OUT["Type"] == type_filter]
            
            if diameter_filter != "All":
                filtered_OUT = filtered_OUT[filtered_OUT["Diameter"] == diameter_filter]

            # Show filtered data
            st.dataframe(filtered_OUT,use_container_width=True)

            #Show count
            st.info(f"Showing {len(filtered_OUT)} of {len(OUT_df)} machines")

except FileNotFoundError:
    st.error(f"❌ File not found: {file_path}")
    st.info("Please make sure the Excel file is in the same folder as this app")

except Exception as e:
    st.error(f"❌ Error loading Excel file: {e}")
