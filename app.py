import streamlit as st
import pandas as pd
import plotly.express as px
from typing import List, Optional
from Automatic_Chart import automatic_chart_analysis, create_ppt

st.set_page_config(page_title="Auto Chart PPT Generator", layout="wide")
st.title("📊 Automatic Chart Generator + Data Export")

# Sidebar for file upload and visualization type selection
st.sidebar.header('Upload Excel File')
uploaded_file = st.sidebar.file_uploader('Choose a file', type=['xlsx', 'xls'])

# Automatic analysis option
auto_analysis = st.sidebar.checkbox('Enable Automatic Chart Analysis')

def read_excel_file(file) -> Optional[pd.DataFrame]:
    """Read an Excel file and return a DataFrame."""
    try:
        return pd.read_excel(file)
    except Exception as e:
        st.error(f"Error reading the file: {e}")
        return None

def create_plotly_chart(df: pd.DataFrame, visualization_type: str, x_axis: str, y_axis: List[str], color: Optional[str] = None):
    """Create a Plotly chart based on the selected type."""
    plot_functions = {
        'Bar Chart': px.bar,
        'Line Chart': px.line,
        'Scatter Plot': px.scatter,
        'Pie Chart': lambda df, **kwargs: px.pie(df, names=kwargs['x'], values=kwargs['y']),
        'Donut Chart': lambda df, **kwargs: px.pie(df, names=kwargs['x'], values=kwargs['y'], hole=0.5),
        'Bubble Chart': lambda df, **kwargs: px.scatter(df, x=kwargs['x'], y=kwargs['y'], size=kwargs['y']),
        'Area Chart': px.area,
        'Radar Chart': lambda df, **kwargs: px.line_polar(df, r=kwargs['y'], theta=kwargs['x'], line_close=True),
        'Mixed Chart': lambda df, **kwargs: px.line(df, x=kwargs['x'], y=kwargs['y']).add_scatter(x=df[kwargs['x']], y=df[kwargs['y']], mode='markers'),
        'Funnel Chart': px.funnel
    }

    if visualization_type in plot_functions:
        if visualization_type in ['Bar Chart', 'Line Chart', 'Scatter Plot', 'Bubble Chart', 'Area Chart']:
            fig = plot_functions[visualization_type](df, x=x_axis, y=y_axis, color=color)
        else:
            fig = plot_functions[visualization_type](df, x=x_axis, y=y_axis)
        st.plotly_chart(fig, use_container_width=True)
        add_html_download_button(fig, x_axis, y_axis)
    else:
        st.warning("Unable to create the selected visualization.")

def add_html_download_button(fig, x_axis: str, y_axis: List[str]):
    """Add a download button for the chart as HTML."""
    html = fig.to_html(include_plotlyjs='cdn')
    st.download_button(
        label="Download Chart as HTML",
        data=html,
        file_name=f"{x_axis}_{'_'.join(y_axis)}_chart.html",
        mime="text/html"
    )

# Main application logic
if uploaded_file is not None:
    df = read_excel_file(uploaded_file)
    if df is not None:
        st.header('Data Preview')
        st.dataframe(df)

        # Automatic analysis option
        if auto_analysis:
            st.subheader("📊 Automatic Analysis")
            charts = automatic_chart_analysis(df)
            for idx, chart in enumerate(charts):
                st.plotly_chart(chart, use_container_width=True)
                add_html_download_button(chart, f"chart_{idx + 1}", ["auto"])

            # Button to create PPT with data summaries
            if st.button("📤 Generate PPT with Data Summaries"):
                try:
                    with st.spinner("Creating PowerPoint..."):
                        pptx_file = create_ppt(charts, df)
                        with open(pptx_file, "rb") as f:
                            st.download_button("📥 Download PPT", f, file_name=pptx_file)
                    st.success("PPT generated successfully!")
                except Exception as e:
                    st.error(f"Failed to create PPT: {e}")

            # CSV download for data
            if st.button("📤 Download Data as CSV"):
                csv = df.to_csv(index=False)
                st.download_button(
                    label="Download CSV",
                    data=csv,
                    file_name="chart_data.csv",
                    mime="text/csv"
                )
        else:
            st.header('📈 Manual Visualization')
            x_axis_column = st.selectbox('Select the X-axis column', df.columns)
            y_axis_columns = st.multiselect('Select the Y-axis columns', df.columns)
            color_column = st.selectbox('Select color column (optional)', [None] + list(df.columns))

            if y_axis_columns:
                visualization_type = st.selectbox('Select Visualization Type', [
                    'Bar Chart', 'Line Chart', 'Scatter Plot', 'Pie Chart', 'Donut Chart', 'Bubble Chart', 'Area Chart',
                    'Radar Chart', 'Mixed Chart', 'Funnel Chart'
                ])
                create_plotly_chart(df, visualization_type, x_axis_column, y_axis_columns, color=color_column)
            else:
                st.warning("Please select at least one Y-axis column.")
else:
    st.info('📁 Please upload an Excel file to get started.')