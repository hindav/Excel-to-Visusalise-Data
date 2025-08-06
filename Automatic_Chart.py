import pandas as pd
import plotly.express as px
from io import BytesIO
from pptx import Presentation
from pptx.util import Inches

def automatic_chart_analysis(df: pd.DataFrame):
    """Automatically generate charts based on the analysis of the DataFrame."""
    charts = []  # List to hold generated charts

    # Identify numerical and categorical columns
    numerical_cols = df.select_dtypes(include=['number']).columns.tolist()
    categorical_cols = df.select_dtypes(include=['object', 'category']).columns.tolist()

    # Generate automatic visualizations
    if numerical_cols and categorical_cols:
        for cat_col in categorical_cols:
            for num_col in numerical_cols:
                fig = px.bar(df, x=cat_col, y=num_col, title=f"{cat_col} vs {num_col}")
                charts.append(fig)

    # Generate histograms for numerical columns
    for num_col in numerical_cols:
        fig = px.histogram(df, x=num_col, title=f"Histogram of {num_col}")
        charts.append(fig)

    # Generate scatter plots for pairs of numerical columns
    for i in range(len(numerical_cols)):
        for j in range(i + 1, len(numerical_cols)):
            fig = px.scatter(df, x=numerical_cols[i], y=numerical_cols[j], title=f"{numerical_cols[i]} vs {numerical_cols[j]}")
            charts.append(fig)

    # Correlation matrix
    if len(numerical_cols) > 1:
        correlation_matrix = df[numerical_cols].corr()
        fig = px.imshow(correlation_matrix, text_auto=True, title="Correlation Matrix")
        charts.append(fig)

    return charts

def create_ppt(charts: list, file_path: str = "automatic_chart_analysis.pptx"):
    """Create a PowerPoint presentation from the list of charts."""
    prs = Presentation()

    for idx, fig in enumerate(charts):
        try:
            # Save the figure to a BytesIO object
            img_stream = BytesIO()
            fig.write_image(img_stream, format='png')  # Requires kaleido
            img_stream.seek(0)

            # Add a slide with a title and content layout
            slide = prs.slides.add_slide(prs.slide_layouts[5])  # Title only layout
            slide.shapes.title.text = f"Chart {idx+1}"

            # Add the image to the slide
            left = Inches(1)
            top = Inches(1.5)
            slide.shapes.add_picture(img_stream, left, top, width=Inches(8))
        except Exception as e:
            print(f"Error adding chart {idx+1}: {e}")

    # Save the presentation to a file
    prs.save(file_path)
    return file_path
