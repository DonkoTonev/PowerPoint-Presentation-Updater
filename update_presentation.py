from pptx import Presentation
from pptx.enum.chart import XL_CHART_TYPE
from pptx.util import Inches

def update_presentation(presentation_path):
    # Open the presentation
    presentation = Presentation(presentation_path)

    # Update the title of the first slide
    first_slide = presentation.slides[0]
    title = first_slide.shapes.title
    
    title.text = "New Title"
    
    subtitle = first_slide.placeholders[1]  # Assuming the subtitle is the second placeholder
    subtitle.text = "New Subtitle"    
    

    # Find and update a chart in the presentation
    # for slide in presentation.slides:
    #     for shape in slide.shapes:
    #         if shape.has_text_frame:
    #             text_frame = shape.text_frame
    #             if "Chart" in text_frame.text:
    #                 # Assume the chart is a bar chart for this example
    #                 chart_data = [
    #                     ("Category 1", 10),
    #                     ("Category 2", 20),
    #                     ("Category 3", 15),
    #                 ]
    #                 update_chart_data(shape, chart_data)

    # Save the updated presentation
    presentation.save("updated_presentation.pptx")

# def update_chart_data(chart_shape, data):
#     # Assuming the chart type is a bar chart for this example
#     chart = chart_shape.chart
#     chart_type = chart.chart_type
#     if chart_type == XL_CHART_TYPE.BAR_CLUSTERED:
#         # Clear existing data
#         chart.series[0].points.clear()

#         # Add new data
#         for label, value in data:
#             chart.series[0].points.add(category=label, values=(value,))

# Update the path to your presentation file
presentation_path = "C:\\Users\\admin\\OneDrive\\Desktop\\Presentation_example4.pptx"
update_presentation(presentation_path)
