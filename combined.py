from pptx import Presentation
from pptx.enum.chart import XL_CHART_TYPE
from pptx.util import Inches, Pt
from pptx.chart.data import CategoryChartData
import csv


def update_presentation(presentation_path):
    # Open the presentation
    presentation = Presentation(presentation_path)

    # Update the title of the first slide
    first_slide = presentation.slides[0]
    title = first_slide.shapes.title
    
    title.text = "New Title"
    
    subtitle = first_slide.placeholders[1]  # Assuming the subtitle is the second placeholder
    subtitle.text = "New Subtitle"

    # Update the second slide with a new chart
    second_slide = presentation.slides[1]

    for shape in second_slide.shapes:
        if shape.has_chart:
            chart = shape.chart
            chart_data = CategoryChartData()
            chart_data.categories = ['Category 1', 'Category 2', 'Category 3', 'Category 4']
            chart_data.add_series('Series 1', (10, 10, 10, 10))
            chart_data.add_series('Series 2', (10, 10, 10, 10))
            chart_data.add_series('Series 3', (10, 10, 10, 10))
            chart.replace_data(chart_data)
            break

    # Update the third slide with a new table cell
    third_slide = presentation.slides[2]

    # for shape in third_slide.shapes:
    #     if shape.has_table:
    #         cell = shape.table.cell(1, 0)
    #         cell.text = 'Pina Colada'

    #         # Set the font size to a smaller value, e.g., 10 points
    #         cell.text_frame.paragraphs[0].runs[0].font.size = Pt(10)

    #         break
    
    # for shape in third_slide.shapes:
    # # Check if the shape has a table
    #     if shape.has_table:
    #         # Iterate through rows and columns of the table
    #         for row in shape.table.rows:
    #             for cell in row.cells:
    #                 # Set the text in each cell to 'Pina Colada'
    #                 cell.text = 'Pina Colada'

    #                 # Set the font size to 10 points
    #                 cell.text_frame.paragraphs[0].runs[0].font.size = Pt(10)
    
    
    csv_file_path = 'data.csv'  # Update with your CSV file path
    with open(csv_file_path, 'r') as csv_file:
        csv_reader = csv.reader(csv_file)
        table_data = [row for row in csv_reader]

    # Iterate through shapes on the third slide
    for shape in third_slide.shapes:
        # Check if the shape has a table
        if shape.has_table:
            # Iterate through rows and columns of the table
            for i, row in enumerate(shape.table.rows):
                # print(i, ':', row)
                for j, cell in enumerate(row.cells):
                    # print(j, ':', cell)
                    
                    # Replace the text in each cell with data from the CSV file
                    if i < len(table_data) and j < len(table_data[i]):
                        cell.text = table_data[i][j]
                    else:
                        cell.text = ''  # In case the CSV data is smaller than the table

                    # Set the font size to 10 points
                    cell.text_frame.paragraphs[0].runs[0].font.size = Pt(10)
    
    
    

    # Save the updated presentation
    presentation.save("combined_updated_presentation.pptx")


if __name__ == "__main__":
    pptx_file_path = "C:\\Users\\admin\\OneDrive\\Desktop\\Presentation_example4.pptx"
    update_presentation(pptx_file_path)
