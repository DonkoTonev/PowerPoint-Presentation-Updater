from pptx import Presentation

def extract_text_from_first_slide(pptx_file):
    # Load the PowerPoint presentation
    presentation = Presentation(pptx_file)

    # Access the first slide
    first_slide = presentation.slides[0]

    # Extract text from shapes in the first slide
    text_content = []
    for shape in first_slide.shapes:
        if hasattr(shape, "text"):
            text_content.append(shape.text)

    return text_content

if __name__ == "__main__":
    # Specify the path to your PowerPoint file
    pptx_file_path = "C:\\Users\\admin\\OneDrive\\Desktop\\Presentation_example4.pptx"

    # Extract text from the first slide
    content = extract_text_from_first_slide(pptx_file_path)

    # Print the extracted text content
    for text in content:
        print(text)
