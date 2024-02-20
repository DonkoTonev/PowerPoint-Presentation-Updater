This Python script allows you to update various elements within a PowerPoint presentation using data from an Excel file or CSV. It automates the process of updating the title, subtitle, master slide logo, chart data, table data, and text boxes within the presentation.


## Features

- **Automated Updating**: The script loads a PowerPoint template, updates the slides, chart data, table data, and text boxes according to the provided Excel file or CSV, and then saves the new PowerPoint file.

## Requirements

- Python 3.x
- `python-pptx` library

## Usage

1. Install the required Python libraries:
    ```bash
    pip install python-pptx
    ```

2. Prepare your PowerPoint template (.pptx) and Excel file or CSV containing the updated data.

3. Ensure that your Excel file or CSV has columns corresponding to the elements you want to update (title, subtitle, chart data, table data, etc.).

4. Run the script:
    ```bash
    python powerpoint_updater.py --presentation <presentation_file> --data <data_file>
    ```

    Replace `<presentation_file>` with the path to your PowerPoint template file and `<data_file>` with the path to your Excel file or CSV.

5. The script will update the presentation elements with the data provided in the Excel file or CSV and save the updated PowerPoint file.

## Example

```bash
python powerpoint_updater.py --presentation template.pptx --data data.xlsx
