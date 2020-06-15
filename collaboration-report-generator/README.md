# Collaboration Report Generator Automation
Python script traverse a folder hierarchy, get collaboration, last login activity, and last file activity.

## Pre-Requisites
1. Ensure you've completed pre-requisites in the [parent project documentation](../README.md)
2. Install dependencies:
    * [dateutil](https://dateutil.readthedocs.io/en/stable/): Used for datetime conversion and parsing utilities.
    * [openpyxl](https://openpyxl.readthedocs.io/en/stable/): Used for create the Excel workbook and spreadsheet.
3. Adjust the [event_types](/collaboration-report-generator/collab_report_generator.py#L11) variable as needed.
4. Adjust the [limit](/collaboration-report-generator/collab_report_generator.py#L14) variable as needed.
    * More details on the [Enterprise Events Stream endpoint.](https://developer.box.com/reference/get-events/#request)
5. Run the collab_report_generator.py Python script with the following parameters
    * --box_config: Path to your JWT public/private configuration json file
    * --parent_folderid: The folder_id for the folder in which you want to begin the traversing process
    * --day_lookback: The integer for the number of days you want to look for new enterprise events being created
    ```
    python3 collab_report_generator.py --box_config /path/to/my/box_config.json --parent_folder_id 123456789 --day_lookback 1
    ```  

## Disclaimer
This project is a collection of open source examples and should not be treated as an officially supported product. Use at your own risk. If you encounter any problems, please log an [issue](https://github.com/kylefernandadams/box-python-automations/issues).

## License

The MIT License (MIT)

Copyright (c) 2020 Kyle Adams

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
