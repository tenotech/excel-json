# excel-json
Format excel to json and json to excel.

## Prerequisites

Before you begin, ensure you have met the following requirements:

- [Node.js](https://nodejs.org/) installed on your machine.

## Installation

1. Clone the repository to your local machine.

    ```bash
    git clone https://github.com/tenotech/excel-json.git
    ```

2. Navigate to the project directory.

    ```bash
    cd excel-json
    ```

3. Install the necessary dependencies.

    ```bash
    npm install
    ```

## Usage

1. Open the `paths.json` file in the project root.

    ```json
    {
      "excel": "path/to/your/excel/file.xlsx",
      "json": "path/to/your/output/file.json"
    }
    ```

    Update the `"excel"` and `"json"` values with the correct paths for your Excel file and desired output JSON file.
    
    This will ensure that the project uses the correct paths for the Excel input and JSON output during development.


2. Save the `paths.json` file.

3. Run the `excel-to-json.ts` script to convert your Excel file to JSON.

    ```bash
    npm run excel-to-json
    ```

    This will generate an `output.json` file in the specified output path.

5. Make modifications to the `output.json` file according to your requirements:

   Open Chatgpt and request this request:
    ```
    I have Excel data representing a backlog, which I have formatted into a JSON file in the following structure:
    {
        "Sheet1": [
            {
            "coordinates": {
                "row": 1,
                "column": 1
            },
            "data": "data 1"
            },
            {
            "coordinates": {
                "row": 1,
                "column": 2
            },
            "data": "data 2"
            },
            // ...
        ]
    }

    I've some mistakes in terms of language, spelling, and grammar. Could you please review and correct them? Here is the JSON file: (provide the content of the JSON file).
    ```
    You can ask anything, this json data is like the excel data.

3. Save the modified `output.json` file.

4. Run the `modify-excel-from-json.ts` script to update the original Excel file based on your modifications.

    ```bash
    npm run update-excel
    ```

    This will update the Excel file with the changes you made in the `output.json` file.



