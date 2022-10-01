# pesu-ics-calendar-export
Export the PESU Calendar of Events from a PDF to a `.ics` file

Allows you to import the PESU Calendar to your favourite calendar application - Google Calendar, Apple Calendar, Outlook, etc.

## Usage
1. Download the PDF from the PESU website/email
2. Clone this repository
3. Create a virtual environment using any method listed and install the requirements
    - `conda`
        ```bash
        conda create -n pesu-ics-calendar-export python=3.9
        conda activate pesu-ics-calendar-export
        pip install -r requirements.txt
        ```
    - `virtualenv`   
        ```bash
        virtualenv venv
        source venv/bin/activate
        pip install -r requirements.txt
        ```
4. Run the script to export the calendar
    ```bash
    PESU ICS Calendar Export [-h] -i INPUT [-o OUTPUT]

    optional arguments:
    -h, --help            show this help message and exit
    -i INPUT, --input INPUT
                            Input calendar PDF file
    -o OUTPUT, --output OUTPUT
                            Output calendar ics file
    ```
    Example: 
    ```bash
    python calendar2ics.py -i data/calendar.pdf -o data/calendar.ics
    ```
5. Import the `.ics` file into your calendar application