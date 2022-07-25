# Generate activity log

This program harvests metadata related to sampling activities from
IMR's onboard Toktlogger and dumps them into an XLSX template that
is compliant with the Nansen Legacy logging system.

This work is a development of the Nansen Legacy labelling system, developed Pål Ellingsen and colleagues:
https://github.com/SIOS-Svalbard/AeN_data
https://github.com/SIOS-Svalbard/darwinsheet
https://doi.org/10.5334/DSJ-2021-034

## Setup and Installation

This application was developed with Python version 3.8.10

```
git clone <repo-url>

pip install -r requirements.txt
```

## Running the application

```
./main.py
```

The application creates a file named 'activty_log_n.xlsx' where 'n' is
a number denoting the version of the file.
