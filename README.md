# Install requirements

```
pip install -r requirements.txt
```

# Run the script

```
python main.py
```

This script will get the start_date and end_dates at A1, B2 then it will fix the headers and begin looping thru each site
In the end an excel will be built as output "output_31_days_report.xlsx" with the correct proposed format

# Assumptions:

I assume that we will always receive the excel in the given format (considering start_dates and end_dates at A1, A2)
I assume that no other params will be added (Needs a fix to include all params dinamically)
I assume that no frameworks are needed for this challenge, only pure functional python (I'm not making use of OOP for this script)
