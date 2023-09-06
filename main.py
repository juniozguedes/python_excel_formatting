import pandas as pd
import datetime
import math


def format_excel_headers(data):
    # Shift the contents of Column A up by one row
    data.iloc[:, 0] = data.iloc[:, 0].shift(-1)

    # Delete the first row
    data = data.iloc[1:]

    # Set the current first row as the header
    data.columns = data.iloc[0]

    # Reset the index
    data = data.reset_index(drop=True)
    return data


def setup():
    # Read the Excel file without a header
    data = pd.read_excel("Analytics Template for Exercise.xlsx", header=None)

    # Extract data from A1 and A2
    start_date = data.iloc[0, 0]
    end_date = data.iloc[1, 0]

    # Fix excel format
    data = format_excel_headers(data)
    # Extract header parameters dynamically from the second row
    headers = data.iloc[0, :].tolist()

    # Initialize an empty set to store unique items
    unique_items = set()
    # Use a list comprehension to filter out duplicates while preserving order
    result_list = [
        x for x in headers if x not in unique_items and not unique_items.add(x)
    ]
    headers = ["Day of the month", "Date"] + result_list
    site_data = {}
    current_site = None
    for index, row in data.iterrows():
        for column, value in row.items():
            # 'column' contains the column name (header) and 'value' contains the cell value
            print(f"Column: {column}, Value: {value}")
            # It means that it's duplicated header, we should skip
            if value == "Site ID":
                break
            elif isinstance(value, str) and value.startswith("site"):
                current_site = value
                # Initialize the site_data dictionary with dynamically extracted headers
                site_data[current_site] = {header: [] for header in headers}
                continue

            if isinstance(value, datetime.datetime):
                site_date = value
                if start_date <= site_date <= end_date:
                    # Check if the date is not already in the "Date" list
                    if site_date not in site_data[current_site]["Date"]:
                        site_data[current_site]["Date"].append(site_date)
                        # Extract the day from the date and append it to "Day of the month"
                        day_of_month = site_date.day
                        site_data[current_site]["Day of the month"].append(day_of_month)

            if isinstance(value, (int, float)):
                if math.isnan(value):
                    continue
                # Only append data values if the value above in the same column is in "Date" list
                if column in site_data[current_site]:
                    # Get the index of the current column
                    current_column_index = row.index.get_loc(column)
                    # Get the value in the cell above in the same column
                    value_above = data.iloc[index - 1, current_column_index]

                    # Convert value_above to a list of datetime objects
                    value_above = value_above.tolist()

                    # Convert site_data[current_site]["Date"] to a list of datetime objects
                    date_list = site_data[current_site]["Date"]

                    # Find matching dates and append the value
                    matching_dates = [date for date in date_list if date in value_above]
                    if matching_dates:
                        site_data[current_site][column].append(value)

    return site_data


if __name__ == "__main__":
    # Step 1: Process Excel data and build site_data dictionary
    site_data = setup()

    # Step 2: Iterate through the dictionary to calculate Site ID counts
    for site, info in site_data.items():
        site_id = "site " + site.split()[1]  # Extract the site ID from the site name
        num_dates = len(info["Date"])
        site_ids = [site_id] * num_dates  # Create a list of site IDs
        info["Site ID"] = site_ids  # Update the "site ID" key with the list

    # Step 3: Create a new DataFrame in sequence
    result_df = pd.DataFrame()
    for site, data in site_data.items():
        site_df = pd.DataFrame(data)
        result_df = pd.concat([result_df, site_df])

    # Save the result DataFrame to a new Excel file
    result_df.to_excel("output_31_days_report.xlsx", index=False)
