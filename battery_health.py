"""
Generate Windows battery report, extract data, and create an information dashboard to 
help the user understand their battery health.

Created on 12/11/2021

@author: Bryce Smith (brycegsmith@hotmail.com)
"""

import os
import codecs
from datetime import datetime, timedelta
import matplotlib.pyplot as plt


def main():
    """
    Main Function.

    Parameters:
        None

    Returns:
        None

    """

    # Generate and get report
    generate_battery_report()
    report = get_battery_report()

    # Extract information
    information = get_battery_information(report)
    capacity_data = get_capacity_data(report)
    usage_data = get_usage_data(report)

    # Generate dashboard
    generate_dashboard(information, capacity_data, usage_data)


def generate_battery_report() -> None:
    """
    Function to create Windows battery report and save it in C drive.

    Parameters:
        None

    Returns:
        None

    """

    # Command to save battery-report.html to current directory
    cwd = os.getcwd()
    cmd = r'powercfg /batteryreport /output "' + cwd + r'\battery-report.html"'

    # Execute command
    os.system(cmd)


def get_battery_report() -> str:
    """
    Function to get Windows battery report (HTML) as a string.

    Parameters:
        None

    Returns:
        report (str): HTML battery report as a string

    """

    file = codecs.open("battery-report.html", "r", "utf-8")
    report = file.read()

    return report


def get_battery_information(report: str) -> list:
    """
    Function to get number of installed batteries and information about batteries.

    Output Format:
        battery_count (int): number of batteries
        names (list of strs): list of name of each battery
        manufacturers (list of strs): list of manufacturer of each battery
        serials (list of strs): list of serial of each battery
        chemistries (list of strs): list of chemistry of each battery
        design_capacities (list of ints): list of design capacity of each battery in mWh
        actual_capacities (list of ints): list of actual capacity of each battery in mWh
        cycles (list of ints): list of cycle count of each battery

    Parameters:
        report (str): HTML battery report as a string

    Returns:
        information (list): list with battery infomation following format above

    """

    # Get section of report with battery information
    battery_info = report.split("Installed batteries")[1].split("Recent usage")[0]

    # Split into list of rows (first row is useless)
    table = __extract_html_table(battery_info)
    battery_count = len(table[0]) - 1
    names = table[1][1:]
    manufacturers = table[2][1:]
    serials = table[3][1:]
    chemistries = table[4][1:]

    design_capacities = table[5][1:]
    for i in range(len(design_capacities)):
        design_capacities[i] = __str_to_int(design_capacities[i])

    actual_capacities = table[6][1:]
    for i in range(len(actual_capacities)):
        actual_capacities[i] = __str_to_int(actual_capacities[i])

    cycles = table[7][1:]
    for i in range(len(cycles)):
        cycles[i] = __str_to_int(cycles[i])

    # Combine into information
    information = [
        battery_count,
        names,
        manufacturers,
        serials,
        chemistries,
        design_capacities,
        actual_capacities,
        cycles,
    ]

    return information


def get_capacity_data(report: str) -> list:
    """
    Function to get capacity vs date data for batteries.

    Parameters:
        report (str): HTML battery report as a string

    Returns:
        capacity_data (list): [[list of dates], [list of capacities as ints]]

    """

    # Get section of report with battery information
    capacity_info = report.split("Battery capacity history")[1]
    capacity_info = capacity_info.split("Battery life estimates")[0]

    # Split into list of rows
    table = __extract_html_table(capacity_info)

    # Iterate through table rows (first row is useless)
    dates = []
    capacities = []
    for i in range(1, len(table)):
        datetime = __extract_date(table[i][0])
        dates.append(datetime)

        capacity = __str_to_int(table[i][1])
        capacities.append(capacity)

    capacity_data = [dates, capacities]

    return capacity_data


def get_usage_data(report: str) -> list:
    """
    Function to get cumulative active usage hours on battery power vs date data for
    batteries.

    Parameters:
        report (str): HTML battery report as a string

    Returns:
        usage_data (list): [[list of dates], [list of hours as ints]]

    """

    # Get section of report with battery information
    usage_info = report.split("Usage history")[1]
    usage_info = usage_info.split("Battery capacity history")[0]

    # Split into list of rows (first row is useless)
    table = __extract_html_table(usage_info)

    # Iterate through table rows (first two rows are useless)
    dates = []
    usage_hours = []
    cumulative_usage = timedelta(seconds=0)
    for i in range(2, len(table)):
        # Only append if active usage data is present
        if len(table[i][1]) > 1:
            # Extracting active seconds on battery powers
            split_time = table[i][1].split(":")
            active_hours = int(split_time[0])
            active_minutes = int(split_time[1]) + (active_hours * 60)
            active_seconds = int(split_time[2]) + (active_minutes * 60)

            # Handle Windows bug (usage time should never be > 1 week)
            if active_seconds <= (7 * 24 * 60 * 60):
                # Append to usage hours
                cumulative_usage += timedelta(seconds=active_seconds)
                usage_hours.append(int(cumulative_usage.total_seconds() / 60 / 60))

                # Extract date and append to dates
                date = __extract_date(table[i][0])
                dates.append(date)

    capacity_data = [dates, usage_hours]

    return capacity_data


def generate_dashboard(
    information: list, capacity_data: list, usage_data: list
) -> None:
    """
    Generate an information dashboard to help the user understand their battery health.

    Parameters:
        information (list): list with battery infomation
        capacity_data (list): [[list of dates], [list of capacities as ints]]
        usage_data (list): [[list of dates], [list of hours as ints]]

    Returns:
        None
    """

    # Figure Formatting
    fig = plt.figure()
    fig.set_size_inches(14, 7.5)
    fig.suptitle("Battery Health Dashboard", size=22)

    # Spacing between subplots
    plt.subplots_adjust(
        left=0.1, bottom=0.1, right=0.9, top=0.9, wspace=0.4, hspace=0.4
    )

    import numpy as np

    """
    battery_count (int): number of batteries
    names (list of strs): list of name of each battery
    manufacturers (list of strs): list of manufacturer of each battery
    serials (list of strs): list of serial of each battery
    chemistries (list of strs): list of chemistry of each battery
    design_capacities (list of ints): list of design capacity of each battery in mWh
    actual_capacities (list of ints): list of actual capacity of each battery in mWh
    cycles (list of ints): list of cycle count of each battery
    """

    # Battery Information Table
    plt.subplot(2, 2, 1)
    plt.axis("tight")
    plt.axis("off")

    # Extract and format data
    column_labels = [""]
    names = ["Name"] + information[1]
    manufacturers = ["Manufacturer"] + information[2]
    serials = ["Serial Number"] + information[3]
    chemistries = ["Chemistry"] + information[4]
    design_capacities = ["Design Capacity"]
    actual_capacities = ["Current Capacity"]
    capacity_retention = ["Capacity Retention"]
    cycles = ["Cycle Count"] + information[7]

    for i in range(information[0]):
        column_labels += ["Battery " + str(i + 1)]
        design_capacities += [str(information[5][i]) + " mWh"]
        actual_capacities += [str(information[6][i]) + " mWh"]
        retention = round((information[6][i] / information[5][i]), 3) * 100
        capacity_retention += [str(retention) + "%"]

    data = [
        names,
        manufacturers,
        serials,
        chemistries,
        design_capacities,
        actual_capacities,
        capacity_retention,
        cycles,
    ]

    plt.table(cellText=data, colLabels=column_labels, loc="center")

    # Battery Usage Plot
    plt.subplot(2, 2, 2)
    plt.plot(usage_data[0], usage_data[1])
    plt.xticks(rotation=35)
    plt.title("Cumulative Battery Usage", size=12)
    plt.xlabel("Date")
    plt.ylabel("Battery Usage (hrs)")

    # Battery Capacity Plot
    plt.subplot(2, 1, 2)

    total_design_capcity = 0
    for i in range(information[0]):
        total_design_capcity += information[5][i]

    for i in range(len(capacity_data[1])):
        capacity_data[1][i] = (capacity_data[1][i] / total_design_capcity) * 100

    plt.plot(capacity_data[0], capacity_data[1], marker="x", markersize=3)
    plt.axhline(90, label="90% Capacity", color="y", alpha=0.3)
    plt.axhline(80, label="80% Capacity", color="#ff9900", alpha=0.3)
    plt.axhline(70, label="70% Capacity", color="r", alpha=0.3)
    plt.legend(loc="lower left")
    plt.title("Battery Capacity Fade", size=12)
    plt.xlabel("Date")
    plt.ylabel("Capacity (%)")
    plt.xticks(rotation=35)

    plt.show()


def __extract_html_table(html_text: str) -> list:
    """
    Given a string with exactly one embedded HTML table, extract the table.

    Output Format: [[r1c1, r1c2, r1c3, ...], [r2c1, r2c2, r2c3, ...], ...]
                   where r = row and c = column

    Parameters:
        html_text (str): string with excactly one embedded HTML table

    Returns:
        table (list of lists of strs): extracted HTML table
    """

    # Clean table text
    html_text = html_text.replace(" ", "")
    html_text = html_text.replace("\r", "")
    html_text = html_text.replace("\n", "")
    html_text = html_text.replace("<trstyle", "")

    # Get the table with metadata
    table = html_text.split("<table>")[1]
    table = table.split("</table>")[0]

    # Drop metadata
    table = table.split("<tr")  # Works for <tr> and <trclass= ...>
    table = table[1:]

    # Extract data for every row
    for i in range(len(table)):
        tr_end = table[i].find(">")
        table[i] = table[i][tr_end + 1 :]
        table[i] = table[i].split("</tr>")[0]
        table[i] = __extract_row_data(table[i])

    return table


def __extract_row_data(row: str) -> list:
    """
    Extract individual data entires from the row of an HTML formatted table.

    Parameters:
        row (str): row of an HTML table

    Returns:
        data (list of strs): list of extracted data from HTML row

    """

    data = row.split("<td")

    for i in range(len(data)):
        td_end = data[i].find(">")
        data[i] = data[i][td_end + 1 :]
        data[i] = data[i].replace("</td>", "")

    data = data[1:]

    return data


def __str_to_int(string: str) -> int:
    """
    Given any string get all numbers and combine to form an integer.

    Examples:
        "1.109" -> 1109
        "0993mWh"-> 993

    Parameters:
        string (str): any string

    Returns:
        integer (int): an integer extracted from the string

    """

    integer_string = ""
    for i in range(len(string)):
        if string[i].isdigit():
            integer_string += string[i]

    integer = int(integer_string)

    return integer


def __extract_date(date_str: str) -> int:
    """
    Given a string with either a single date or a date range, return a single date. If
    given a date range, the midpoint will be returned. Dates must follow format used in
    Windows battery report w/ stripped spaces.

    Examples:
        "2018-01-05-2018-01-07" -> datetime(2018, 01, 06)
        "2021-06-20" -> datetime(2021, 06, 20)

    Parameters:
        date_str (str): a string with a date or date range

    Returns:
        date (datetime): extracted date from string

    """

    dash_count = date_str.count("-")

    # If single date
    if dash_count == 2:
        date = datetime.strptime(date_str, "%Y-%m-%d").date()

    # If date range
    elif dash_count == 5:
        date_1 = datetime.strptime(date_str[:10], "%Y-%m-%d").date()
        date_2 = datetime.strptime(date_str[11:21], "%Y-%m-%d").date()

        date = date_1 + (date_2 - date_1) / 2

    return date


if __name__ == "__main__":
    main()
