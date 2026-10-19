import openpyxl
from datetime import datetime, date, timedelta

FILE_NAME = "12602713.xlsx"
SHEET_NAME = "Daily Log"

START_DATE = date(2026, 8, 13)
END_DATE = date(2026, 9, 21)
EXPECTED_DAYS = 40


def rating(value, kind):
    if kind == "feeling":
        if value == "Excellent":
            return 5
        elif value == "Good":
            return 4
        elif value == "Neutral":
            return 3
        elif value == "Low":
            return 2
        elif value == "Stressed":
            return 1

    elif kind == "satisfaction":
        if value == "Very Satisfied":
            return 5
        elif value == "Satisfied":
            return 4
        elif value == "Neutral":
            return 3
        elif value == "Unsatisfied":
            return 2
        elif value == "Very Unsatisfied":
            return 1

    elif kind == "energy":
        if value == "High":
            return 5
        elif value == "Medium":
            return 3
        elif value == "Low":
            return 1

    return None


def get_date(value):
    if isinstance(value, datetime):
        return value.date()
    elif isinstance(value, date):
        return value
    return None


def read_file():
    wb = openpyxl.load_workbook(FILE_NAME, data_only=True)
    ws = wb[SHEET_NAME]

    data = []
    invalid = 0

    for row in ws.iter_rows(min_row=6, values_only=True):
        if row[0] is None:
            continue

        try:
            d = get_date(row[0])

            if d is None or d < START_DATE or d > END_DATE:
                invalid += 1
                continue

            sleep = float(row[1])
            fitness = float(row[2])
            study = float(row[3])
            coding = float(row[4])
            class_time = float(row[5])
            tracked = float(row[8])
            free_time = float(row[9])

            feeling = rating(row[10], "feeling")
            satisfaction = rating(row[11], "satisfaction")
            energy = rating(row[12], "energy")

            if sleep < 0 or fitness < 0 or study < 0:
                invalid += 1
                continue

            if coding < 0 or class_time < 0:
                invalid += 1
                continue

            if tracked < 0 or free_time < 0:
                invalid += 1
                continue

            if feeling is None or satisfaction is None or energy is None:
                invalid += 1
                continue

            data.append([
                d, sleep, fitness, study, coding,
                class_time, tracked, free_time,
                feeling, satisfaction, energy
            ])

        except (ValueError, TypeError):
            invalid += 1

    return data, invalid


def check_continuity(data):
    print("\nDATA CONTINUITY")
    print("----------------")

    dates = []
    i = 0

    while i < len(data):
        d = data[i][0]

        duplicate = False
        j = 0
        while j < len(dates):
            if dates[j] == d:
                duplicate = True
            j += 1

        if duplicate:
            print("Duplicate date:", d)
        else:
            dates.append(d)

        if i > 0 and d < data[i - 1][0]:
            print("Out of order date:", d)

        i += 1

    missing = []
    current = START_DATE

    while current <= END_DATE:
        found = False
        j = 0

        while j < len(dates):
            if dates[j] == current:
                found = True
            j += 1

        if not found:
            missing.append(current)

        current = current + timedelta(days=1)

    print("Expected days:", EXPECTED_DAYS)
    print("Recorded valid days:", len(dates))

    if len(missing) == 0:
        print("Missing dates: None")
    else:
        print("Missing dates:")
        for d in missing:
            print(d)

    return len(dates)


def calculate_indices(data, recorded_days):
    sleep_total = 0
    fitness_total = 0
    study_total = 0
    coding_total = 0
    class_total = 0
    tracked_total = 0
    free_total = 0
    experience_total = 0

    

    i = 0
    while i < len(data):
        sleep_total += data[i][1]
        fitness_total += data[i][2]
        study_total += data[i][3]
        coding_total += data[i][4]
        class_total += data[i][5]
        tracked_total += data[i][6]
        free_total += data[i][7]

        

        experience_total += (
            data[i][8] + data[i][9] + data[i][10]
        ) / 3

        i += 1

    if recorded_days == 0:
        print("\nNo valid data available.")
        return

    TPI = coding_total / recorded_days
    AAI = (study_total + class_total) / recorded_days
    PhAI = fitness_total / recorded_days
    SRI = sleep_total / recorded_days
    ABI = free_total / recorded_days
    TUI = tracked_total / recorded_days
    EI = experience_total / recorded_days
    DCI = (recorded_days / EXPECTED_DAYS) * 100
 
    


    PAI = (
        0.15 * TPI +
        0.20 * AAI +
        0.15 * PhAI +
        0.20 * SRI +
        0.15 * TUI +
        0.10 * EI +
        0.05 * DCI
    )

    print("\nACTIVITY INDICES")
    print("----------------")
    print("TPI :", TPI)
    print("AAI :", AAI)
    print("PhAI:", PhAI)
    print("SRI :", SRI)
    print("ABI :", ABI)
    print("TUI :", TUI)
    print("EI  :", EI)
    print("DCI :", DCI)
    print("PAI :", PAI)
    print("Average sleep:" ,SRI)
    print("Average Fitness:", PhAI)
    print("Average Study :", AAI)
    print("Average Coding:", TPI)
    print("Average Class:",(class_total) / recorded_days)
    print("Average free Unaccounted Time:" , (free_total) / recorded_days)


def correlation(data, x_index, y_index):
    n = len(data)

    if n < 2:
        return 0

    x_total = 0
    y_total = 0

    i = 0
    while i < n:
        x_total += data[i][x_index]
        y_total += data[i][y_index]
        i += 1

    x_mean = x_total / n
    y_mean = y_total / n

    numerator = 0
    x_square = 0
    y_square = 0

    i = 0
    while i < n:
        x_difference = data[i][x_index] - x_mean
        y_difference = data[i][y_index] - y_mean

        numerator += x_difference * y_difference
        x_square += x_difference * x_difference
        y_square += y_difference * y_difference

        i += 1

    denominator = (x_square * y_square) ** 0.5

    if denominator == 0:
        return 0

    return numerator / denominator


def relationship_finding(value):
    if value > 0:
        return "Positive relationship"
    elif value < 0:
        return "Negative relationship"
    else:
        return "No relationship"


def show_relationships(data):
    coding_energy = correlation(data, 4, 10)
    sleep_energy = correlation(data, 1, 10)
    study_satisfaction = correlation(data, 3, 9)

    print("\nRELATIONSHIPS")
    print("-------------")

    print("Coding - Energy:", coding_energy)
    print("Finding:", relationship_finding(coding_energy))

    print("\nSleep - Energy:", sleep_energy)
    print("Finding:", relationship_finding(sleep_energy))

    print("\nStudy - Satisfaction:", study_satisfaction)
    print("Finding:", relationship_finding(study_satisfaction))
    



def main():
    try:
        data, invalid = read_file()

        print("MY DATA - MY STORY")
        print("==================")
        print("Valid records:", len(data))
        print("Invalid records:", invalid)

        recorded_days = check_continuity(data)
        calculate_indices(data, recorded_days)
        show_relationships(data)

    except FileNotFoundError:
        print("Excel file not found.")
    except KeyError:
        print("Sheet name 'Daily Log' not found.")
    except Exception as e:
        print("Error:", e)

main()