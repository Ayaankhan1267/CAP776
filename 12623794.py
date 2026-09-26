import pandas as pd

class InvalidActivityDataError(Exception):
    pass

class DailyActivity:

    def __init__(
        self,
        date,
        sleep,
        fitness,
        study,
        coding,
        class_time,
        classes_attended,
        other_activities,
        total_tracked,
        free_time,
        feeling,
        satisfaction,
        energy,
        notes
    ):
        self.date = date
        self.sleep = sleep
        self.fitness = fitness
        self.study = study
        self.coding = coding
        self.class_time = class_time
        self.classes_attended = classes_attended
        self.other_activities = other_activities
        self.total_tracked = total_tracked
        self.free_time = free_time
        self.feeling = feeling
        self.satisfaction = satisfaction
        self.energy = energy
        self.notes = notes

    def feeling_score(self):

        scores = {
            "Excellent": 5,
            "Good": 4,
            "Neutral": 3,
            "Low": 2,
            "Stressed": 1
        }

        return scores.get(str(self.feeling).strip(), None) #1

    def satisfaction_score(self):

        scores = {
            "Very Satisfied": 5,
            "Satisfied": 4,
            "Neutral": 3,
            "Unsatisfied": 2,
            "Very Unsatisfied": 1
        }

        return scores.get(str(self.satisfaction).strip(), None)

    def energy_score(self):

        scores = {
            "High": 3,
            "Medium": 2,
            "Low": 1
        }

        return scores.get(str(self.energy).strip(), None)

    def experience_score(self): 

        feeling = self.feeling_score()
        satisfaction = self.satisfaction_score()
        energy = self.energy_score()

        if feeling is None or satisfaction is None or energy is None:
            return None

        return (feeling + satisfaction + energy) / 3
    
def validate_activity(activity):
    time_values = [
        activity.sleep,
        activity.fitness,
        activity.study,
        activity.coding,
        activity.class_time,
        activity.other_activities
    ]

    for value in time_values:
        if value < 0:
            raise InvalidActivityDataError(
                "Time values cannot be negative."
            )

    if activity.classes_attended < 0:
        raise InvalidActivityDataError(
            "Classes attended cannot be negative."
        )

def read_excel_data(filename):

    df = pd.read_excel(           
        filename,
        sheet_name="Daily Log",
        header=4
    )

    valid_records = []
    invalid_records = 0

    for _, row in df.iterrows():    

        if pd.isna(row["Date"]):
            continue

        try:

            date = pd.to_datetime(row["Date"])

            sleep = float(row["Sleep (min)"])
            fitness = float(row["Fitness (min)"])
            study = float(row["Study (min)"])
            coding = float(row["Coding (min)"])
            class_time = float(row["Class (min)"])
            classes_attended = float(row["Classes Attended"])
            other_activities = float(row["Other Activities (min)"])

            total_tracked = (
                sleep
                + fitness
                + study
                + coding
                + class_time
                + other_activities
            )

            free_time = 1440 - total_tracked

            feeling = row["Day's Feeling"]
            satisfaction = row["Satisfaction Level"]
            energy = row["Energy Level"]

            notes = ""

            if not pd.isna(row["Notes"]):
                notes = str(row["Notes"])

            durations = [
                sleep,
                fitness,
                study,
                coding,
                class_time,
                other_activities
            ]

            if any(value < 0 for value in durations):
                raise InvalidActivityDataError(
                    "Negative activity time"
                )

            if classes_attended < 0:
                raise InvalidActivityDataError(
                    "Negative class attendance"
                )

            if total_tracked > 1440:
                raise InvalidActivityDataError(
                    "Total tracked time is greater than 1440 minutes"
                )

            temp_activity = DailyActivity(
                date,
                sleep,
                fitness,
                study,
                coding,
                class_time,
                classes_attended,
                other_activities,
                total_tracked,
                free_time,
                feeling,
                satisfaction,
                energy,
                notes
            )
            validate_activity(temp_activity)

            if temp_activity.feeling_score() is None:
                raise InvalidActivityDataError(
                    "Invalid feeling value"
                )

            if temp_activity.satisfaction_score() is None:
                raise InvalidActivityDataError(
                    "Invalid satisfaction value"
                )

            if temp_activity.energy_score() is None:
                raise InvalidActivityDataError(
                    "Invalid energy value"
                )

            valid_records.append(temp_activity)

        except (ValueError, TypeError, InvalidActivityDataError):

            invalid_records += 1

    return valid_records, invalid_records

def calculate_average(values):

    if len(values) == 0:
        return 0

    return sum(values) / len(values)


def calculate_total(values):

    return sum(values)

def calculate_activity_summary(activities):

    sleep = [a.sleep for a in activities]
    fitness = [a.fitness for a in activities]
    study = [a.study for a in activities]
    coding = [a.coding for a in activities]
    class_time = [a.class_time for a in activities]
    other = [a.other_activities for a in activities]
    free_time = [a.free_time for a in activities]
    total_tracked = [a.total_tracked for a in activities]

    summary = {

        "Valid Days": len(activities),

        "Average Sleep": calculate_average(sleep),

        "Average Fitness": calculate_average(fitness),

        "Average Study": calculate_average(study),

        "Average Coding": calculate_average(coding),

        "Average Class": calculate_average(class_time),

        "Average Other Activities":
            calculate_average(other),

        "Average Free / Unaccounted Time":
            calculate_average(free_time),

        "Average Total Tracked Time":
            calculate_average(total_tracked)
    }

    return summary

def calculate_indices(activities):

    coding_values = [
        activity.coding
        for activity in activities
    ]

    TPI = calculate_average(coding_values)

    academic_values = [
        activity.study + activity.class_time
        for activity in activities
    ]

    AAI = calculate_average(academic_values)

    fitness_values = [
        activity.fitness
        for activity in activities
    ]

    PhAI = calculate_average(fitness_values)

    sleep_values = [
        activity.sleep
        for activity in activities
    ]

    SRI = calculate_average(sleep_values)

    free_values = [
        activity.free_time
        for activity in activities
    ]

    ABI = calculate_average(free_values)

    tracked_values = [
        activity.total_tracked
        for activity in activities
    ]

    TUI = calculate_average(tracked_values)

    experience_values = []

    for activity in activities:

        score = activity.experience_score()

        if score is not None:
            experience_values.append(score)

    EI = calculate_average(experience_values)

    dates = [
        activity.date
        for activity in activities
    ]

    first_date = min(dates)
    last_date = max(dates)

    expected_days = (
        last_date - first_date
    ).days + 1

    valid_days = len(activities)

    DCI = (
        valid_days / expected_days
    ) * 100

    PAI = (
        0.15 * TPI
        + 0.20 * AAI
        + 0.15 * PhAI
        + 0.20 * SRI
        + 0.15 * TUI
        + 0.10 * EI
        + 0.05 * DCI
    )

    return {

        "TPI": TPI,
        "AAI": AAI,
        "PhAI": PhAI,
        "SRI": SRI,
        "ABI": ABI,
        "TUI": TUI,
        "EI": EI,
        "DCI": DCI,
        "PAI": PAI,

        "Expected Days": expected_days,
        "Valid Days": valid_days
    }

def calculate_correlations(activities):

    data = {

        "Sleep": [
            activity.sleep
            for activity in activities
        ],

        "Energy": [
            activity.energy_score()
            for activity in activities
        ],

        "Study": [
            activity.study
            for activity in activities
        ],

        "Satisfaction": [
            activity.satisfaction_score()
            for activity in activities
        ],

        "Coding": [
            activity.coding
            for activity in activities
        ]
    }

    df = pd.DataFrame(data)

    sleep_energy = df["Sleep"].corr(
        df["Energy"]
    )

    study_satisfaction = df["Study"].corr(
        df["Satisfaction"]
    )

    coding_energy = df["Coding"].corr(
        df["Energy"]
    )

    return {

        "Sleep ↔ Energy": sleep_energy,

        "Study ↔ Satisfaction":
            study_satisfaction,

        "Coding ↔ Energy":
            coding_energy
    }

def correlation_note(value):

    if value >= 0.7:
        return "Strong positive relationship"

    elif value >= 0.3:
        return "Moderate positive relationship"

    elif value > -0.3:
        return "Weak or little relationship"

    elif value > -0.7:
        return "Moderate negative relationship"

    else:
        return "Strong negative relationship"
    
def print_activity_summary(summary, invalid_records):

    print("\n")
    print("=" * 70)
    print("1. ACTIVITY DATA SUMMARY")
    print("=" * 70)

    print(
        f"Valid days recorded           : "
        f"{summary['Valid Days']}"
    )

    print(
        f"Invalid / excluded records    : "
        f"{invalid_records}"
    )

    print(
        f"Average Sleep/day             : "
        f"{summary['Average Sleep']:.2f} min"
    )

    print(
        f"Average Fitness/day           : "
        f"{summary['Average Fitness']:.2f} min"
    )

    print(
        f"Average Study/day             : "
        f"{summary['Average Study']:.2f} min"
    )

    print(
        f"Average Coding/day            : "
        f"{summary['Average Coding']:.2f} min"
    )

    print(
        f"Average Class/day             : "
        f"{summary['Average Class']:.2f} min"
    )

    print(
        f"Average Other Activities/day : "
        f"{summary['Average Other Activities']:.2f} min"
    )

    print(
        f"Average Free/Unaccounted/day  : "
        f"{summary['Average Free / Unaccounted Time']:.2f} min"
    )

def print_indices(indices):

    print("\n")
    print("=" * 70)
    print("2. INDEX VALUES")
    print("=" * 70)

    print(
        f"TPI  - Tech Productivity       : "
        f"{indices['TPI']:.2f} min/day"
    )

    print(
        f"AAI  - Academic Activity       : "
        f"{indices['AAI']:.2f} min/day"
    )

    print(
        f"PhAI - Physical Activity       : "
        f"{indices['PhAI']:.2f} min/day"
    )

    print(
        f"SRI  - Sleep & Recovery        : "
        f"{indices['SRI']:.2f} min/day"
    )

    print(
        f"ABI  - Activity Balance        : "
        f"{indices['ABI']:.2f} min/day"
    )

    print(
        f"TUI  - Time Utilization        : "
        f"{indices['TUI']:.2f} min/day"
    )

    print(
        f"EI   - Experience Index        : "
        f"{indices['EI']:.2f} / 5"
    )

    print(
        f"DCI  - Data Continuity         : "
        f"{indices['DCI']:.2f} %"
    )

    print("-" * 70)

    print(
        f"PAI  - Personal Activity Index : "
        f"{indices['PAI']:.2f}"
    )

def print_correlations(correlations):

    print("\n")
    print("=" * 70)
    print("3. KEY FINDINGS / CORRELATIONS")
    print("=" * 70)

    for name, value in correlations.items():

        print(
            f"{name:<30} : "
            f"{value:.3f}"
        )

def save_results(
    activities,
    summary,
    indices,
    correlations,
    invalid_records,
    filename
):

    daily_data = []

    for activity in activities:

        daily_data.append({

            "Date": activity.date,

            "Sleep (min)": activity.sleep,

            "Fitness (min)": activity.fitness,

            "Study (min)": activity.study,

            "Coding (min)": activity.coding,

            "Class (min)": activity.class_time,

            "Classes Attended":
                activity.classes_attended,

            "Other Activities (min)":
                activity.other_activities,

            "Total Tracked (min)":
                activity.total_tracked,

            "Free/Unaccounted (min)":
                activity.free_time,

            "Feeling": activity.feeling,

            "Feeling Score":
                activity.feeling_score(),

            "Satisfaction":
                activity.satisfaction,

            "Satisfaction Score":
                activity.satisfaction_score(),

            "Energy": activity.energy,

            "Energy Score":
                activity.energy_score(),

            "Experience Score":
                activity.experience_score(),

            "Notes": activity.notes
        })

    daily_df = pd.DataFrame(daily_data)

    summary_data = {

        "Expected Days":
            indices["Expected Days"],

        "Valid Days":
            indices["Valid Days"],

        "Invalid / Excluded Records":
            invalid_records,

        "Average Sleep (min/day)":
            summary["Average Sleep"],

        "Average Fitness (min/day)":
            summary["Average Fitness"],

        "Average Study (min/day)":
            summary["Average Study"],

        "Average Coding (min/day)":
            summary["Average Coding"],

        "Average Class (min/day)":
            summary["Average Class"],

        "Average Other Activities (min/day)":
            summary["Average Other Activities"],

        "Average Free / Unaccounted (min/day)":
            summary["Average Free / Unaccounted Time"],

        "TPI":
            indices["TPI"],

        "AAI":
            indices["AAI"],

        "PhAI":
            indices["PhAI"],

        "SRI":
            indices["SRI"],

        "ABI":
            indices["ABI"],

        "TUI":
            indices["TUI"],

        "EI":
            indices["EI"],

        "DCI":
            indices["DCI"],

        "PAI":
            indices["PAI"],

        "Sleep ↔ Energy":
            correlations["Sleep ↔ Energy"],

        "Study ↔ Satisfaction":
            correlations["Study ↔ Satisfaction"],

        "Coding ↔ Energy":
            correlations["Coding ↔ Energy"]
    }

    summary_df = pd.DataFrame(
        list(summary_data.items()),
        columns=["Parameter", "Value"]
    )

    with pd.ExcelWriter(filename) as writer:

        daily_df.to_excel(
            writer,
            sheet_name="Calculated Daily Data",
            index=False
        )

        summary_df.to_excel(
            writer,
            sheet_name="Project Results",
            index=False
        )

def main():

    input_file = "12623794.xlsx"

    output_file = "12623794_results.xlsx"

    try:

        activities, invalid_records = read_excel_data(
            input_file
        )

        if len(activities) == 0:

            print("No valid activity data found.")

            return

        summary = calculate_activity_summary(
            activities
        )

        indices = calculate_indices(
            activities
        )

        correlations = calculate_correlations(
            activities
        )
        print("\nCORRELATION OBSERVATIONS")
        print("-" * 50)

        for relationship, value in correlations.items():
            print(
                f"{relationship}: {value:.2f} - "
                f"{correlation_note(value)}"
            )

        print_activity_summary(
            summary,
            invalid_records
        )

        print_indices(
            indices
        )

        print_correlations(
            correlations
        )

        save_results(
            activities,
            summary,
            indices,
            correlations,
            invalid_records,
            output_file
        )

        print("\n")
        print("=" * 70)
        print("RESULTS SAVED SUCCESSFULLY")
        print("=" * 70)

        print(
            f"Output file: {output_file}"
        )

    except FileNotFoundError:

        print(
            f"ERROR: {input_file} was not found."
        )

    except Exception as error:

        print(
            f"ERROR: {error}"
        )

if __name__ == "__main__":
    main()