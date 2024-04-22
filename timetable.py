import numpy
from ics import Calendar, Event
from datetime import datetime, timedelta
import pandas as pd


class Exam:
    # date_time_str in form "01 Jan 02:00PM"
    def __init__(self, subject_name: str, start_date_time: datetime, duration, location: str):
        self.subject_name = subject_name
        self.start_date_time = start_date_time
        self.duration = duration
        self.location = location

    def to_event(self):
        exam_event = Event()

        exam_event.name = f"{self.subject_name} Exam"
        exam_event.begin = self.start_date_time
        exam_event.duration = self.duration
        exam_event.location = self.location

        return exam_event


def main():
    file_path = 'exjun24_final_tt.xlsx'
    df = pd.read_excel(file_path, sheet_name='Sheet1', header=4)

    cal = Calendar()

    modules_str = input("Enter list of module codes separated by commas: ").upper()

    modules = map(lambda s: s.strip(), modules_str.split(","))

    print("[Exams]")

    for module_code in modules:
        subject_df = df[df['Paper Code'].str.contains(module_code)]
        subject_df = subject_df[~subject_df['Paper Title'].str.contains('Resit')]

        if subject_df.empty:
            raise ValueError(f"No module with code: {module_code}")

        subject_row = subject_df.iloc[0].to_dict()

        date = subject_row['Date']
        start_date_time = datetime.strptime(f"{date.year}-{date.month}-{date.day} {subject_row['Time']}", "%Y-%m-%d %H:%M:%S")

        split_duration = subject_row['Duration'].split(':')
        duration = timedelta(hours=int(split_duration[0]), minutes=int(split_duration[1]), seconds=int(split_duration[2]))
        cal.events.add(Exam(f"{subject_row['Paper Title']}", start_date_time, duration, subject_row['Room/Platform']).to_event())

        print(f"{module_code}: {subject_row['Paper Title']} | {start_date_time} | {subject_row['Room/Platform']}")

    with (open("timetable.ics", "w")) as time_table:
        time_table.write(cal.serialize())


if __name__ == "__main__":
    main()


