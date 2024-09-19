import glob

files=glob.glob("data/Attendance_xlsx/*.xlsx",recursive=True)
for file in files:
    print(file)