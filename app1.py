
import requests
import tkinter as tk
import openpyxl
import os
import re
from tkinter import messagebox



window = tk.Tk()
window.title('canvas scheduler 1.0.0')
window.geometry('500x200')

label1 = tk.Label(window, text = 'Enter your api token:')
label1.pack()
input1 = tk.Entry(width = 100)
input1.pack()

label2 = tk.Label(window, text = 'Enter your course ID:')
label2.pack()
input2 = tk.Entry()
input2.pack()

label3 = tk.Label(window, text = 'love you baby have fun in school :) <3')
label3.pack()



def generate_schedule():
    # === Setup ===
    API_TOKEN = input1.get().strip() #'1303~2GYMDQyrt8hhP48TzGmL2ZVmmLxu2ACyufPaGECNePyQt8LLE2nBvxXkKx42QZP4'  # Replace this!
    COURSE_ID = input2.get().strip() #'220239'
    BASE_URL = 'https://baylor.instructure.com/api/v1'

    if not API_TOKEN or not COURSE_ID:
        messagebox.showerror("Missing input", "Please enter both API token and Course ID.")
        return

    url = f"{BASE_URL}/courses/{COURSE_ID}/assignment_groups"
    params = {
        "exclude_assignment_submission_types[]": "wiki_page",
        "exclude_response_fields[]": ["description", "rubric"],
        "include[]": ["assignments", "discussion_topic", "assessment_requests"],
        "override_assignment_dates": "true",
        "per_page": 50
    }

    headers = {
        "Authorization": f"Bearer {API_TOKEN}",
        "Accept": "application/json"
    }

    try:
        response = requests.get(url, headers=headers, params=params)
    except Exception as e:
        messagebox.showerror("Network Error", str(e))
        return

    # === Fetch Assignment Data ===
    response = requests.get(url, headers=headers, params=params)
    print("Status Code:", response.status_code)

    if response.status_code != 200:
        print("❌ Failed to fetch assignments.")
        print("Response Text:", response.text)
        exit()

    assignment_groups = response.json()

    # === Determine File Number ===
    def get_next_file_number(prefix='canvas_assignments_with_weights_', ext='.xlsx'):
        existing_files = [f for f in os.listdir('.') if f.startswith(prefix) and f.endswith(ext)]
        max_num = 0
        for filename in existing_files:
            match = re.search(rf"{re.escape(prefix)}(\d+){re.escape(ext)}", filename)
            if match:
                num = int(match.group(1))
                if num > max_num:
                    max_num = num
        return max_num + 1

    file_number = get_next_file_number()
    filename = f"canvas_assignments_with_weights_{file_number}.xlsx"

    # === Write to Excel ===
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Assignments'
    ws.append(['Assignment Name', 'Due Date', 'Due Time', 'Group Weight (%)' 'Assignment Weight (%)'])

    
    for group in assignment_groups:
        group_weight = group.get('group_weight', 0.0)
        assignments = group.get('assignments', [])

        total_group_points = sum(
            a.get('points_possible', 0.0) for a in assignments if a.get('points_possible') is not None
        )

        for assignment in assignments:
            name = assignment.get('name', 'Unnamed')
            due_at = assignment.get('due_at')
            points = assignment.get('points_possible', 0.0)

            if due_at:
                date, time = due_at.split('T')
                time = time.split('Z')[0]
            else:
                date, time = 'No Due Date', 'N/A'

            if total_group_points > 0:
                assignment_weight = (points / total_group_points) * group_weight

            ws.append([name, date, time, group_weight, assignment_weight])

    wb.save(filename)
    status_label.config(text=f"✅ Saved as '{filename}'")


submit = tk.Button(window, text = 'Generate your schedule!', width = 20, command = generate_schedule)
submit.pack()

status_label = tk.Label(window, text="", fg="green")
status_label.pack()


window.mainloop()


