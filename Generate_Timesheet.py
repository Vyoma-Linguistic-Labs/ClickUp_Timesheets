# -*- coding: utf-8 -*-
"""
Created on Fri May 26 18:52:02 2023

@author: shank
"""

import tkinter as tk
from tkinter import messagebox
from tkcalendar import Calendar
import pandas as pd
from datetime import datetime, timezone
import requests
import time
import webbrowser
import numpy as np
import pytz
from PIL import Image, ImageTk
from io import BytesIO

__version__ = "v2.0.1"
# Dictionary mapping month names to numbers
month_dict = {
    "Jan": 1, "Feb": 2, "Mar": 3, "Apr": 4, "May": 5, "Jun": 6, "Jul": 7,
    "Aug": 8, "Sep": 9, "Oct": 10, "Nov": 11, "Dec": 12
}

# Define the list of columns to check for NaN
columns_to_check = [
    "Course", "Product", "Proj-Common-Activity", "Proj-Outside-Office",
    "Management-Project", "Technology-Project", "Linguistic-Project",
    "MMedia-Project", "Project-CST", "Sales-Mktg-Project", "Project-ELA",
    "Proj-KidsPersona", "FinAcc-Project", "Website", "SFH-Admin-Project", "Admin-Project"
]

# Create a timezone object for IST
ist_timezone = pytz.timezone('Asia/Kolkata')

# Exchange keys and values
month_flipped = {value: key for key, value in month_dict.items()}
def download_latest(root):
        
    # Define the GitHub repository, file path, and file URL
    repository = "Vyoma-Linguistic-Labs/ClickUp_Timesheets"
    file_path = "Generate_Timesheet.py"  # Replace with the path to the file you want to download
    file_url = f"https://raw.githubusercontent.com/{repository}/main/{file_path}"

    # Define the local file name where you want to save the downloaded file
    local_file_name = "Generate_Timesheet.py"

    try:
        # Send an HTTP GET request to the file URL
        response = requests.get(file_url)

        # Check if the request was successful (status code 200)
        if response.status_code == 200:
            # Save the file content to a local file
            with open(local_file_name, "wb") as file:
                file.write(response.content)
            print(f"File '{local_file_name}' downloaded successfully.")
            message = "The latest version of Generate_Timesheet.py is downloaded."\
                "\nClose this window and re-run the script."
            messagebox.showinfo("Success!", message)
            root.destroy()
            import sys
            sys.exit()
        else:
            print(f"Failed to download file. Status code: {response.status_code}")
            messagebox.showerror("Error!", 
                                 f"Failed to download file. Status code: {response.status_code}")

    except requests.exceptions.RequestException as e:
        print(f"An error occurred: {str(e)}")
        messagebox.showerror("Error!", f"An error occurred: {str(e)}")
    return

def check_for_update(current_version):
    url = "https://api.github.com/repos/Vyoma-Linguistic-Labs/ClickUp_Timesheets/releases/latest"
    response = requests.get(url)
    if response.status_code == 200:
        latest_version = response.json()["tag_name"]
        if latest_version != current_version:
            print(f"A new version ({latest_version}) is available. Please download it from GitHub.")
            root = tk.Tk()
            root.title("Newer Version available")
            
            message = f"A new version ({latest_version}) is available."\
                "\nClick on the button below to automatically download the"\
                    " \nlatest version from the GitHub repository."
            # link = "https://github.com/Vyoma-Linguistic-Labs/ClickUp_Timesheets/releases/latest"
            label = tk.Label(root, text=message, #fg="blue", cursor="hand2",
                             font=("Times", 12, "bold"))
            label.pack()        
            # # Bind the label to the open_link function with the corresponding link_url
            # label.bind("<Button-1>", lambda e, 
            #            url=link: webbrowser.open_new(url))
            
            # Submit Button
            submit_button = tk.Button(root, text=f"Download Latest Version {latest_version}",
                                      command=lambda arg=root: download_latest(arg))
            submit_button.pack(pady=10)
                        
            root.mainloop()
            import sys
            sys.exit()


def convert_milliseconds_to_hours_minutes(milliseconds):
    seconds = milliseconds / 1000
    minutes = seconds // 60
    hours = minutes // 60
    minutes = minutes % 60
    return (int(hours), int(minutes))

def memberInfo():
    url = "https://api.clickup.com/api/v2/team"
    headers = {"Authorization": "pk_3326657_EOM3G6Z3CKH2W61H8NOL5T7AGO9D7LNN"}
    response = requests.get(url, headers=headers)
    data = response.json()
    
    # Extract member id and username
    members_dict = {}
    for team in data['teams']:
        for member in team['members']:
            member_id = member['user']['id']
            member_username = member['user']['username']
            members_dict[member_id] = member_username
          
    # Exchange keys and values - keep last 4 digits corresponding to emp ID
    members_dict = {value[-4:]: key for key, value in members_dict.items()}
        
    return members_dict

def open_link(link):    
    webbrowser.open_new("app.clickup.com/t/"+link)
    
    
def get_selected_dates():
    start_date = start_cal.selection_get()
    end_date = end_cal.selection_get()
    key = key_entry.get().upper()

    # Format dates
    start_date_str = start_date.strftime("%b %d")
    end_date_str = end_date.strftime("%b %d")
    year_str = str(start_date.year)

    # Generate filename
    filename = f"{key}_{start_date_str}_to_{end_date_str}_{year_str}.xlsx"
    
    # Retrieve information from ClickUp
    start = time.time() 
    
    members_dict = memberInfo()
    employee_key = members_dict[key] # Convert our key to ClickUp key
       
    date_obj = datetime.strptime(str(start_date), '%Y-%m-%d')
    start_timestamp = int(date_obj.replace(tzinfo=timezone.utc).timestamp())
    date_obj = datetime.strptime(str(end_date), '%Y-%m-%d')
    end_timestamp = int(date_obj.replace(tzinfo=timezone.utc).timestamp())

    team_id = "3314662"
    url = "https://api.clickup.com/api/v2/team/" + team_id + "/time_entries"
    query = {
      "start_date": str(int(start_timestamp - 19800)*1000),#int(start_date*1000), # Converting to milliseconds from seconds
      "end_date": str(int((end_timestamp+86399)*1000) - 19800000),#int(end_date*1000),
      "assignee": employee_key,
    }
    
    headers = {
      "Content-Type": "application/json",
      "Authorization": "pk_3326657_EOM3G6Z3CKH2W61H8NOL5T7AGO9D7LNN"
    }
    
    response = requests.get(url, headers=headers, params=query)
    data = response.json()    
    
    # Initialize empty lists for each column
    task_names = []
    task_ids = []
    task_status = []
    # total_time = []
    durations = []
    dates = []
    days = []
    
    # Loop through the data and extract the required fields
    for entry in data['data']:
        try:
            task_names.append(entry['task']['name'])
            task_ids.append(entry['task']['id'])
            task_status.append(entry['task']['status']['status'])
        except:
            task_names.append('0')
            task_ids.append('0')
            task_status.append('0')
        durations.append(int(entry['duration']))
        start_time = int(entry['start']) // 1000 # Convert to seconds
        
        date = pd.Timestamp(start_time, unit='s').date()
        dates.append(date)

        # Convert start_time to a datetime object in UTC, Localize the datetime object to UTC                
        localized_start_datetime = pytz.utc.localize(datetime.utcfromtimestamp(start_time))
        # Convert the datetime object from UTC to IST
        day = localized_start_datetime.astimezone(ist_timezone).strftime('%A')
        days.append(day)
        
    # Create a pandas dataframe
    df = pd.DataFrame({
        'Task Name': task_names,
        'Task ID': task_ids,
        'Task Status': task_status,
        # 'Total Time Tracked': total_time,
        'Duration': durations,
        'Date': dates,
        'Day': days
    })
    
    # Create a new DataFrame with only unique Task IDs
    task_ids = df['Task ID'].unique()
    new_df = pd.DataFrame({'Task ID': task_ids})
    
    # Add columns for each day of the week
    days_of_week = ['Saturday', 'Sunday', 'Monday', 'Tuesday', 'Wednesday', 
                    'Thursday', 'Friday']
    for day in days_of_week:
        new_df[day] = 0
    
    # Loop through each task and add duration to the corresponding day column
    for task in task_ids:
        task_entries = df[df['Task ID'] == task]
        grouped_entries = task_entries.groupby(['Day']).sum(numeric_only=True)
        for day in days_of_week:
            if day in grouped_entries.index:
                new_df.loc[new_df['Task ID'] == task, day] = grouped_entries.loc[day]['Duration']
    
    # Merge the new DataFrame with the original DataFrame
    df_h = pd.merge(df, new_df, on='Task ID')
    
    # Drop duplicates
    df_h.drop_duplicates(subset='Task ID', inplace=True)
    
    # Convert the durations to hours format
    df_h[days_of_week] = df_h[days_of_week].apply(lambda x: x / 3600000).round(2)
    df_h = df_h.drop(['Duration', 'Date', 'Day'], axis=1)
    
    # define the API parameters
    headers = {"Authorization": "pk_3326657_EOM3G6Z3CKH2W61H8NOL5T7AGO9D7LNN"}
    
    # iterate over the unique task IDs in the dataframe
    for task_id in df_h['Task ID'].unique():
        # construct the API URL for the task ID
        url = "https://api.clickup.com/api/v2/task/" + task_id
    
        # make the API request and parse the JSON response
        response = requests.get(url, headers=headers)
        tasks = response.json()
        
        hrs_mins = convert_milliseconds_to_hours_minutes(tasks['time_spent'])
        df_h.loc[df_h['Task ID'] == task_id, 
                 'Total Time tracked for this task till now (hrs)'] = str(hrs_mins[0]) + 'h ' + str(hrs_mins[1]) + 'm'
        # If there is no Custom field just continue
        try:
            # iterate over the custom fields for the task
            for custom_field in tasks['custom_fields']:
                if 'value' in custom_field:
                    if custom_field['type'] == 'drop_down':
                        # set the value in the dataframe for the current task ID and custom field name
                        df_h.loc[df_h['Task ID'] == task_id, custom_field['name']] = custom_field['type_config']['options'][custom_field['value']]['name']
        except:
           pass        
    
    # Check if 'Goal Type' column exists
    if 'Goal Type' not in df_h.columns:
        # Add a new column with 'nan' values
        df_h['Goal Type'] = np.nan
    # Initialize a list to collect the names of rows that do not fit the criterion
    rows_with_missing_data = []
    row_id_with_missing_data = []
    rows_missing_goal_type = []
    row_id_missing_goal_type = []
    project_columns = list(set(df_h.columns.tolist()).intersection(columns_to_check))
    # Iterate through rows in the DataFrame
    for index, row in df_h.iterrows():
        # Extract the 'Task Name' column value for the current row
        task_name = row['Task Name']
        task_id = row['Task ID']
        # Check if all specified columns are NaN for the current row
        if row[project_columns].isna().all():
            # All specified columns are NaN, so add the 'Task Name' to the list
            rows_with_missing_data.append(task_name)
            row_id_with_missing_data.append(task_id)
            
        if pd.isna(row['Goal Type']):
            rows_missing_goal_type.append(task_name)
            row_id_missing_goal_type.append(task_id)
            
    # Output the names of rows that do not fit the criterion
    if rows_with_missing_data or rows_missing_goal_type:
        # Create a new top-level window for the error message
        error_window = tk.Toplevel(root)
        error_window.title("ERROR")
        # second_win = tkinter.Toplevel(root)
        root.eval(f'tk::PlaceWindow {str(error_window)} center')
        
        if rows_with_missing_data:
            # Add widgets to the error window to display the error message
            error_label = tk.Label(error_window, 
                                   text="‘Project/Product/Course/Website’ is not"\
                                       " set for the below task(s) (links provided)",
                                   font=("Times", 12, "bold"))
            error_label.pack()
                  
            # Create labels for each link
            link_label = []
            count = 0
            for link_text, link_url in zip(rows_with_missing_data, row_id_with_missing_data):
                            
                label = tk.Label(error_window, text=link_text, fg="blue", cursor="hand2")
                label.pack()        
                # Bind the label to the open_link function with the corresponding link_url
                label.bind("<Button-1>", lambda e, url=link_url: open_link(url))
                link_label.append(label)
                
                count += 1
        # You can add more widgets here to provide additional information
        if rows_missing_goal_type:
            # Add widgets to the error window to display the error message
            goal_error_label = tk.Label(error_window, 
                                   text="Goal Type not set for:",
                                   font=("Times", 12, "bold"))
            goal_error_label.pack()
                  
            # Create labels for each link
            goal_link_label = []
            count = 0
            for link_text, link_url in zip(rows_missing_goal_type, row_id_missing_goal_type):
                            
                label = tk.Label(error_window, text=link_text, fg="blue", cursor="hand2")
                label.pack()        
                # Bind the label to the open_link function with the corresponding link_url
                label.bind("<Button-1>", lambda e, url=link_url: open_link(url))
                goal_link_label.append(label)
                
                count += 1
        
        info_label = tk.Label(error_window, 
                               text="You can try generating your timesheet again once"\
                                   " you set the above information in these tasks.",
                               font=("Times", 12, "bold"))
        info_label.pack()
        # Start the mainloop for the error window
        error_window.mainloop()
            
    # Add the 'time_this_week' column by summing the values of all days_of_week columns
    df_h['Total Tracked this week in this task'] = df_h[days_of_week].sum(axis=1)
    # Calculate the totals of the days_of_week columns
    totals = df_h[days_of_week].sum(axis=0)

    # Append totals as a new row to the DataFrame
    df_h = pd.concat([df_h, totals.to_frame().T], ignore_index=True)
    # Sum the values in the last row of the DataFrame
    weekly_total = df_h.iloc[-1].sum()
    # Update the value in the 'Status' column for the last row
    df_h.at[df_h.index[-1], 'Task Status'] = 'Daily Totals ->'
    
    # Create an empty row with NaN values
    empty_row = pd.Series([np.nan] * len(df_h.columns), index=df_h.columns)
    # Append the empty row to the DataFrame
    df_h = pd.concat([df_h, empty_row.to_frame().T], ignore_index=True)    
    # Append a value to the 6th column
    df_h.iloc[-1, 5] = weekly_total
    if int((end_timestamp - start_timestamp)/86400)+1 <= 7:            
        df_h.iloc[-1, 3] = 'Week\'s total ='
        week_number = end_date.isocalendar()[1]
        df_h.at[df_h.index[-1], 'Task Name'] = f'Week #{week_number} - {start_date_str}, {year_str} - {end_date_str}, {year_str}'
    else:
        df_h.iloc[-1, 3] = 'Total Hours Tracked ='
        df_h.at[df_h.index[-1], 'Task Name'] = f'{start_date_str}, {year_str} - {end_date_str}, {year_str}'
    
    # Move the column to the 11th position
    df_h.insert(10, 'Total Tracked this week in this task', 
                df_h.pop('Total Tracked this week in this task'))
    # write the DataFrame to an Excel file
    df_h.to_excel(filename, index=False) #members_dict[key]+    
    # Display output filename
    if filename:
        output_label.config(text=f"Note: Please find the generated Excel output in this folder itself.\nFile: {filename}")
    total_time = df_h['Total Tracked this week in this task'].sum()
    print("Total Hours for this time frame: ", total_time)
    print(time.time()-start)    

    # Clear the selected dates and employee key
    start_cal.selection_clear()
    end_cal.selection_clear()
    # key_entry.delete(0, tk.END)
    if checkbox_var_1.get():
        # URLs to open
        url1 = 'https://docs.google.com/spreadsheets/d/1XLDSTT5m952eiOXhiUtxldIIoEAfQgiVKv5XY2HFOBg/edit?usp=sharing'
        # Open URLs in separate browser tabs
        # time.sleep(5)
        webbrowser.open_new_tab(url1)
    
    ### Draft mail Feature - for later version
    # if checkbox_var_2.get():
    #     time.sleep(5)            
    #     recipient = "venkat.s@vyomalabs.in"
    #     recipient_cc = "srilatha.vyoma@gmail.com"
    #     subject = f"Timesheet for Week #{week_number} - {start_date_str}, {year_str} - {end_date_str}, {year_str}"
    #     body = "Ram ram ram, \nPlease find my weekly timesheet in this Google tracker -"
    #     # Encode special characters in the subject and body
    #     subject = urllib.parse.quote(subject)
    #     body = urllib.parse.quote(body)
    #     # Compose the Gmail URL with recipient, subject, and body
    #     gmail_url = f"https://mail.google.com/mail/?view=cm&to={recipient}&cc={recipient_cc}&su={subject}&body={body}"    
    #     # Open the Gmail URL in a web browser
    #     webbrowser.open(gmail_url)
    # root.destroy()
if __name__ == "__main__":
    current_version = __version__  # Replace with your current version
    check_for_update(current_version)
    # Rest of your script
    
    root = tk.Tk()
    root.title("Timesheet Generator")
    
    # Increase the size of the window
    window_width = 900
    window_height = 750
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    x = (screen_width - window_width) // 2
    y = (screen_height - window_height) // 2
    root.geometry(f"{window_width}x{window_height}+{x}+{y}")
    # Change the background color of the window
    root.configure(background="lightblue")
    
    try:
        # URL of the image you want to display
        image_url = "https://digitalsanskritguru.com/wp-content/uploads/" \
                    "2020/05/Vyoma_Logo_Blue_500x243.png"
        # Fetch the image from the URL
        response = requests.get(image_url)
        image_data = response.content
        # Create a PIL Image object from the image data
        image = Image.open(BytesIO(image_data))
        # Resize the image to the desired size
        image = image.resize((167, 81), Image.LANCZOS)
        # Convert the PIL Image to a PhotoImage object
        photo = ImageTk.PhotoImage(image)
        # Create a label to display the image
        image_label = tk.Label(root, image=photo)
        image_label.photo = photo 
        # Position the label at the top right corner
        image_label.pack(anchor=tk.NE, padx=5, pady=5)
    except:
        pass
    
    # Start Date Frame
    start_frame = tk.Frame(root)
    start_frame.pack(anchor=tk.CENTER)
    
    # Start Date Label
    start_label = tk.Label(start_frame, text="Start Date:", bg="black", fg="white")
    start_label.pack(side="left")
    
    # Start Date Calendar
    start_cal = Calendar(start_frame, selectmode="day", date_pattern="yyyy-mm-dd")
    start_cal.pack(side="left")
    
    # End Date Frame
    end_frame = tk.Frame(root)
    end_frame.pack(pady=10)
    
    # End Date Label
    end_label = tk.Label(end_frame, text="End Date:", bg="black", fg="white")
    end_label.pack(side="left")
    
    # End Date Calendar
    end_cal = Calendar(end_frame, selectmode="day", date_pattern="yyyy-mm-dd")
    end_cal.pack(side="left")
    
    # Employee Key Label
    key_label = tk.Label(root, text="Employee ID: (Eg: C047)", bg="black", fg="white")
    key_label.pack(pady=10)
    
    # Employee Key Entry
    key_entry = tk.Entry(root)
    key_entry.pack()
    
    # Create a variable to store the checkbox state
    checkbox_var_1 = tk.IntVar()
    # Create the checkbox
    checkbox_1 = tk.Checkbutton(root, text="Open the Google sheet for TS Submission Status", 
                                variable=checkbox_var_1)
    checkbox_1.pack()
    
    # # Create a variable to store the checkbox state
    # checkbox_var_2 = tk.IntVar()
    # # Create the checkbox
    # checkbox_2 = tk.Checkbutton(root, text="Open draft mail to send TS", 
    #                             variable=checkbox_var_2)
    # checkbox_2.pack()
    
    # Output Label
    output_label = tk.Label(root, 
                            text="Please update the above tracker once you send the Timesheet mail.")
                            # font=font.Font(weight="bold"))
    output_label.pack()
    
    # Submit Button
    submit_button = tk.Button(root, text="Submit", command=get_selected_dates)
    submit_button.pack(pady=10)
    
    # Output Label
    output_label = tk.Label(root, 
                            text="Note: Please find the generated Excel output in this folder itself.")
                            # font=font.Font(weight="bold"))
    output_label.pack()
    
    # Create the footer label
    footer_label = tk.Label(root, text="Version " + __version__ +" (14th Sept 2023)", 
                            relief=tk.RAISED, anchor=tk.W)
    footer_label.pack(side=tk.BOTTOM, fill=tk.X)
    
    root.mainloop()
