import tkinter as tk
from tkinter import ttk, filedialog, messagebox  
import psutil
import threading
import os
from PIL import Image, ImageTk
import win32gui # type: ignore
import win32ui # type: ignore
import win32con # type: ignore
import win32process # type: ignore
import csv


# Global variable to toggle sorting direction
sort_directions = {
    "Name": True,
    "Status": True,
    "CPU (%)": True,
    "Memory (MB)": True,
    "PID (Processes ID)": True,
    "Description": True  # Add this line
}

# Global variable for auto-refresh toggle and countdown
refresh_timer = None
countdown_time = 10  # 10 seconds for countdown

# Center Screen
def center_window(window, width, height):
    screen_width = window.winfo_screenwidth()
    screen_height = window.winfo_screenheight()
    x = (screen_width - width) // 2
    y = (screen_height - height) // 2
    window.geometry(f"{width}x{height}+{x}+{y}")

# Function to refresh process list every 10 seconds when checkbox is checked
def auto_refresh():
    global countdown_time

    if auto_refresh_enabled.get():
        countdown_label.config(text=f"Refreshing in {countdown_time} seconds...")
        countdown_time -= 1
        if countdown_time <= 0:
            update_process_list()  # Refresh process list
            countdown_time = 10  # Reset countdown to 10 seconds
        refresh_timer = root.after(1000, auto_refresh)  # Refresh every second
    else:
        countdown_label.config(text="Auto-refresh Disabled")

# Function to toggle auto-refresh state when checkbox is clicked
def toggle_auto_refresh():
    if auto_refresh_enabled.get():
        auto_refresh()  # Start auto-refresh if enabled
    else:
        if refresh_timer:
            root.after_cancel(refresh_timer)  # Cancel the countdown timer


# Function to get the executable path and icon for a process
def get_process_icon(pid):
    try:
        handle = win32process.OpenProcess(win32con.PROCESS_QUERY_INFORMATION | win32con.PROCESS_VM_READ, False, pid)
        exe_path = win32process.GetModuleFileNameEx(handle, 0)

        # Extract icon from the executable
        large, small = win32gui.ExtractIconEx(exe_path, 0)
        icon = large[0] if large else None
        return exe_path, icon
    except Exception:
        return None, None


# Function to update the process list
def update_process_list():
    # Clear the search box when updating the process list
    search_entry.delete(0, tk.END)
    #loading_label.pack(pady=10)  # Show loading spinner
    tree.delete(*tree.get_children())  # Clear the treeview

    # Modify the fetch_data function to collect PID and Description
    def fetch_data():
        process_list = []
        for process in psutil.process_iter(attrs=['pid', 'name', 'status', 'cpu_percent', 'memory_info']):
            try:
                pid = process.info['pid']
                name = process.info['name'] or "N/A"
                status = process.info['status']
                cpu_percent = process.info['cpu_percent']
                memory_mb = process.info['memory_info'].rss / (1024 * 1024)

                # Get executable path and icon
                exe_path, icon = get_process_icon(pid)
                process_list.append((pid, name, status, cpu_percent, memory_mb, exe_path, icon))
            except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                pass

        # Sort the list alphabetically by Name
        process_list.sort(key=lambda x: x[1])

        # Update the Treeview in the main thread
        tree.after(0, lambda: populate_tree(process_list))
        #tree.after(0, lambda: loading_label.pack_forget())  # Hide loading spinner

    threading.Thread(target=fetch_data, daemon=True).start()


# Modify the populate_tree function to include PID and Description
def populate_tree(process_list):
    global current_processes
    current_processes = process_list  # Store the full list for searching

    # Update the total count label
    total_count_label.config(text=f"Active Total: {len(process_list)}")

    for process in process_list:
        pid, name, status, cpu, memory, exe_path, icon = process

        # Determine the tag based on the status
        tag = None
        if status == 'stopped':
            tag = 'stopped'
        elif status == 'running':
            tag = 'running'

        # Create an icon image if available
        if icon:
            icon_image = Image.open(win32gui.GetIconInfo(icon)[4])
            icon_image = icon_image.resize((16, 16))  # Resize icon to fit in the Treeview
            icon_tk = ImageTk.PhotoImage(icon_image)
            icon_cache[process[0]] = icon_tk  # Cache icon to prevent garbage collection

            tree.insert('', 'end', values=(pid, name, status, f"{cpu:.1f}", f"{memory:.2f} MB", exe_path), image=icon_tk, tags=(tag,))
        else:
            tree.insert('', 'end', values=(pid, name, status, f"{cpu:.1f}", f"{memory:.2f} MB", exe_path), tags=(tag,))

    # Update the selected count
    update_selected_count()


# Filter the treeview based on the search query
def filter_treeview(event=None):
    query = search_entry.get().strip().lower()
    tree.delete(*tree.get_children())  # Clear the treeview

    # Check if the query contains multiple PIDs separated by commas
    if ',' in query:
        try:
            # Parse the query into a list of integers (PIDs)
            pid_list = [int(pid.strip()) for pid in query.split(',') if pid.strip().isdigit()]
            # Filter the processes based on PIDs
            filtered_list = [process for process in current_processes if process[0] in pid_list]
        except ValueError:
            messagebox.showerror("Invalid Input", "Please enter valid PIDs separated by commas.")
            return
    else:
        # Single search, check against PID or Name
        filtered_list = [
            process for process in current_processes 
            if query in process[1].lower() or query.isdigit() and int(query) == process[0]
        ]

    # Repopulate the treeview with the filtered list
    for process in filtered_list:
        pid, name, status, cpu, memory, exe_path, icon = process  # Unpack all 7 values

        if icon:
            icon_image = Image.open(win32gui.GetIconInfo(icon)[4])
            icon_image = icon_image.resize((16, 16))  # Resize icon to fit in the Treeview
            icon_tk = ImageTk.PhotoImage(icon_image)
            icon_cache[process[0]] = icon_tk  # Cache icon to prevent garbage collection

            tree.insert('', 'end', values=(pid, name, status, f"{cpu:.1f}", f"{memory:.2f} MB", exe_path), image=icon_tk)
        else:
            tree.insert('', 'end', values=(pid, name, status, f"{cpu:.1f}", f"{memory:.2f} MB", exe_path))

    # Update the selected count (if relevant)
    update_selected_count()


# Update the count of selected rows
def update_selected_count(event=None):
    selected_count = len(tree.selection())
    selected_count_label.config(text=f"Selected: {selected_count}")


# Sort the treeview by a column
def sort_column(col):
    # Sort the current_processes list based on the clicked column
    if col == "PID (Processes ID)":
        current_processes.sort(key=lambda x: x[0], reverse=not sort_directions[col])
    elif col == "Name":
        current_processes.sort(key=lambda x: x[1].lower(), reverse=not sort_directions[col])
    elif col == "Status":
        current_processes.sort(key=lambda x: x[2].lower(), reverse=not sort_directions[col])
    elif col == "CPU (%)":
        current_processes.sort(key=lambda x: float(x[3]), reverse=not sort_directions[col])
    elif col == "Memory (MB)":
        current_processes.sort(key=lambda x: float(x[4]), reverse=not sort_directions[col])
    elif col == "Description":
        current_processes.sort(
            key=lambda x: (x[5] or "").lower(),  # Handle None values
            reverse=not sort_directions[col]
        )
    
    # Reverse the sort direction for next click
    sort_directions[col] = not sort_directions[col]

    # Clear the treeview and repopulate with sorted data
    tree.delete(*tree.get_children())  # Clear the treeview
    
    for process in current_processes:
        pid, name, status, cpu, memory, exe_path, icon = process

        # Create an icon image if available
        if icon:
            icon_image = Image.open(win32gui.GetIconInfo(icon)[4])
            icon_image = icon_image.resize((16, 16))  # Resize icon to fit in the Treeview
            icon_tk = ImageTk.PhotoImage(icon_image)
            icon_cache[process[0]] = icon_tk  # Cache icon to prevent garbage collection

            tree.insert('', 'end', values=(pid, name, status, f"{cpu:.1f}", f"{memory:.2f} MB", exe_path), image=icon_tk)
        else:
            tree.insert('', 'end', values=(pid, name, status, f"{cpu:.1f}", f"{memory:.2f} MB", exe_path))

    update_selected_count()


# Function to copy selected row values to clipboard
def copy_selected_row():
    selected_items = tree.selection()  # Get selected items
    if selected_items:
        values = [tree.item(item, "values") for item in selected_items]  # Get values of the selected rows
        # Format values as strings and join with newlines for multiple selections
        copied_data = "\n".join([", ".join(value) for value in values])
        root.clipboard_clear()  # Clear clipboard
        root.clipboard_append(copied_data)  # Append data to clipboard
        root.update()  # Update clipboard to reflect changes
        print(f"Copied to clipboard:\n{copied_data}")  # Optional: Log the copied data for debugging
    else:
        print("No row selected.")  # Optional: Log when no row is selected

# Function to export selected rows to a CSV file with a dialog for file save
def export_selected_rows():
    selected_items = tree.selection()  # Get selected items
    if selected_items:
        # Open a file save dialog to choose the directory and enter file name
        file_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files", "*.csv")])
        
        if file_path:  # If a valid file path is selected
            # Open the file in write mode
            with open(file_path, mode="w", newline="") as file:
                writer = csv.writer(file)
                # Write the header row
                writer.writerow(columns)

                # Write data for each selected row
                for item in selected_items:
                    values = tree.item(item, "values")
                    writer.writerow(values)

            print(f"Selected rows exported to '{file_path}'.")
        else:
            print("Export canceled or invalid file path.")
    else:
        print("No row selected.")  # Optional: Log when no row is selected


# Function to select all rows in the treeview
def select_all_rows():
    tree.selection_set(tree.get_children())  # Select all rows currently visible in the Treeview
    update_selected_count()  # Update the selected count display


# Function to end a task
def end_task():
    selected_items = tree.selection()
    if selected_items:
        for item in selected_items:
            pid = int(tree.item(item, "values")[0])  # Get the PID from the selected row
            try:
                process = psutil.Process(pid)
                process.terminate()  # Attempt to terminate the process
                print(f"Terminated process PID {pid}")
            except Exception as e:
                print(f"Failed to terminate process PID {pid}: {e}")
        update_process_list()  # Refresh the process list
    else:
        print("No process selected.")

# Function to kill a task forcefully with confirmation for single row and clear search input
def kill_task():
    selected_items = tree.selection()  # Get selected items
    if selected_items:
        for item in selected_items:
            pid = int(tree.item(item, "values")[0])  # Get PID from the selected row
            try:
                process = psutil.Process(pid)
                
                # Confirm killing a single process
                if len(selected_items) == 1:
                    confirm = messagebox.askyesno(
                        "Confirm Kill",
                        f"Are you sure you want to kill the process with PID {pid}?",
                    )
                    if not confirm:
                        return  # If user cancels, do nothing

                # Forcefully kill the process
                process.kill()
                print(f"Killed process PID {pid}")
            except Exception as e:
                print(f"Failed to kill process PID {pid}: {e}")
        
        # After killing the process, clear the search input
        search_entry.delete(0, tk.END)
        
        update_process_list()  # Refresh the process list after killing
    else:
        messagebox.showwarning("No Selection", "No process selected. Please select a row to kill.")


# Function to show the context menu
def show_context_menu(event):
    # Get the row that was clicked
    item = tree.identify_row(event.y)
    if item:
        tree.selection_set(item)  # Select the row under the cursor
        context_menu.post(event.x_root, event.y_root)  # Display the context menu


# Function to kill all selected processes with confirmation
def kill_all_selected():
    selected_items = tree.selection()  # Get all selected rows
    
    if len(selected_items) > 1:  # Ensure multiple rows are selected
        confirm = messagebox.askyesno(
            "Confirm Kill",
            f"You are about to kill {len(selected_items)} processes. Are you sure?",
        )
        if confirm:
            for item in selected_items:
                pid = int(tree.item(item, "values")[0])  # Get PID from each selected row
                try:
                    process = psutil.Process(pid)
                    process.kill()  # Forcefully kill the process
                    print(f"Killed process PID {pid}")
                except Exception as e:
                    print(f"Failed to kill process PID {pid}: {e}")
            update_process_list()  # Refresh the process list after killing

            # After killing, clear the search input
            search_entry.delete(0, tk.END)
    elif len(selected_items) == 1:
        messagebox.showinfo("Action Restricted", "This button is for killing multiple processes. Use 'Kill PID' for single rows.")
    else:
        messagebox.showwarning("No Selection", "No rows selected. Please select multiple processes to use this button.")


def deselect_all_rows():
    tree.selection_remove(tree.selection())  # Remove all selections
    update_selected_count()  # Update count



# Fetch device name
device_name = os.environ['COMPUTERNAME']

# Set up the main UI window
root = tk.Tk()
root.title(f"Task Manager    -   {device_name}    [ PHEKDEY.PHORN - V.1.0]")  # Use device name in title
window_width = 900
window_height = 600

# Center the window on the screen
center_window(root, window_width, window_height)

# Initialize the auto-refresh toggle variable after the root window is created
auto_refresh_enabled = tk.BooleanVar(value=False)

# Add a search box
search_frame = tk.Frame(root)
search_frame.pack(fill=tk.X, padx=10, pady=5)

search_label = tk.Label(search_frame, text="Search here:")
search_label.pack(side=tk.LEFT, padx=5)

search_entry = tk.Entry(search_frame)
search_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
search_entry.bind("<KeyRelease>", filter_treeview)

# Add a loading spinner (text-based)
loading_label = tk.Label(root, text="Loading...", font=("Arial", 10), fg="blue")

# Frame to hold the Treeview and scrollbar
frame = tk.Frame(root)
frame.pack(fill=tk.BOTH, expand=True)

# Create a Treeview widget
columns = ("PID (Processes ID)", "Name", "Status", "CPU (%)", "Memory (MB)", "Description")
tree = ttk.Treeview(frame, columns=columns, show='headings', height=20, selectmode="extended")

# Define column headings and add sorting functionality
for col in columns:
    tree.heading(col, text=col, command=lambda _col=col: sort_column(_col))
    tree.column(col, anchor=tk.W, width=180 if col == "Name" else 100)

# Add vertical scrollbar
scrollbar = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=tree.yview)
tree.configure(yscroll=scrollbar.set)
scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

# Add the Treeview to the frame
tree.pack(fill=tk.BOTH, expand=True)

# Configure row styles based on status
tree.tag_configure('stopped', foreground='dark red')  # Dark red text for 'stopped'
tree.tag_configure('running', foreground='black')     # Default text color for 'running'

# Bottom frame for additional info
bottom_frame = tk.Frame(root)
bottom_frame.pack(fill=tk.X)

# Add a label to display the count of total processes
total_count_label = tk.Label(bottom_frame, text="Active Total: 0", font=("Arial", 10))
total_count_label.pack(side=tk.LEFT, padx=10)

# Add a label to display the count of selected rows
selected_count_label = tk.Label(bottom_frame, text="Selected: 0", font=("Arial", 10))
selected_count_label.pack(side=tk.LEFT, padx=10)

# Create an "Update" button
update_button = tk.Button(bottom_frame, text="Update", command=update_process_list)
update_button.pack(side=tk.RIGHT, padx=10, pady=10)

# Add a "Kill All Selected" button to the search frame (next to Select All)
kill_all_button = tk.Button(search_frame, text="Kill All Selected", command=kill_all_selected)
kill_all_button.pack(side=tk.RIGHT, padx=5)  # Place it next to the "Select All" button

# Add a "Select All" button to the search frame (top-right)
select_all_button = tk.Button(search_frame, text="Select All", command=select_all_rows)
select_all_button.pack(side=tk.RIGHT, padx=5)  # Place it on the right side of the search frame

deselect_all_button = tk.Button(search_frame, text="Deselect All", command=deselect_all_rows)
deselect_all_button.pack(side=tk.RIGHT, padx=5)

# Add a "Copy" button
copy_button = tk.Button(bottom_frame, text="Copy", command=copy_selected_row)
copy_button.pack(side=tk.RIGHT, padx=10, pady=10)

# Add an "Export" button
export_button = tk.Button(bottom_frame, text="Export", command=export_selected_rows)
export_button.pack(side=tk.RIGHT, padx=10, pady=10)

# Add checkbox for auto-refresh
auto_refresh_checkbox = tk.Checkbutton(bottom_frame, text="Enable Auto-Refresh (10s)", variable=auto_refresh_enabled, command=toggle_auto_refresh)
auto_refresh_checkbox.pack(side=tk.LEFT, padx=10)

# Label to show the countdown or auto-refresh status
countdown_label = tk.Label(bottom_frame, text="Auto-refresh Disabled", font=("Arial", 10))
countdown_label.pack(side=tk.LEFT, padx=10)

# Create a context menu
context_menu = tk.Menu(root, tearoff=0)
#context_menu.add_command(label="End Task", command=end_task)
context_menu.add_command(label="Kill PID", command=kill_task)

# Bind right-click to the Treeview
tree.bind("<Button-3>", show_context_menu)  # Right-click

root.bind("<Control-a>", lambda event: select_all_rows())
root.bind("<Control-d>", lambda event: deselect_all_rows())

# Cache for icons to prevent garbage collection
icon_cache = {}

# Store current processes for searching
current_processes = []

# Bind Treeview selection event to update the selected count
tree.bind("<<TreeviewSelect>>", update_selected_count)

# Fetch initial process list
update_process_list()

# Run the UI application
root.mainloop()
