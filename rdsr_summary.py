# -*- coding: utf-8 -*-
"""
Created on Mon Apr 21 13:26:47 2025

@author: nabih
"""

import pydicom
from pydicom.dataset import Dataset
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import plotly.express as px
import plotly.io as pio
import os
import pandas as pd
from typing import List, Dict, Any

pio.renderers.default = "browser"  

# =============================================================================
# Extract desired data
# =============================================================================

# List of CodeMeanings we want to extract
target_concepts = {"SOPInstanceUID",
                   "ContentDate", 
                   "ContentTime", 
                   "StationName", 
                   "PatientName", 
                   "PatientID", 
                   "PatientBirthDate", 
                   "PatientSex", 
                   "SoftwareVersions", 
                   "StudyInstanceUID", 
                   "SeriesInstanceUID", 
                   "SeriesNumber", 
                   "Person Observer Name", 
                   "Start of X-Ray Irradiation", 
                   "End of X-Ray Irradiation", 
                   "Total Number of Irradiation Events", 
                   "CT Dose Length Product Total", 
                   "Acquisition Protocol", 
                   "Irradiation Event UID", 
                   "Exposure Time", 
                   "Scanning Length", 
                   "Nominal Single Collimation Width", 
                   "Nominal Total Collimation Width", 
                   "Identification of the X-Ray Source", 
                   "KVP", 
                   "Maximum X-Ray Tube Current", 
                   "X-Ray Tube Current", 
                   "Exposure Time per Rotation", 
                   "Mean CTDIvol", 
                   "DLP"
                   }

def extract_targeted_data(ds: Dataset) -> Dict[str, Any]:
    """
    Extracts selected concept values from a DICOM Structured Report (SR).

    Recursively parses the ContentSequence of a DICOM SR object, collecting values 
    for specified target concepts. Handles common DICOM SR value types such as 
    MeasuredValueSequence, TextValue, DateTime, UID, and PersonName. Also checks 
    top-level attributes for matching concepts.

    Target concepts are defined externally in the `target_concepts` set.

    Args:
        ds (Dataset): A pydicom Dataset representing a DICOM Structured Report.

    Returns:
        Dict[str, Any]: A dictionary where keys are concept names and values are 
        the extracted data (e.g., numeric values, strings, dates).

    Raises:
        None

    Example:
        >>> extracted = extract_targeted_data(dicom_ds)
        >>> print(extracted)
        {'StationName': HALO, 'ContentDate': '20250424', ...}
    """
    
    # Initialise a dictionary to store extracted concept values
    data: Dict[str, Any] = {}

    def parse_sequence(sequence: List[Dataset]):
        """
        Recursively parses a list of DICOM datasets to extract targeted concept values.

        Args:
            sequence (List[Dataset]): A list of DICOM dataset items (ContentSequence).
        """
        
        for item in sequence:
            # Retrieve the Concept Name CodeSequence (0040,A043) if available
            concept_name_seq = item.get((0x0040, 0xA043), [])
            concept_name = None
            if concept_name_seq:
                concept_name = concept_name_seq[0].get("CodeMeaning", None)
                
            # If the concept name matches one of the targeted concepts, extract its value            
            if concept_name and concept_name in target_concepts:
                
                # Handle different types of value representations
                if "MeasuredValueSequence" in item:
                    measured = item.MeasuredValueSequence[0]
                    data[concept_name] = measured.get("NumericValue", None)

                elif "TextValue" in item:
                    data[concept_name] = item.TextValue

                elif "DateTime" in item:
                    data[concept_name] = item.DateTime

                elif "UID" in item:
                    data[concept_name] = item.UID

                elif "PersonName" in item:
                    data[concept_name] = item.PersonName

            # Recurse if there is nested ContentSequence
            if "ContentSequence" in item:
                parse_sequence(item.ContentSequence)

    # Parse if top-level ContentSequence exists
    if "ContentSequence" in ds:
        parse_sequence(ds.ContentSequence)

    # Check top-level attributes that may match the target concepts
    for tag in target_concepts:
        if tag in ds.dir():
            try:
                value = ds.data_element(tag).value
                data[tag] = str(value)  
            except Exception:
                pass

    return data


# =============================================================================
# Create the app
# =============================================================================
    
class DoseSummaryApp:
    """
    A GUI application for displaying, filtering, analysing, and exporting
    DICOM Radiation Dose Structured Report (RDSR) data.

    Main Features:
        - Load DICOM files from a selected folder
        - Extract and display targeted DICOM SR data in a scrollable table
        - Filter data dynamically by text fields and date ranges
        - Sort data by clicking table headers
        - Export full dataset to a CSV file
        - Plot histograms for numeric fields
        - View summary statistics for filtered data
        - Visualise exposure counts over time
        - Summarise multiple exposure events

    GUI Layout:
        - Tkinter-based interface
        - Filter panel for date and text filtering
        - Treeview table for data display with scrollbars
        - Bottom toolbar with buttons for export, plotting, and statistics
        - Dynamic dropdown for selecting histogram fields
        - Active filters displayed and manageable

    Attributes:
        root (tk.Tk): Main application window
        data (pd.DataFrame): Full extracted DICOM dataset
        filtered_data (pd.DataFrame): Currently filtered dataset
        sort_directions (dict): Tracks ascending/descending sort order per column
        active_filters (list): List of (column, value) tuples representing active text filters
        
        select_btn (tk.Button): Button to select folder containing DICOM files
        filter_frame (tk.Frame): Frame containing filtering widgets
        filter_display_frame (tk.Frame): Frame showing currently active filters
        filter_list_container (tk.Frame): Container frame to dynamically display filters

        start_date_entry (tk.Entry): Entry box for start date filter
        end_date_entry (tk.Entry): Entry box for end date filter
        dynamic_filter_column (tk.StringVar): Selected filter column.
        dynamic_filter_dropdown (ttk.Combobox): Dropdown to choose column for text filter
        dynamic_filter_value (tk.Entry): Entry for text value filter
        add_filter_button (tk.Button): Button to add a dynamic text filter
        clear_filters_btn (tk.Button): Button to clear all active filters

        tree (ttk.Treeview): Table widget displaying data
        
        histogram_column_var (tk.StringVar): Selected numeric column for histogram
        histogram_dropdown (ttk.Combobox): Dropdown for selecting numeric field for plotting
        plot_button (ttk.Button): Button to plot a histogram
        summary_button (ttk.Button): Button to show summary statistics popup
        plot_exposure_button (ttk.Button): Button to plot exposures over time
        multiple_exposures_btn (ttk.Button): Button to summarise multiple exposures

    Methods:
        __init__(self, root): Sets up the full GUI layout and binds all actions
        select_folder(self): Opens file dialog to select a folder and loads DICOM files
        load_dicom_data(self, folder_path): Loads and parses DICOM files from a given folder
        setup_table(self): Initialises the Treeview table for data display
        display_table(self, df): Populates the table with data from a DataFrame
        apply_date_range(self): Filters dataset by start and end dates
        add_dynamic_filter(self): Applies a dynamic text filter based on selected column and value
        clear_all_filters(self): Removes all active filters and refreshes display
        sort_column(self, col, reverse): Sorts the Treeview table by a selected column
        export_csv(self): Exports the full dataset to a CSV file
        plot_histogram(self): Plots a histogram for a selected numeric column
        show_summary_stats(self): Displays summary statistics (count, mean, median, etc.) for filtered data
        plot_exposures_over_time(self): Plots the number of exposures over time
        show_multiple_exposures_table(self): Displays a summary of multiple exposure events

    Dependencies:
        - pydicom
        - pandas
        - plotly
        - tkinter (built-in Python GUI toolkit)

    Notes:
        The class depends on an external `extract_targeted_data` function and
        a predefined set of `target_concepts` for extracting relevant DICOM information.
    """    
    
    def __init__(self, root):
        
        # Initialise main window
        self.root = root
        self.root.title("RDSR Summary")
    
        # Data storage
        self.data = pd.DataFrame()
        self.filtered_data = pd.DataFrame()
        self.sort_directions = {}  # Tracks sort direction per column
        self.active_filters = []   # List of (column, value) tuples
    
        # Folder Selection 
        self.select_btn = tk.Button(root, text="Select Folder", command=self.select_folder)
        self.select_btn.pack(pady=5)
    
        # --- Filter Section ---
        self.filter_frame = tk.Frame(root)
        self.filter_frame.pack(pady=5)
    
        # Date Range Filter
        tk.Label(self.filter_frame, text="Start Date (YYYYMMDD):").grid(row=1, column=0, padx=5)
        self.start_date_entry = tk.Entry(self.filter_frame)
        self.start_date_entry.grid(row=1, column=1, padx=5)
    
        tk.Label(self.filter_frame, text="End Date (YYYYMMDD):").grid(row=1, column=2, padx=5)
        self.end_date_entry = tk.Entry(self.filter_frame)
        self.end_date_entry.grid(row=1, column=3, padx=5)
    
        self.apply_dr_btn = tk.Button(self.filter_frame, text="Apply Date Range", command=self.apply_date_range)
        self.apply_dr_btn.grid(row=1, column=4, padx=5)
    
        # Dynamic Text Filter
        tk.Label(self.filter_frame, text="Filter Column:").grid(row=2, column=0, padx=5)
        self.dynamic_filter_column = tk.StringVar()
        self.dynamic_filter_dropdown = ttk.Combobox(
            self.filter_frame,
            textvariable=self.dynamic_filter_column,
            values=[
                "StationName", "PatientName", "PatientID", "PatientSex",
                "SoftwareVersions", "Person Observer Name",
                "Total Number of Irradiation Events", "Acquisition Protocol",
                "Target Region", "Identification of the X-Ray Source"
            ],
            state="readonly", width=30
        )
        self.dynamic_filter_dropdown.grid(row=2, column=1, padx=5)
    
        tk.Label(self.filter_frame, text="Filter Value:").grid(row=2, column=2, padx=5)
        self.dynamic_filter_value = tk.Entry(self.filter_frame)
        self.dynamic_filter_value.grid(row=2, column=3, padx=5)
    
        self.add_filter_button = tk.Button(self.filter_frame, text="Add Filter", command=self.add_dynamic_filter)
        self.add_filter_button.grid(row=2, column=4, padx=5)
    
        self.clear_filters_btn = tk.Button(self.filter_frame, text="Clear All Filters", command=self.clear_all_filters)
        self.clear_filters_btn.grid(row=2, column=5, padx=5)
    
        # Active Filters Display
        self.filter_display_frame = tk.Frame(self.root)
        self.filter_display_frame.pack(fill="x", padx=10, pady=5)
    
        self.filter_list_container = tk.Frame(self.filter_display_frame)
        self.filter_list_container.pack(side="left", fill="x", expand=True)
    
        # Data Table Setup
        self.setup_table()
    
        # Bottom Controls 
        bottom_frame = tk.Frame(root)
        bottom_frame.pack(pady=10)
    
        tk.Button(bottom_frame, text="Export to CSV", command=self.export_csv).pack(side=tk.LEFT, padx=10)
    
        self.histogram_column_var = tk.StringVar()
        self.histogram_dropdown = ttk.Combobox(bottom_frame, textvariable=self.histogram_column_var, state="readonly", width=30)
        self.histogram_dropdown.pack(side=tk.LEFT, padx=10)
    
        self.plot_button = ttk.Button(bottom_frame, text="Plot Histogram", command=self.plot_histogram)
        self.plot_button.pack(side=tk.LEFT, padx=10)
    
        self.summary_button = ttk.Button(bottom_frame, text="Show Summary Stats", command=self.show_summary_stats)
        self.summary_button.pack(side=tk.LEFT, padx=10)
    
        self.plot_exposure_button = ttk.Button(bottom_frame, text="Plot Exposures Over Time", command=self.plot_exposures_over_time)
        self.plot_exposure_button.pack(side=tk.LEFT, padx=10)
    
        self.multiple_exposures_btn = ttk.Button(bottom_frame, text="Multiple Exposure Summary", command=self.show_multiple_exposures_table)
        self.multiple_exposures_btn.pack(side=tk.LEFT, padx=10)

      
    def select_folder(self):
        """
        Open a folder selection dialog, load DICOM data from the selected folder,
        and display the extracted data in the table.
    
        If a folder is selected:
            - Loads and parses DICOM files into a DataFrame
            - Initialises filtered_data as a copy of the full dataset
            - Updates the table display with the filtered data
    
        Returns:
            None
        """
        
        folder_path = filedialog.askdirectory()
        if folder_path:
            self.data = self.load_dicom_data(folder_path)
            self.filtered_data = self.data.copy()
            self.display_table(self.filtered_data)


    def load_dicom_data(self, folder_path):
        """
        Recursively load DICOM `.dcm` files from a folder, extract targeted data,
        and organise the results into a structured DataFrame.
    
        For each DICOM file found:
            - Reads the file using pydicom
            - Extracts selected fields using `extract_targeted_data`
            - Appends the extracted data into a list of records
    
        The final DataFrame is reindexed to match a predefined column order.
    
        Parameters:
            folder_path (str): Path to the folder containing DICOM files.
    
        Returns:
            pd.DataFrame: DataFrame containing extracted DICOM data.
        """
        
        column_order = ["SOPInstanceUID",
                   "ContentDate", 
                   "ContentTime", 
                   "StationName", 
                   "PatientName", 
                   "PatientID", 
                   "PatientBirthDate", 
                   "PatientSex", 
                   "SoftwareVersions", 
                   "StudyInstanceUID", 
                   "SeriesInstanceUID", 
                   "SeriesNumber", 
                   "Person Observer Name", 
                   "Start of X-Ray Irradiation", 
                   "End of X-Ray Irradiation", 
                   "Total Number of Irradiation Events", 
                   "CT Dose Length Product Total", 
                   "Acquisition Protocol", 
                   "Irradiation Event UID", 
                   "Exposure Time", 
                   "Scanning Length", 
                   "Nominal Single Collimation Width", 
                   "Nominal Total Collimation Width", 
                   "Identification of the X-Ray Source", 
                   "KVP", 
                   "Maximum X-Ray Tube Current", 
                   "X-Ray Tube Current", 
                   "Exposure Time per Rotation", 
                   "Mean CTDIvol", 
                   "DLP"
                   ]
        
        records = []
        for root, _, files in os.walk(folder_path):
            for f in files:
                if f.endswith(".dcm"):
                    full_path = os.path.join(root, f)
                    read_dicom = pydicom.dcmread(full_path)
                    info = extract_targeted_data(read_dicom)
                    if info:
                        records.append(info)
        
        df = pd.DataFrame(records)
        
        # Ensure consistent column order
        df = df.reindex(columns=column_order, fill_value=None)
    
        return df


    def display_table(self, df):
        """
        Populate the Treeview table with data from the provided DataFrame.
    
        This method:
            - Clears any existing rows and columns
            - Sets up new column headers with sorting functionality
            - Inserts all rows into the Treeview
            - Updates the histogram dropdown to include available numeric columns
    
        Parameters:
            df (pd.DataFrame): The DataFrame containing the data to display.
    
        Returns:
            None
        """
        
        self.tree.delete(*self.tree.get_children())
        self.tree["columns"] = list(df.columns)
        self.tree["show"] = "headings"
        for col in df.columns:
            col_name = col.rstrip(" ↑↓")  # Clean if reused
            self.tree.heading(col_name, text = col_name, command=lambda _col=col_name: self.sort_column(_col, False))
            self.tree.column(col_name, width = 150, anchor = "w")
        for _, row in df.iterrows():
            self.tree.insert("", tk.END, values = list(row))
            
        # Update histogram dropdown with numeric columns
        numeric_cols = df.select_dtypes(include = ['number']).columns.tolist()
        self.histogram_dropdown['values'] = numeric_cols
        if numeric_cols:
            self.histogram_dropdown.current(0)


    def setup_table(self):
        """
        Create and configure the Treeview widget for displaying tabular data,
        along with horizontal and vertical scrollbars.
    
        This method:
            - Initialises a Frame to contain the table and scrollbars
            - Creates vertical and horizontal scrollbars
            - Creates a Treeview widget for data display
            - Links scrollbars to the Treeview for smooth navigation
    
        Returns:
            None
        """
        
        table_frame = tk.Frame(self.root)
        table_frame.pack(expand = True, fill = 'both')
    
        # Scrollbars
        self.tree_scroll_y = tk.Scrollbar(table_frame, orient = "vertical")
        self.tree_scroll_y.pack(side = "right", fill = "y")
    
        self.tree_scroll_x = tk.Scrollbar(table_frame, orient = "horizontal")
        self.tree_scroll_x.pack(side = "bottom", fill = "x")
    
        # Treeview
        self.tree = ttk.Treeview(table_frame, yscrollcommand = self.tree_scroll_y.set, xscrollcommand = self.tree_scroll_x.set)
        self.tree.pack(side = "left", fill = "both", expand = True)
    
        self.tree_scroll_y.config(command = self.tree.yview)
        self.tree_scroll_x.config(command = self.tree.xview)
        
        
    def refresh_filter_display(self):
        """
        Update the active filter display section in the GUI.
    
        This method:
            - Clears any existing filter display widgets
            - Adds visual indicators for active start date, end date, and dynamic filters
            - Provides "Remove" buttons to clear individual filters
            - Displays "None" if no filters are currently active
    
        Returns:
            None
        """
        
        # Clear previous filter display
        for widget in self.filter_list_container.winfo_children():
            widget.destroy()
    
        filters_exist = False
    
        def add_filter_button(text, clear_func):
            nonlocal filters_exist
            filters_exist = True
            frame = tk.Frame(self.filter_list_container)
            frame.pack(anchor="w", fill="x", pady=1)
            tk.Label(frame, text=text).pack(side="left")
            tk.Button(frame, text="Remove", command=clear_func, padx=3).pack(side="right")
    
        # Start Date
        start_date = self.start_date_entry.get().strip()
        if start_date:
            add_filter_button(f"Start Date ≥ {start_date}", self.clear_start_date_filter)
    
        # End Date
        end_date = self.end_date_entry.get().strip()
        if end_date:
            add_filter_button(f"End Date ≤ {end_date}", self.clear_end_date_filter)
    
        # Dynamic dropdown filters
        for idx, (col, val) in enumerate(self.active_filters):
            def remove(idx=idx):
                self.active_filters.pop(idx)
                self.apply_all_filters()
            add_filter_button(f"{col} contains '{val}'", remove)
    
        if not filters_exist:
            tk.Label(self.filter_list_container, text="None").pack(anchor="w")


    def apply_date_range(self):
        """
        Filter the dataset based on the specified start and end dates, 
        then update the displayed table and active filter list.
    
        This method:
            - Reads date inputs from the GUI entries
            - Converts the 'ContentDate' column to datetime format
            - Applies filtering for start and/or end dates if provided
            - Handles invalid date formats gracefully by showing error messages
            - Updates the filtered dataset for downstream operations (e.g. plotting)
            - Refreshes the table view and filter display
    
        Returns:
            None
        """
        
        df = self.data.copy()
    
        # Filter by Date Range
        start_date = self.start_date_entry.get().strip()
        end_date = self.end_date_entry.get().strip()
    
        if start_date or end_date:
            # Convert ContentDate column to datetime
            df["ContentDate"] = pd.to_datetime(df["ContentDate"], format="%Y%m%d", errors="coerce")
    
            if start_date:
                try:
                    start = pd.to_datetime(start_date, format="%Y%m%d")
                    df = df[df["ContentDate"] >= start]
                except ValueError:
                    messagebox.showerror("Date Error", "Invalid start date format. Use YYYYMMDD.")
            if end_date:
                try:
                    end = pd.to_datetime(end_date, format="%Y%m%d")
                    df = df[df["ContentDate"] <= end]
                except ValueError:
                    messagebox.showerror("Date Error", "Invalid end date format. Use YYYYMMDD.")
    
        # Store the filtered data so other methods (like plotting) can use it
        self.filtered_data = df
    
        # Update table
        self.display_table(df)
        self.refresh_filter_display()

        
    def add_dynamic_filter(self):
        """
        Add a new dynamic filter based on user input and reapply all filters.
    
        This method:
            - Retrieves the selected column and entered value from the GUI
            - Validates that both a column and a value are provided
            - Appends the (column, value) pair to the list of active filters
            - Reapplies all filters (including date filters) to refresh the displayed dataset
            - Displays a warning message if column or value input is missing
    
        Returns:
            None
        """
        
        column = self.dynamic_filter_column.get()
        value = self.dynamic_filter_value.get().strip()
    
        if not column or not value:
            messagebox.showwarning("Input Error", "Select a column and enter a value.")
            return
    
        self.active_filters.append((column, value))
        self.apply_all_filters()

        
    def apply_all_filters(self):
        """
        Apply all active filters (date range and dynamic column filters) to the dataset, 
        then update the table and filter display.
    
        This method:
            - Makes a fresh copy of the original dataset
            - Applies start and/or end date filters if provided
            - Converts the 'ContentDate' column to datetime for comparison
            - Applies user-defined dynamic filters based on partial text matches
            - Handles invalid date formats gracefully with error messages
            - Updates the `filtered_data` attribute for use by other functions
            - Refreshes the GUI table and filter indicators
    
        Returns:
            None
        """
        
        df = self.data.copy()
    
        # Date filter
        start_date = self.start_date_entry.get().strip()
        end_date = self.end_date_entry.get().strip()
        
        df["ContentDate"] = pd.to_datetime(df["ContentDate"], format="%Y%m%d", errors="coerce")
    
        if start_date:
            try:
                start = pd.to_datetime(start_date, format="%Y%m%d")
                df = df[df["ContentDate"] >= start]
            except ValueError:
                messagebox.showerror("Date Error", "Invalid start date format. Use YYYYMMDD.")
        if end_date:
            try:
                end = pd.to_datetime(end_date, format="%Y%m%d")
                df = df[df["ContentDate"] <= end]
            except ValueError:
                messagebox.showerror("Date Error", "Invalid end date format. Use YYYYMMDD.")
    
        # Dynamic filters
        for col, val in self.active_filters:
            if col in df.columns:
                df = df[df[col].astype(str).str.contains(val, case=False, na=False)]
    
        # Update display and internal filtered data
        self.filtered_data = df
        self.display_table(df)
        self.refresh_filter_display()


    def clear_all_filters(self):
        """
        Clear all active filters, reset filter fields, and reapply to refresh the dataset.
    
        This method:
            - Clears the list of dynamic filters
            - Resets the dynamic filter dropdown and input field
            - Clears both start and end date entries
            - Reapplies all filters to update the table to show unfiltered data
    
        Returns:
            None
        """
        
        self.active_filters.clear()
        self.dynamic_filter_column.set("")
        self.dynamic_filter_value.delete(0, tk.END)
        self.start_date_entry.delete(0, tk.END)
        self.end_date_entry.delete(0, tk.END)
        self.apply_all_filters()
        

    def remove_filter(self, column, value):
        """
        Remove a specific dynamic filter and reapply the remaining filters.
    
        Args:
            column (str): The name of the column for the filter to remove
            value (str): The filter value associated with the column
    
        This method:
            - Checks if the (column, value) pair exists in active filters
            - Removes the specified filter if found
            - Reapplies the remaining filters to the dataset
    
        Returns:
            None
        """
        
        if (column, value) in self.active_filters:
            self.active_filters.remove((column, value))
            self.apply_all_filters()

        
    def clear_start_date_filter(self):
        """
        Clear the start date filter and reapply all remaining filters.
    
        This method:
            - Deletes the user input from the start date entry field
            - Reapplies filters to update the displayed data accordingly
    
        Returns:
            None
        """
        
        self.start_date_entry.delete(0, tk.END)
        self.apply_all_filters()
    
    def clear_end_date_filter(self):
        """
        Clear the end date filter and reapply all remaining filters.
    
        This method:
            - Deletes the user input from the end date entry field
            - Reapplies filters to update the displayed data accordingly
    
        Returns:
            None
        """

        self.end_date_entry.delete(0, tk.END)
        self.apply_all_filters()
        
        
    def sort_column(self, col, reverse):
        """
        Sort the Treeview table by a specific column, toggling between ascending and descending order.
    
        Args:
            col (str): The name of the column to sort by
            reverse (bool): Whether to sort in descending order (True) or ascending (False)
    
        This method:
            - Gathers all rows from the Treeview
            - Attempts to sort numerically when possible; otherwise sorts alphabetically (case-insensitive)
            - Rearranges the rows based on the new sort order
            - Updates the column header to display an up (↑) or down (↓) arrow
            - Stores the new sort direction for future toggling
    
        Returns:
            None
        """
        
        # Get all items in the tree
        data = [(self.tree.set(k, col), k) for k in self.tree.get_children('')]
    
        # Try to convert values to float or leave as string
        def try_sort(val):
            try:
                return float(val)
            except (ValueError, TypeError):
                return val.lower() if isinstance(val, str) else val
    
        data.sort(key=lambda t: try_sort(t[0]), reverse=reverse)
    
        # Rearrange items in sorted order
        for index, (_, k) in enumerate(data):
            self.tree.move(k, '', index)
    
        # Update all column headers to remove arrows
        for c in self.tree["columns"]:
            base_text = c.rstrip(" ↑↓")  # remove any previous arrow
            self.tree.heading(c, text=base_text)

        # Set the arrow for this column
        arrow = "↓" if reverse else "↑"
        self.tree.heading(col, text=f"{col} {arrow}", command=lambda: self.sort_column(col, not reverse))
    
        # Update sort direction memory
        self.sort_directions[col] = not reverse


    def export_csv(self):
        """
        Export the currently displayed (filtered) Treeview data to a CSV file.
    
        This method:
            - Checks if there is data to export
            - Reconstructs a DataFrame from the visible Treeview entries
            - Opens a file dialog for the user to choose a save location
            - Saves the DataFrame as a CSV without an index column
            - Displays a success or warning message based on the outcome
    
        Returns:
            None
        """
        
        if not self.tree.get_children():
            messagebox.showwarning("No Data", "No filtered data to export.")
            return
    
        # Rebuild DataFrame from the Treeview
        columns = self.tree["columns"]
        rows = []
        for child_id in self.tree.get_children():
            row_values = self.tree.item(child_id)["values"]
            rows.append(row_values)
        filtered_df = pd.DataFrame(rows, columns=columns)
    
        file_path = filedialog.asksaveasfilename(defaultextension=".csv")
        if file_path:
            filtered_df.to_csv(file_path, index=False)
            messagebox.showinfo("Exported", f"Filtered data saved to {file_path}")

       
    def plot_histogram(self):
        """
        Plot an interactive histogram of a selected numeric column from the filtered data.
    
        This method:
            - Extracts the selected column's values from the Treeview (only currently visible data).
            - Skips non-numeric or invalid entries.
            - Uses Plotly to create and display a histogram.
            - Warns the user if no valid data exists or if an error occurs during plotting.
    
        Returns:
            None
        """
        
        selected_col = self.histogram_column_var.get()
        if not selected_col:
            messagebox.showwarning("Select Column", "Please select a column to plot.")
            return
    
        try:
            # Extract data from the Treeview (only visible/filtered data)
            values = []
            for item in self.tree.get_children():
                val = self.tree.set(item, selected_col)
                try:
                    values.append(float(val))
                except ValueError:
                    continue  # Skip non-numeric or empty cells
    
            if not values:
                messagebox.showwarning("No Data", f"No numeric data found in column '{selected_col}'.")
                return
    
            # Create a simple DataFrame for Plotly
            df = pd.DataFrame({selected_col: values})
    
            # Interactive histogram
            fig = px.histogram(df, x=selected_col, nbins=20,
                               title=f"Histogram of {selected_col}",
                               labels={selected_col: selected_col})
    
            fig.update_layout(bargap=0.1)
            fig.show()
    
        except Exception as e:
            messagebox.showerror("Error", f"Failed to plot histogram:\n{str(e)}")


    def show_summary_stats(self):
        
        if not self.tree.get_children():
            messagebox.showinfo("No Data", "No filtered data to summarize.")
            return
    
        # Get column headers from Treeview
        columns = self.tree["columns"]
    
        # Rebuild DataFrame from visible Treeview rows
        rows = []
        for child_id in self.tree.get_children():
            values = self.tree.item(child_id)["values"]
            cleaned_values = []
            for val in values:
                try:
                    cleaned_values.append(float(val))
                except (ValueError, TypeError):
                    cleaned_values.append(val)  # Keep as string if not numeric
            rows.append(cleaned_values)
    
        df = pd.DataFrame(rows, columns=columns)
    
        # Identify numeric columns only
        numeric_columns = []
        for col in df.columns:
            try:
                df[col] = pd.to_numeric(df[col], errors="coerce")
                if df[col].notna().sum() > 0:
                    numeric_columns.append(col)
            except Exception:
                continue
    
        numeric_data = df[numeric_columns]
    
        if numeric_data.empty:
            messagebox.showinfo("No Numeric Data", "No numeric columns available in filtered data.")
            return
        
        # List of variables to exclude from summary
        excluded_variables = {"Start of X-Ray Irradiation",
                              "End of X-Ray Irradiation",
                              "ContentTime",
                              "SeriesNumber",
                              "ContentDate", 
                              "PatientBirthDate"
                              }
        
        # Remove unwanted columns
        numeric_data = numeric_data.drop(columns=[col for col in excluded_variables if col in numeric_data.columns])

    
        # Summary statistics
        stats = numeric_data.describe().loc[["count", "mean", "50%", "min", "max", "std"]]
        stats.rename(index={"50%": "median"}, inplace=True)
        stats = stats.transpose()  # Variables as rows
    
        # Build popup window
        summary_window = tk.Toplevel(self.root)
        summary_window.title("Summary Statistics (Filtered Data Only)")
    
        tree = ttk.Treeview(summary_window, columns=["Variable"] + list(stats.columns), show="headings", height=len(stats))
    
        tree.heading("Variable", text="Variable")
        tree.column("Variable", anchor="w", width=200)
    
        for col in stats.columns:
            tree.heading(col, text=col.capitalize())
            tree.column(col, anchor="center", width=100)
    
        for idx, row in stats.iterrows():
            values = [idx] + [f"{v:.2f}" if pd.notnull(v) else "" for v in row]
            tree.insert("", "end", values=values)
    
        tree.pack(fill="both", expand=True)
    
        scrollbar = ttk.Scrollbar(summary_window, orient="horizontal", command=tree.xview)
        tree.configure(xscrollcommand=scrollbar.set)
        scrollbar.pack(fill="x")
        
    def plot_exposures_over_time(self):
        # Use filtered data if available
        if hasattr(self, 'filtered_data'):
            df = self.filtered_data.copy()
        else:
            df = self.data.copy()
    
        if df.empty:
            messagebox.showinfo("No Data", "No data available to plot.")
            return
    
        if "ContentDate" not in df.columns or "PatientID" not in df.columns:
            messagebox.showerror("Missing Columns", "Required columns 'ContentDate' or 'PatientID' not found.")
            return
    
        # Ensure datetime and cleanup
        df["ContentDate"] = pd.to_datetime(df["ContentDate"], errors="coerce")
        df["PatientID"] = df["PatientID"].astype(str).str.strip()
        df = df.dropna(subset=["ContentDate", "PatientID"])
        df = df[df["PatientID"] != ""]
    
        # Group by date and count exposures
        df["ContentDate"] = df["ContentDate"].dt.date
        exposure_counts = df.groupby("ContentDate").size().reset_index(name="Exposure Count")
    
        if exposure_counts.empty:
            messagebox.showinfo("No Exposure Data", "No exposure data found to plot.")
            return
    
        # Plot
        fig = px.bar(
            exposure_counts,
            x="ContentDate",
            y="Exposure Count",
            title="Exposure Count Over Time",
            labels={"ContentDate": "Date", "Exposure Count": "Number of Exposures"},
        )
        fig.update_layout(xaxis_tickformat="%Y-%m-%d")
        fig.show()
        
    def export_multi_exposures_to_csv(self):
        if not hasattr(self, "latest_multi_exposures_df") or self.latest_multi_exposures_df.empty:
            messagebox.showinfo("No Data", "There is no data to export.")
            return
    
        file_path = filedialog.asksaveasfilename(defaultextension=".csv",
                                                 filetypes=[("CSV files", "*.csv")],
                                                 title="Save Multiple Exposures Summary")
        if file_path:
            try:
                self.latest_multi_exposures_df.to_csv(file_path, index=False)
                messagebox.showinfo("Export Successful", f"Data exported to:\n{file_path}")
            except Exception as e:
                messagebox.showerror("Export Error", f"Failed to export CSV:\n{e}")


    def show_multiple_exposures_table(self):
        """
        Display a summary table of patients who had three or more exposures on the same day.
    
        This method:
            - Extracts the currently visible (filtered) data from the Treeview
            - Validates that required columns ('PatientID' and 'ContentDate') are present
            - Safely parses the 'ContentDate' column to datetime objects
            - Identifies patients with three or more exposures on a single day
            - Opens a new window listing grouped exposures by Patient ID and date
            - Creates individual mini-tables for each patient/date group
            - Adds vertical and horizontal scrollbars to handle large datasets
            - Provides an "Export to CSV" button to save the results externally
    
        Edge cases handled:
            - Missing columns
            - Malformed or missing date entries
            - No qualifying patients found
            - Empty or improperly formatted Treeview data
    
        Returns:
            None
        """
        
        if not self.tree.get_children():
            messagebox.showinfo("No Data", "No data is currently displayed.")
            return
    
        # Get data from currently displayed Treeview rows
        columns = self.tree["columns"]
        rows = [self.tree.item(i)["values"] for i in self.tree.get_children()]
        df = pd.DataFrame(rows, columns=columns)
    
        if "ContentDate" not in df.columns or "PatientID" not in df.columns:
            messagebox.showerror("Missing Columns", "Required columns (PatientID, ContentDate) are not present.")
            return
    
        df = df.copy()

        # Coerce ContentDate column to string first (safe for mixed formatting)
        df["ContentDate"] = df["ContentDate"].astype(str).str.strip()
        
        # Try parsing in two steps
        try:
            # If it's in full datetime format like "2019-08-06 00:00:00", this works
            df["ContentDate"] = pd.to_datetime(df["ContentDate"], errors="coerce")
            
            # Normalise to date only (removes time for grouping)
            df["ContentDate"] = df["ContentDate"].dt.normalize()
        
            # Drop invalid rows
            df = df.dropna(subset=["PatientID", "ContentDate"])
        
        except Exception as e:
            messagebox.showerror("Date Parsing Error", f"Could not parse ContentDate:\n{e}")
            return

        # Count exposures per PatientID and ContentDate
        exposure_counts = df.groupby(["PatientID", "ContentDate"]).size().reset_index(name="ExposureCount")
        multi_exposures = exposure_counts[exposure_counts["ExposureCount"] >= 3]
    
        if multi_exposures.empty:
            messagebox.showinfo("No Multiple Exposures", "No patients with ≥3 exposures on the same day.")
            return
    
        # Merge to get full exposure records
        filtered_df = df.merge(multi_exposures, on=["PatientID", "ContentDate"])
        filtered_df = filtered_df.drop(columns=["ExposureCount"])  # Drop the extra column before display
        # After filtering and dropping ExposureCount if needed
        self.latest_multi_exposures_df = filtered_df.copy()  # Store for export

        # Group by ContentDate and PatientID
        grouped = filtered_df.groupby(["ContentDate", "PatientID"])

        # Create new window
        exposure_window = tk.Toplevel(self.root)
        exposure_window.title("Multiple Exposures Summary")
        exposure_window.geometry("1200x600")
    
        # Canvas for scrolling
        canvas_frame = tk.Frame(exposure_window)
        canvas_frame.pack(fill="both", expand=True)
    
        canvas = tk.Canvas(canvas_frame)
        scrollbar_y = tk.Scrollbar(canvas_frame, orient="vertical", command=canvas.yview)
        scrollbar_x = tk.Scrollbar(exposure_window, orient="horizontal", command=canvas.xview)
    
        scrollable_frame = tk.Frame(canvas)
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        
        export_btn = tk.Button(exposure_window, text="Export to CSV", command=self.export_multi_exposures_to_csv)
        export_btn.pack(pady=(10, 0))

        def on_configure(event):
            canvas.configure(scrollregion=canvas.bbox("all"))
    
        scrollable_frame.bind("<Configure>", on_configure)
        canvas.configure(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)
    
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar_y.pack(side="right", fill="y")
        scrollbar_x.pack(side="bottom", fill="x")
    
        for (date, patient_id), group_df in grouped:
            # Header label
            header_text = f"{date.strftime('%d %b %Y')} — Patient ID: {patient_id} — {len(group_df)} Exposures"
            header = tk.Label(
                scrollable_frame,
                text=header_text,
                font=("Arial", 12, "bold"),
                bg="#e0e0e0",
                anchor="w"
                )
            header.pack(fill="x", padx=10, pady=(10, 0))
        
            # Frame for table and scrollbars
            table_frame = tk.Frame(scrollable_frame)
            table_frame.pack(fill="both", expand=True, padx=20, pady=5)
        
            # Scrollbars
            tree_scroll_y = tk.Scrollbar(table_frame, orient="vertical")
        
            # Treeview
            tree = ttk.Treeview(
                table_frame,
                columns=list(group_df.columns), show="headings", yscrollcommand=tree_scroll_y.set)
            
            tree_scroll_y.config(command=tree.yview)
        
            tree.grid(row=0, column=0, sticky="nsew")
            tree_scroll_y.grid(row=0, column=1, sticky="ns")
        
            table_frame.grid_rowconfigure(0, weight=1)
            table_frame.grid_columnconfigure(0, weight=1)
        
            tree.configure(height=8)
        
            # Insert headings and data
            for col in group_df.columns:
                tree.heading(col, text=col)
                tree.column(col, width=150, anchor="w")
        
            for _, row in group_df.iterrows():
                values = [row[col] if pd.notna(row[col]) else "" for col in group_df.columns]
                tree.insert("", "end", values=values)


# =============================================================================
# Run the app        
# =============================================================================

if __name__ == "__main__":
    root = tk.Tk()
    app = DoseSummaryApp(root)
    root.mainloop()