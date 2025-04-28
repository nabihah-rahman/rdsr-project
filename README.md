# RDSR Project

A Tkinter-based interactive tool for filtering, analysing, visualising and exporting DICOM radiation dose structured reports (RDSRs) in a readable format.  

## Description

DoseSummaryApp is an interactive Python application that allows users to work with DICOM RDSRs to extract data about imaging exposures. It offers features like dynamic filtering, summary statistics, detection of multiple exposures for the same patient, histogram plotting, and exporting filtered results — all inside a Tkinter interface. Designed for quick review of exposures and QA workflows, where efficient filtering and visualisation of patient imaging exposures are important.

## Getting Started

### Dependencies

* Python 3.7 or higher 
* tkinter (comes with Python standard library)
* pandas
* plotly
* pydicom
* Works on Windows, macOS, and Linux

Install the necessary libraries: 
```
pip install -r requirements.txt
```
### Installing

* Clone or download this repository to your local machine. 
* No additional setup required if dependencies are installed.
* Ensure your RDSR files have PatientID and ContentDate data. 

### Executing program

* Navigate to the project folder. 
* Run the main application file (rdsr_summary.py) 
```
python rdsr_summary.py
```
In the app: 
* Select a folder containing DICOM RDSR files.
* Apply filters, view summary statistics, plot histograms, detect multiple exposures, or export results. 
  
## Help

* Ensure your ContentDate field is properly formatted (YYYYMMDD or a standard datetime string).
* Non-numeric fields will be automatically handled or skipped when plotting histograms.

```
python rdsr_summary.py --help
```

## Authors

Nabihah Rahman
[@nabihah-rahman](https://github.com/nabihah-rahman)

## Version History

* 0.2
    * Updated documentation
    * See [commit change]() or See [release history]()
* 0.1
    * Initial release

## License

This project is licensed under the MIT License - see the LICENSE.md file for details
