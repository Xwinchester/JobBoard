# JobBoard

Machine Shop Job Board is a Python GUI application developed using the Tkinter library. The purpose of this application is to serve as a job board for machine shops, providing a visual representation of job information, statuses, and relevant details.

Overview
The key features of the Machine Shop Job Board include:

Graphical User Interface (GUI):

Utilizes Tkinter for a user-friendly full-screen window.
Features a dedicated frame housing a Treeview widget for displaying job information.
Configuration File Management:

Implements ConfigParser for efficient handling of configuration settings.
The config.ini file stores details such as Excel file paths, colors, font size, and update intervals.
File Operations:

Allows users to open Excel files through a file dialog.
Reads specified Excel sheets to gather essential job data.
Tabular Display:

Utilizes the Treeview widget to present job information in a clear tabular format.
Implements color coding for different job statuses, enhancing visibility.
Dynamic Table and Font Size:

Adjusts font size and table height dynamically based on screen size and content.
Automated Job Board Updates:

Regularly updates the job board to reflect changes in the linked Excel file.
Menu Bar:

Provides options to open new Excel files, exit, access color settings, and more.
Color Picker:

Enables users to customize colors for different job statuses.
Excel and Configuration Menu:

Offers convenient options to open the active Excel sheet and access configuration settings.
Purpose
The Machine Shop Job Board offers a simple way to show machinist's the current state of each machine and where they will be working as well as what jobs they will be working on.
