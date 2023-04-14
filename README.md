# WebMD-TaskWeights

This is the macro I developed to automatically populate cells in a client's Workfront data spreadsheet. The code enables the user to select a READ column and a WRITE column while including failsafes to prevent overwriting existing data.

Using a key provided by the department, the macro maps each string value in the "TaskNameColumn" to a specific value for up to five additional columns per row. For instance, if "Newsletter Team Dev" is present in the "read" column, the macro will populate the value "2" in a newly added column named "NL/WS" and "1" in a newly added column named "CES." Each of the strings maps to between one and five new fields.

To avoid any mishaps, if the macro detects text already present in a designated WRITE column, an error message promptly pops up, and the code refrains from running until the user adjusts the column settings appropriately. Additionally, error messages notify the user if they have mistakenly indicated the READ column.

I have created documentation for the primary users of this macro, outlining how to troubleshoot the macro and adjust its settings as necessary. All modifications are performed within the block of code that I have outlined with asterisks at the start of the script.

Previously, data input for spreadsheets with over 1,800 line items was done manually. However, with the use of this macro, the time taken to complete the task is significantly reduced.
