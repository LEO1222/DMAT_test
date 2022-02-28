#!/usr/bin/env python3.7
import mhi.enerplot

# Launch or Connect to the Enerplot application
with mhi.enerplot.application() as enerplot:
    enerplot.silence = True


    # Workspace & Book names
    folder = r"~\Documents"
    workspace_name = "ScriptDemo"
    book_name = "ScriptDemo"
    # Load the workspace if it exists, otherwise create a new one.
    try:
        enerplot.load_workspace(workspace_name, folder=folder)
        # Get reference to book, if exists, otherwise create a new one.
        try:
            book = enerplot.book(book_name)
        except ValueError:
            book = enerplot.new_book(book_name, folder=folder)
    except FileNotFoundError:
        enerplot.new_workspace()
        book = enerplot.book("Untitled")
    
    # Locate (or load) "xxxx" datafile
    # enerplot.load_datafiles('C:\\Users\\Public\\Documents\\Enerplot\\1.0.0\\Examples\\DataFiles\\CSV_Files\\xxxx.csv', load_data=True)
    try:
        data_set = enerplot.datafile("xxxxx.csv")
    except ValueError as e:
        data_set = enerplot.load_datafiles("DataFiles\\CSV_Files\\xxxx.csv",folder=enerplot.examples)[0]

    # Get references to Rectifier AC phase voltages
    ph_a = data_set["Rectifier\\AC Voltage:1"]
    ph_b = data_set["Rectifier\\AC Voltage:2"]
    ph_c = data_set["Rectifier\\AC Voltage:3"]

    # Locate (or create) a sheet called "Sheet1"
    try:
        sheet1 = book.sheet("Sheet1")
    except ValueError:
        sheet1 = book.new_sheet("Sheet1")
    
        
    # Locate (or create) a Graph Frame with "Rectifier AC Voltage" title
    frame = sheet1.find("GraphFrame", title="Rectifier AC Voltage")
    if not frame:
        frame = sheet1.graph_frame(1, 1, 45, 32,title="Rectifier AC Voltage", xtitle="Time")
        top = frame.panel(0)
        top.properties(title="Phase Voltages (kV)")
        top.add_curves(ph_a, ph_b, ph_c)
        top.zoom(xmin=0, xmax=0.2, ymax=1.2, ymin=-0.2)
    
    # Locate (or create) the Graph with title "Zero Sequence Voltage (V)"
    bottom = frame.find(title="Zero Sequence Voltage (V)")
    if not bottom:
        bottom = frame.add_overlay_graph()
        bottom.properties(title="Zero Sequence Voltage (V)")
    
    # Remove any curves accidentally saved in "bottom" graph
    curves = bottom.list()
    if curves:
        for curve in curves:
            curve.cut()
    # Calculate and add the Zero Sequence channel to bottom
    v_zero = (ph_a + ph_b + ph_c) * 1000
    zero_seq = data_set.set_channel(v_zero, "Zero Sequence", "Script Output")
    bottom.add_curves(zero_seq)
    
    # Save the book and workspace
    book.save_as(book_name, folder=folder)
    enerplot.save_workspace_as(workspace_name, folder=folder)


