from itertools import count
import time
import win32com.client
import json
import pythoncom

def attach_powerpoint():
    """
    Attach to an existing PowerPoint instance or start a new one.
    
    Parameters:
        None

    Returns:
        ppt_app (COM Object): PowerPoint Application object.
    """
    try:
        # Try to get an active PowerPoint instance
        ppt_app = win32com.client.GetActiveObject("PowerPoint.Application")
        print("Attached to existing PowerPoint instance")
    except:
        # If no active instance, launch a new one
        ppt_app = win32com.client.Dispatch("PowerPoint.Application")
        print("Launched new PowerPoint instance")
    ppt_app.Visible = True
    return ppt_app

def attach_excel():
    """
    Attach to an existing Excel instance or start a new one.

    Parameters:
        None

    Returns:
        excel (COM Object): Excel Application object.
    """
    # Launch a new Excel instance
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = True
    return excel

def wait_for_ppt_ready(ppt, timeout=10):
    start = time.time()
    while True:
        try:
            _ = ppt.Slides.Count
            return
        except Exception:
            if time.time() - start > timeout:
                raise TimeoutError("PowerPoint slides not ready")
            pythoncom.PumpWaitingMessages()
            time.sleep(0.3)

def refresh_ppt_objects(ppt, excel, config):
    """
    Refresh all linked objects and charts in a PowerPoint presentation.

    Parameters:
        ppt (COM Object): The PowerPoint Presentation object.

    Returns:
        None
    """
    
    wb = excel.Workbooks.Open(config["excel_path"])
    # ws = wb.Sheets("Overall Metrics")
    # print(f"\nCell vaue is : {ws.Range("E1").Value}\n")
    # print("Bye")

    print("\nRefreshing charts and linked Excel objects...")
    # for slide in ppt.Slides:
    slides = ppt.Slides
    count = slides.Count

    for i in range(1, count + 1):
        slide = slides(i)
        print(f"  Slide {slide.SlideIndex}...")
        # for shape in slide.Shapes:
        shapes = slide.Shapes
        shape_count = shapes.Count
        for j in range(1, shape_count + 1):
            shape = shapes(j)
            try:
                # Update linked objects
                if hasattr(shape, "LinkFormat") and shape.LinkFormat is not None:
                    shape.LinkFormat.Update()
                    print(f"Updated linked object: {shape.Name}")
                # Refresh charts
                if hasattr(shape, "HasChart") and shape.HasChart:
                    shape.Chart.Refresh()
                    print(f"Chart refreshed: {shape.Name}")
            except:
                print(f"Skipping shape '{shape.Name}'")

def update_ppt_tables(ppt, excel, config):
    """
    Update PowerPoint tables with data from Excel based on config.

    Parameters:
        ppt (COM Object): The PowerPoint Presentation object.
        excel (COM Object): The Excel Application object.
        config (dict): Dictionary loaded from the JSON config file containing paths and table mappings.

    Returns:
        None
    """
    # Open the Excel workbook
    wb = excel.Workbooks.Open(config["excel_path"])
    print("\nUpdating tables from Excel...")
    
    # Iterate through slides and tables defined in the config
    for slide_num, slide_config in config["slides"].items():
        slide = ppt.Slides(int(slide_num))
        print(f"Slide {slide_num}...")
        
        for table_name, tbl_config in slide_config.get("tables", {}).items():
            for shape in slide.Shapes:
                # Find the table shape by name
                if shape.HasTable and shape.Name.strip().lower() == table_name.strip().lower():
                    print(f"Updating table: {shape.Name}")
                    table = shape.Table
                    # Iterate through the rows and columns defined in the config
                    for r, row in enumerate(range(tbl_config["ppt_rows"][0], tbl_config["ppt_rows"][1] + 1)):
                        for c, col in enumerate(range(tbl_config["ppt_cols"][0], tbl_config["ppt_cols"][1] + 1)):
                            # Get the corresponding Excel cell value
                            excel_row = tbl_config["excel_rows"][0] + r
                            excel_col = tbl_config["excel_cols"][0] + c
                            cell = wb.Sheets(tbl_config["sheet"]).Cells(excel_row, excel_col)
                            value = cell.Text
                            # table.Cell(row, col).Shape.TextFrame.TextRange.Text = value
                            # Get font properties from Excel (for example, font size and color)
                            font = cell.Font
                            # font_size = font.Size
                            # font_color = font.Color
                            font_color = cell.DisplayFormat.Font.Color


                            # Update PowerPoint cell
                            ppt_cell = table.Cell(row, col).Shape.TextFrame.TextRange
                            ppt_cell.Text = value

                            # Set the font size and color in PowerPoint
                            # ppt_cell.Font.Size = font_size
                            ppt_cell.Font.Color.RGB = font_color
    
    wb.Save()
    wb.Close(SaveChanges=0)

def save_and_close(ppt, ppt_app, excel, config):
    """
    Save and close PowerPoint and Excel instances.

    Parameters:
        ppt (COM Object): The PowerPoint Presentation object.
        ppt_app (COM Object): The PowerPoint Application object.
        excel (COM Object): The Excel Application object.
        config (dict): Dictionary loaded from the JSON config file, includes output path.

    Returns:
        None
    """
    try:
        print("\nSaving updated PowerPoint...")
        if ppt:
            # Save the PowerPoint presentation
            if ppt.FullName.lower() != config["ppt_output_path"].lower():
                ppt.SaveAs(config["ppt_output_path"])
            else:
                ppt.Save()
            print("PowerPoint saved successfully!")
        else:
            print("Warning: PowerPoint file was not opened properly.")

    except Exception as save_error:
        print(f"Failed to save PowerPoint: {save_error}")

    finally:
        print("Closing PowerPoint and Excel...")
        if ppt:
            try:
                ppt.Close()
            except Exception as close_error:
                print(f"Warning: Could not close PowerPoint: {close_error}")

        # Quit PowerPoint and Excel
        ppt_app.Quit()
        excel.Quit()
        print("Excel to PPT Automation complete.")

def reset_excel_cell(excel_path, sheet_name, cell_address, value):
    """
    Reset a specific cell in an Excel workbook to a given value.

    Parameters:
        excel (COM Object): The Excel Application object.
        excel_path (str): Path to the Excel workbook.
        sheet_name (str): Name of the sheet containing the cell.
        cell_address (str): Address of the cell to reset (e.g., "E1").
        value: Value to set in the specified cell.

    Returns:
        None
    """
    excel = attach_excel()
    wb = excel.Workbooks.Open(excel_path)
    ws = wb.Sheets(sheet_name)
    ws.Range(cell_address).Value = value
    wb.Save()
    if 'wb' in locals() and wb:
        wb.Close(SaveChanges=False)
    if excel:
        excel.Quit()

def update_ppt_from_excel(config_path):
    """
    Main function to update PowerPoint from Excel using the given config file.

    Parameters:
        config_path (str): Path to the JSON config file.

    Returns:
        None
    """
    # Load the configuration from the JSON file
    with open(config_path, "r") as f:
        config = json.load(f)
    
    # Attach to PowerPoint and Excel instances
    ppt_app = attach_powerpoint()
    excel = attach_excel()

    wb = excel.Workbooks.Open(config["excel_path"])
    ws = wb.Sheets("Overall Metrics")
    ws.Range("E1").Value = 45730
    wb.Save()
    if 'wb' in locals() and wb:
        wb.Close(SaveChanges=False)
    if excel:
        excel.Quit()

    excel = attach_excel()

    # Open the PowerPoint presentation
    ppt = ppt_app.Presentations.Open(config["ppt_path"], WithWindow=True)
    print("\nPowerPoint is being processed...\n")
    wait_for_ppt_ready(ppt)
    
    if not ppt:
        raise Exception("Failed to open PowerPoint file!")
    
    # Refresh linked objects and update tables
    refresh_ppt_objects(ppt, excel, config)
    update_ppt_tables(ppt, excel, config)
    # Save and close the instances
    save_and_close(ppt, ppt_app, excel, config)
