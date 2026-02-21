import shutil
import os
import time
"""
Main automation controller script.

This script provides an interactive interface to perform one or more of the following tasks:
1. Download spend report files from Outlook email attachments.
2. Transfer values and perform pivoting in Excel based on the downloaded files.
3. Update a PowerPoint presentation using the Excel data.
4. Combine all of the above in sequence.

The script uses external modules:
- `email_automation` for Outlook email automation
- `excel_automation` for Excel transformation
- `ppt_automation` for PowerPoint updates
- `insight_automation` for pasting generated insights onto ppt
"""

# Import required functions from custom modules
from excel_automation import automate_excel_pivoting
from ppt_automation import update_ppt_from_excel
from email_automation import download_attachments
from insights_automation import generate_and_paste_insights
from ppt_automation import reset_excel_cell

# Define paths to configuration files
excel_config_path = "C:\\Users\\askpr\\Downloads\\PPT_Automation\\config\\excel_config.json"
ppt_config_path = "C:\\Users\\askpr\\Downloads\\PPT_Automation\\config\\ppt_config_mod.json"
original_excel_path = "C:\\Users\\askpr\\Downloads\\PPT_Automation\\files\\SourceExcel.xlsx"
original_ppt_path = "C:\\Users\\askpr\\Downloads\\PPT_Automation\\files\\DeckTemplate.pptx"
new_ppt_path = "C:\\Users\\askpr\\Downloads\\PPT_Automation\\files\\test.pptx"

if __name__ == "__main__":

    # Prompt user for the task to be performed
    # key = input(
    #     'What task is to be performed? \n'
    #     'Download Spend files from email (d) \n'
    #     'Transfer the values from spend file to collated Excel (e) \n'
    #     'Update PPT (p) \n'
    #     'Transfer values and update PPT (b) \n'
    #     'Download files, transfer values, and update PPT: all together (c)\n'
    #     'Choose your option: '
    # )

    key = input(
        'What task is to be performed? \n'
        'Download Spend files from email (d) \n'
        'Transfer the values from spend file to collated Excel (e) \n'
        'Update PPT from Excel (u) \n'
        'Generate Insights in PPT (i) \n'
        'Update PPT from Excel and Generate Insights in PPT: both together (c)\n'
        'Choose your option: '
    )
    
    # # Option to only download email attachments
    # if key.lower()=='d':
    #     download_attachments(excel_config_path)
    #     print('✅ Files have been downloaded from Outlook.')
    # # Option to only run Excel automation (e.g., pivoting or merging)
    # elif key.lower()=='e':
    #     automate_excel_pivoting(excel_config_path)
    #     print("✅ Excel pivoting complete!")
    # # Option to only update the PowerPoint from the Excel data
    # elif key.lower() == 'p':
    #     update_ppt_from_excel(ppt_config_path)
    #     generate_and_paste_insights(ppt_config_path)
    #     print("✅ PPT update complete!")
    # # Option to run Excel automation followed by PPT update
    # elif key.lower() == 'b':
    #     automate_excel_pivoting(excel_config_path)
    #     update_ppt_from_excel(ppt_config_path)
    #     generate_and_paste_insights(ppt_config_path)
    #     print('✅ Excel pivoting and PPT update both complete!')
    # elif key.lower() == 'c':
    #     download_attachments(excel_config_path)
    #     automate_excel_pivoting(excel_config_path)
    #     update_ppt_from_excel(ppt_config_path)
    #     generate_and_paste_insights(ppt_config_path)
    #     print('✅ All tasks complete: Downloaded emails, updated Excel, and PPT.')
    # else:
    #     print('❌ Invalid option. Terminating this session.\nPlease re-run the script with a valid character input.')

    # Option to only update ppt using excel
    if key.lower()=='u':
        update_ppt_from_excel(ppt_config_path)
        print("PPT update complete!")
    # Option to only download email attachments
    elif key.lower()=='d':
        download_attachments(excel_config_path)
        print("Spend file downloaded from outlook!")
    # Option to only run Excel automation (e.g., pivoting or merging)
    elif key.lower()=='e':
        automate_excel_pivoting(excel_config_path)
        print("Excel - to - Excel update complete with data transformation!")
    # Option to only generate insights in ppt
    elif key.lower()=='i':
        generate_and_paste_insights(ppt_config_path)
        print("Insight generation in PPT complete!")
    # Option to only update the PowerPoint from the Excel data
    elif key.lower() == 'c':
        print("Starting Automation...\n")
        shutil.copy2(original_ppt_path, new_ppt_path)
        time.sleep(5)
        update_ppt_from_excel(ppt_config_path)
        generate_and_paste_insights(ppt_config_path)
        if os.path.exists(original_ppt_path):
            os.remove(original_ppt_path)
        os.rename(new_ppt_path, original_ppt_path)
        print('All tasks complete: PPT updation and Insight generation.')
        print("Resetting Excel cell E1 to original value...")
        reset_excel_cell(original_excel_path, "Overall Metrics", "E1", 45723)
    else:
        print('Invalid option. Terminating this session.\nPlease re-run the script with a valid character input.')


    