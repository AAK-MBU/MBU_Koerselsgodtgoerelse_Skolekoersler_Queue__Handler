## Kørselsgodtgørelse for Skolekørsler - Queue Handler

This robot is part of the 'MBU Koerselsgodtgoerelse Skolekoersler' process.

### Process Overview

The process consists of four main robots working in sequence:

1. **Create Excel and Upload to SharePoint**:  
   The first robot retrieves and exports weekly 'Egenbefordring' data from a database to an Excel file, which is then uploaded to SharePoint at the following location: `MBU - RPA - Egenbefordring/Dokumenter/General`. Once the file is processed, personnel will move it to `MBU - RPA - Egenbefordring/Dokumenter/General/Til udbetaling`. Run it with the 'Single Trigger' or with the Scheduled Trigger'.

2. **Queue Uploader**:  
   The second robot retrieves data from the Excel file and uploads it to the **Koerselsgodtgoerelse_egenbefordring** queue using [OpenOrchestrator](https://github.com/itk-dev-rpa/OpenOrchestrator). Run it with the 'Single Trigger'.

3. **Queue Handler (This robot)**:  
   The third robot, triggered by the Queue Trigger in OpenOrchestrator, processes the queue elements by creating tickets in OPUS.

4. **Update SharePoint**:  
   The fourth robot cleans and updates the files in SharePoint by uploading the updated Excel file and attachments of any failed elements. Run it with the 'Single Trigger'.

### The Queue Handler Process

Using the Queue Framework with modifications.

    - If no new elements it breaks.
    - If no browser open - it opens a browser -> opens OPUS
    - Fetches the receipt from OS2Forms.
    - Creates a ticket in OPUS and uploads the receipt.
    - Opens the Excel file and marks the entry as either failed or successfully handled.
    - If the check fails (when clicking the 'kontroller' button) is stops and fetches the next queue element.
    - If another unexpected error occurs is retries up until config.max_retries.

### Process and Related Robots

1. **Create Excel & Upload to SharePoint**: [Create Excel & Upload To SharePoint](https://github.com/AAK-MBU/MBU_Koerselsgodtgoerelse_Skolekoersler_Dan_Excel_Upload_Til_SharePoint)
2. **Queue Uploader** [Queue Uploader](https://github.com/AAK-MBU/MBU_Koerselsgodtgoerelse_Skolekoersler_Queue_Uploader).
3. **Queue Handler**: (This Robot)
4. **Update SharePoint**: [Update Sharepoint](https://github.com/AAK-MBU/MBU_Koerselsgodtgoerelse_Skolekoersler_Update_Sharepoint)

### Arguments

- **path**: The same path as the `path` argument in the uploader robot or the location where the Excel file is stored.
