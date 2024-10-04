## MBU Kørselsgodtgørelse for Skolekørsler - Queue Handler

This robot is built for [OpenOrchestrator](https://github.com/itk-dev-rpa/OpenOrchestrator).

It handles elements from the **Koerselsgodtgoerelse_egenbefordring** queue by fetching a receipt and creating a ticket in OPUS. The related queue uploader robot can be found here: [Queue Uploader](https://github.com/AAK-MBU/MBU_Koerselsgodtgoerelse_Skolekoersler_Queue_Uploader).

### Process:

1. **Processing NEW Queue Elements:**
    - Fetches the receipt from OS2Forms.
    - Creates a ticket in OPUS and uploads the receipt.
    - Opens the Excel file and marks the entry as either failed or successfully handled.

2. **Processing FAILED Queue Elements:**
    - Fetches the receipt from OS2Forms.
    - Creates a ticket in OPUS and uploads the receipt.
    - Opens the Excel file and marks the entry as either failed or successfully handled.

3. **Post process:**
    - Uploads the Excel file to the "Behandlet" (Processed) or "Fejlet" (Failed) folder in SharePoint.
    - For failed queue elements, it also uploads the associated attachments.

### Arguments:

- **path**: The same path as the `path` argument in the uploader robot or the location where the Excel file is stored.
