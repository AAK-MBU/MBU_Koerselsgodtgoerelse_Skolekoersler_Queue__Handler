"""This is the main process file for the robot framework."""
import json
import os
import glob
import pandas as pd
from OpenOrchestrator.orchestrator_connection.connection import OrchestratorConnection
from OpenOrchestrator.database.queues import QueueStatus
from office365.runtime.auth.user_credential import UserCredential
from office365.sharepoint.client_context import ClientContext
from robot_framework import config
from robot_framework.subprocesses.get_os2form_receipt import fetch_receipt
from robot_framework.subprocesses.outlay_ticket_creation import handle_opus


def process(orchestrator_connection: OrchestratorConnection) -> None:
    """Main process function."""
    orchestrator_connection.log_trace("Starting the process.")
    process_args = json.loads(orchestrator_connection.process_arguments)
    path_arg = process_args.get('path')
    os2_api_key = orchestrator_connection.get_credential("os2_api").password
    service_konto_credential = orchestrator_connection.get_credential("SvcRpaMBU002")
    username = service_konto_credential.username
    password = service_konto_credential.password

    first_element = process_queue_elements(orchestrator_connection, config.QUEUE_NAME, QueueStatus.NEW, os2_api_key, path_arg)
    first_element = process_queue_elements(orchestrator_connection, config.QUEUE_NAME, QueueStatus.FAILED, os2_api_key, path_arg)

    if first_element:
        first_element_data = json.loads(first_element.data)
        filename = first_element_data['filename']

        update_sharepoint(orchestrator_connection, path_arg, filename, username, password)

    orchestrator_connection.log_trace("Process completed.")


def process_queue_elements(orchestrator_connection, queue_name, status, os2_api_key, path_arg):
    """Process queue elements based on their status."""
    queue_elements = orchestrator_connection.get_queue_elements(queue_name, status=status)
    if queue_elements:
        first_element = queue_elements[0]
        orchestrator_connection.log_trace(f"Processing {len(queue_elements)} {'new' if status == QueueStatus.NEW else 'failed'} queue elements.")
        for queue_element in queue_elements:
            orchestrator_connection.set_queue_element_status(queue_element.id, QueueStatus.IN_PROGRESS)
            process_single_queue_element(queue_element, os2_api_key, path_arg, orchestrator_connection)
        return first_element

    orchestrator_connection.log_trace(f"No {'new' if status == QueueStatus.NEW else 'failed'} queue elements to process.")
    return None


def process_single_queue_element(queue_element, os2_api_key, path_arg, orchestrator_connection):
    """Process a single queue element."""
    try:
        orchestrator_connection.set_queue_element_status(queue_element.id, QueueStatus.IN_PROGRESS)
        orchestrator_connection.log_trace(f"Processing queue element ID: {queue_element.id}")

        folder_path = fetch_receipt(queue_element, os2_api_key, path_arg, orchestrator_connection)
        status, error = handle_opus(queue_element, folder_path, orchestrator_connection)

        error_status = {"status": status, "error": error}
        handle_processing_status(error_status, queue_element, folder_path, path_arg, orchestrator_connection)

    except Exception as e:  # pylint: disable=broad-except
        handle_unexpected_error(e, queue_element, orchestrator_connection, path_arg)


def handle_processing_status(error_status, queue_element, folder_path, path_arg, orchestrator_connection):
    """Handle the status of processing a queue element."""
    element_data = json.loads(queue_element.data)
    status = error_status.get("status")
    error = error_status.get("error")
    failed = status != "Completed"
    if status == "Completed":
        orchestrator_connection.set_queue_element_status(queue_element.id, QueueStatus.DONE)
        remove_attachment_if_exists(folder_path, element_data, orchestrator_connection)

    if status in ["BusinessError", "RobotError"]:
        handle_error(error, f"{status}: Check failed", queue_element, orchestrator_connection)

    handle_post_process(failed, element_data, orchestrator_connection, path_arg)


def remove_attachment_if_exists(folder_path, element_data, orchestrator_connection):
    """Remove the attachment file if it exists."""
    attachment_path = os.path.join(folder_path, f'receipt_{element_data["uuid"]}.pdf')
    if os.path.exists(attachment_path):
        orchestrator_connection.log_trace(f"Removing attachment file: {attachment_path}")
        os.remove(attachment_path)


def handle_error(error, error_type, queue_element, orchestrator_connection):
    """Handle errors."""
    orchestrator_connection.log_error(f"{error_type}: {error if error else ''}")
    orchestrator_connection.set_queue_element_status(queue_element.id, QueueStatus.FAILED)


def handle_unexpected_error(error, queue_element, orchestrator_connection, path_arg):
    """Handle unexpected errors."""
    orchestrator_connection.log_error(f"Unexpected Error: {error}")
    if queue_element:
        orchestrator_connection.set_queue_element_status(queue_element.id, QueueStatus.FAILED)
        handle_post_process(True, json.loads(queue_element.data), orchestrator_connection, path_arg)


def handle_post_process(failed, element_data, orchestrator_connection, dir_path):
    """Update the Excel file with the status of the element."""
    uuid = element_data['uuid']
    excel_filename = element_data['filename']

    excel_files = glob.glob(os.path.join(dir_path, excel_filename))
    if not excel_files:
        raise FileNotFoundError(f"{excel_filename} not found in {dir_path}.")

    file_to_read = excel_files[0]
    df = pd.read_excel(file_to_read, engine='openpyxl')
    df = ensure_columns(df)
    update_dataframe(df, uuid, failed)

    with pd.ExcelWriter(file_to_read, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)

    orchestrator_connection.log_trace(f"Element status updated to {'failed' if failed else 'succeeded'} in Excel file")


def ensure_columns(df):
    """Ensure that the Excel file has the necessary columns."""
    for col in ['behandlet_fejl', 'behandlet_ok']:
        if col not in df.columns:
            df[col] = ''
    df['behandlet_fejl'] = df['behandlet_fejl'].astype(str)
    df['behandlet_ok'] = df['behandlet_ok'].astype(str)
    return df


def update_dataframe(df, uuid, failed):
    """Update the dataframe with the status of the element."""
    df.loc[df['uuid'] == uuid, 'behandlet_fejl' if failed else 'behandlet_ok'] = 'x'
    if not failed:
        df.loc[df['uuid'] == uuid, 'behandlet_fejl'] = ''


def update_sharepoint(orchestrator_connection, path_arg, filename, username, password):
    """Update the SharePoint folders."""
    orchestrator_connection.log_trace("Updating SharePoint folders.")
    failed_elements = orchestrator_connection.get_queue_elements(config.QUEUE_NAME, status=QueueStatus.FAILED)
    if failed_elements:
        orchestrator_connection.log_trace("Moving Excel file and failed attachments to the failed folder.")
        folder_name = os.path.splitext(filename)[0]
        upload_file_to_sharepoint(username, password, path_arg, filename, "Fejlet")
        upload_folder_to_sharepoint(username, password, path_arg, folder_name, "Fejlet")
    else:
        orchestrator_connection.log_trace("Uploading Excel file to the 'Behandlet' folder.")
        upload_file_to_sharepoint(username, password, path_arg, filename, "Behandlet")
    delete_file_from_sharepoint(username, password, filename)
    orchestrator_connection.log_trace("SharePoint folders updated.")


def upload_file_to_sharepoint(username: str, password: str, path_arg: str, excel_filename: str, sharepoint_folder_name: str) -> None:
    """Upload a file to SharePoint."""
    sharepoint_site_url = "https://aarhuskommune.sharepoint.com/teams/MBU-RPA-Egenbefordring"
    document_library = f"Delte dokumenter/General/Til udbetaling/{sharepoint_folder_name}"
    ctx = ClientContext(sharepoint_site_url).with_credentials(UserCredential(username, password))
    target_folder_url = f"/teams/MBU-RPA-Egenbefordring/{document_library}"
    target_folder = ctx.web.get_folder_by_server_relative_url(target_folder_url)
    file_path = os.path.join(path_arg, excel_filename)
    with open(file_path, "rb") as file_content:
        target_folder.upload_file(excel_filename, file_content).execute_query()

    print(f"File '{excel_filename}' has been uploaded successfully to SharePoint in '{sharepoint_folder_name}'.")


def upload_folder_to_sharepoint(username: str, password: str, path_arg: str, folder_name: str, sharepoint_folder_name: str) -> None:
    """Upload a folder and its contents to SharePoint."""
    sharepoint_site_url = "https://aarhuskommune.sharepoint.com/teams/MBU-RPA-Egenbefordring"
    document_library = f"Delte dokumenter/General/Til udbetaling/{sharepoint_folder_name}"
    ctx = ClientContext(sharepoint_site_url).with_credentials(UserCredential(username, password))
    target_folder_url = f"/teams/MBU-RPA-Egenbefordring/{document_library}/{folder_name}"
    ctx.web.folders.add(target_folder_url).execute_query()
    print(f"Folder '{folder_name}' created in SharePoint.")

    local_folder_path = os.path.join(path_arg, folder_name)
    updated_sharepoint_folder_name = f"{sharepoint_folder_name}/{folder_name}"

    if os.path.exists(local_folder_path):
        for file_name in os.listdir(local_folder_path):
            file_full_path = os.path.join(local_folder_path, file_name)
            if os.path.isfile(file_full_path):
                upload_file_to_sharepoint(username, password, local_folder_path, file_name, updated_sharepoint_folder_name)

    print(f"Folder '{folder_name}' and its contents have been uploaded successfully to SharePoint.")


def delete_file_from_sharepoint(username: str, password: str, file_name: str) -> None:
    """Delete a file from SharePoint."""
    sharepoint_site_url = "https://aarhuskommune.sharepoint.com/teams/MBU-RPA-Egenbefordring"
    document_library = "Delte dokumenter/General/Til udbetaling"
    ctx = ClientContext(sharepoint_site_url).with_credentials(UserCredential(username, password))
    target_file_url = f"/teams/MBU-RPA-Egenbefordring/{document_library}/{file_name}"
    try:
        file = ctx.web.get_file_by_server_relative_url(target_file_url)
        file.delete_object()
        ctx.execute_query()

        print(f"File '{file_name}' has been deleted successfully from SharePoint.")
    except Exception as e:  # pylint: disable=broad-except
        print(f"Error deleting file '{file_name}': {e}")
