"""This is the main process file for the robot framework."""
import json
import os
import glob
import pandas as pd
from OpenOrchestrator.orchestrator_connection.connection import OrchestratorConnection
from OpenOrchestrator.database.queues import QueueStatus
from robot_framework import config
from robot_framework.subprocesses.get_os2form_receipt import fetch_receipt
from robot_framework.subprocesses.outlay_ticket_creation import handle_opus


def process(orchestrator_connection: OrchestratorConnection) -> None:
    """Main process function."""
    orchestrator_connection.log_trace("Starting the process.")
    process_args = json.loads(orchestrator_connection.process_arguments)
    path_arg = process_args.get('path')
    os2_api_key = orchestrator_connection.get_credential("os2_api").password

    process_queue_elements(orchestrator_connection, config.QUEUE_NAME, QueueStatus.NEW, os2_api_key, path_arg)
    process_queue_elements(orchestrator_connection, config.QUEUE_NAME, QueueStatus.FAILED, os2_api_key, path_arg)

    orchestrator_connection.log_trace("Process completed.")


def process_queue_elements(orchestrator_connection, queue_name, status, os2_api_key, path_arg):
    """Process queue elements based on their status."""
    queue_elements = orchestrator_connection.get_queue_elements(queue_name, status=status)
    if queue_elements:
        orchestrator_connection.log_trace(f"Processing {len(queue_elements)} {'new' if status == QueueStatus.NEW else 'failed'} queue elements.")
        for queue_element in queue_elements:
            orchestrator_connection.set_queue_element_status(queue_element.id, QueueStatus.IN_PROGRESS)
            process_single_queue_element(queue_element, os2_api_key, path_arg, orchestrator_connection)
    else:
        orchestrator_connection.log_trace(f"No {'new' if status == QueueStatus.NEW else 'failed'} queue elements to process.")


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
