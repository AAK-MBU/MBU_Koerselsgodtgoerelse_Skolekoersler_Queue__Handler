"""This is the main process file for the robot framework."""
import json
import os
import glob
import pandas as pd
from OpenOrchestrator.orchestrator_connection.connection import OrchestratorConnection
from OpenOrchestrator.database.queues import QueueStatus
from subprocesses.get_os2form_receipt import fetch_receipt

DIR_PATH = None


def process(orchestrator_connection: OrchestratorConnection, queue_element, browser) -> None:
    """Main process function."""
    orchestrator_connection.log_trace("Starting the process.")
    process_args = json.loads(orchestrator_connection.process_arguments)
    path_arg = process_args.get('path')

    global DIR_PATH
    DIR_PATH = path_arg

    os2_api_key = orchestrator_connection.get_credential("os2_api").password
    process_single_queue_element(queue_element, os2_api_key, path_arg, browser, orchestrator_connection)

    orchestrator_connection.log_trace("Process completed.")


def process_single_queue_element(queue_element, os2_api_key, path_arg, browser, orchestrator_connection):
    """Process a single queue element."""
    from robot_framework.subprocesses.outlay_ticket_creation import handle_opus
    element_data = json.loads(queue_element.data)
    orchestrator_connection.set_queue_element_status(queue_element.id, QueueStatus.IN_PROGRESS)
    orchestrator_connection.log_trace(f"Processing queue element ID: {queue_element.id}")

    folder_path = fetch_receipt(queue_element, os2_api_key, path_arg, orchestrator_connection)
    handle_opus(queue_element, folder_path, browser, orchestrator_connection)
    remove_attachment_if_exists(folder_path, element_data, orchestrator_connection)
    handle_post_process(False, queue_element, orchestrator_connection)


def remove_attachment_if_exists(folder_path, element_data, orchestrator_connection):
    """Remove the attachment file if it exists."""
    attachment_path = os.path.join(folder_path, f'receipt_{element_data["uuid"]}.pdf')
    if os.path.exists(attachment_path):
        orchestrator_connection.log_trace(f"Removing attachment file: {attachment_path}")
        os.remove(attachment_path)


def handle_post_process(failed, queue_element, orchestrator_connection):
    """Update the Excel file with the status of the element."""
    element_data = json.loads(queue_element.data)
    uuid = element_data['uuid']
    excel_filename = element_data['filename']

    excel_files = glob.glob(os.path.join(DIR_PATH, excel_filename))
    if not excel_files:
        raise FileNotFoundError(f"{excel_filename} not found in {DIR_PATH}.")

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
        df.loc[df['uuid'] == uuid, 'behandlet_fejl'] = ' '
    else:
        df.loc[df['uuid'] == uuid, 'behandlet_ok'] = ' '
