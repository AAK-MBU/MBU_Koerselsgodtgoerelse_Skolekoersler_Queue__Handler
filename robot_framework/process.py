"""This module contains the main process of the robot."""
import json
from OpenOrchestrator.orchestrator_connection.connection import OrchestratorConnection
from OpenOrchestrator.database.queues import QueueStatus
from robot_framework import config
from robot_framework.subprocesses.get_os2form_receipt import fetch_receipt
from robot_framework.subprocesses.outlay_ticket_creation import initialize_browser, handle_opus


def process(orchestrator_connection: OrchestratorConnection) -> None:
    """Execute the primary process of the robot."""
    orchestrator_connection.log_trace("Starting the process.")

    try:
        process_args = json.loads(orchestrator_connection.process_arguments)
        path_arg = process_args.get('path')

        os2_api_credential = orchestrator_connection.get_credential("os2_api")
        os2_api_key = os2_api_credential.password

        queue_element = orchestrator_connection.get_next_queue_element(config.QUEUE_NAME)

        fetch_receipt(queue_element, os2_api_key, path_arg, orchestrator_connection)
        browser = initialize_browser()
        handle_opus(browser, queue_element, path_arg, orchestrator_connection)

    except Exception as e:
        orchestrator_connection.log_error(f"An error occurred during the process: {e}")
        orchestrator_connection.set_queue_element_status(queue_element.id, QueueStatus.FAILED)
    finally:
        orchestrator_connection.log_trace("Process completed.")
