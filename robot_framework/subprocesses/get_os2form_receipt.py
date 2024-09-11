"""This module contains the logic for fetching a receipt from OS2FORMS."""
import json
import os
from mbu_dev_shared_components.os2forms import documents
from OpenOrchestrator.database.queues import QueueStatus
import requests


class FileDownloadError(Exception):
    """An exception raised when a file download fails."""


def fetch_receipt(queue_element, os2_api_key, path, orchestrator_connection):
    """Fetch a receipt from OS2FORMS and save it to the specified path."""
    element_data = json.loads(queue_element.data)
    url = element_data.get('attachment')
    uuid = element_data.get('uuid')

    if not url or not uuid:
        error_message = "Missing 'attachment' URL or 'uuid' in element data."
        print(error_message)

    try:
        orchestrator_connection.log_info("Downloading file from OS2FORMS.")
        # Download the file bytes
        file_content = documents.download_file_bytes(url, os2_api_key)

        if not os.path.exists(path):
            os.makedirs(path)

        filename = f"receipt_{uuid}.pdf"
        file_path = os.path.join(path, filename)

        # Save the file content
        with open(file_path, 'wb') as f:
            f.write(file_content)

        orchestrator_connection.log_trace(f"File downloaded and saved successfully to {file_path}.")

    except requests.exceptions.RequestException as e:
        error_message = "Network error downloading file from OS2FORMS"
        orchestrator_connection.log_error(error_message)
        orchestrator_connection.set_queue_element_status(queue_element.id, QueueStatus.FAILED, error_message)
        raise FileDownloadError(error_message) from e

    except OSError as e:
        error_message = "File system error while saving the file"
        orchestrator_connection.log_error(error_message)
        orchestrator_connection.set_queue_element_status(queue_element.id, QueueStatus.FAILED, error_message)
        raise FileDownloadError(error_message) from e
