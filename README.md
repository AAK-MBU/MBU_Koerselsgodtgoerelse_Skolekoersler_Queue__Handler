## MBU_Kørselsgodtgørelse for skolekørsler - Queue Handler

This robot is made for [OpenOrchestrator](https://github.com/itk-dev-rpa/OpenOrchestrator).

It handles elements for "Koerselsgodtgoerelse_egenbefordring" queue by fetching a receipt and creating a ticket in OPUS.

Using the queue flow.

### Queue Flow

The queue framework is used when the robot is doing multiple bite-sized tasks defined in an
OpenOrchestrator queue.
The flow of the queue framework is sketched up in the following illustration:

![Queue Flow diagram](Robot-Queue-Framework.svg)

