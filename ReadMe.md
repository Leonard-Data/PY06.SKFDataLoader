# PY06.SKFDataLoader - Automation Workflow to Refresh Monthly Data from Dataverse and Upload to SharePoint

## Overview

The **PY06.SKFDataLoader** project automates the process of retrieving the current month's data from Dataverse and uploading it to SharePoint. This ensures that Power BI always has the latest data source available for reporting and analysis.

The workflow extracts data from Dataverse, processes it as needed, and then uploads it to a designated SharePoint location, eliminating manual intervention and ensuring data accuracy.

## Features

- **Automated Data Retrieval**: Fetches the latest month's data from Dataverse.
- **Data Processing**: Cleans and formats data before uploading.
- **Seamless Integration**: Uploads processed data to SharePoint for Power BI consumption.
- **Error Handling**: Logs errors and includes retry mechanisms to ensure a successful update.
  
## Prerequisites

Before running this automation workflow, ensure that you have the following installed and set up:

- Python 3.x
- Required Python libraries (listed below)
- Access to Dataverse with appropriate credentials
- Access to SharePoint for data upload

### Install Python

You can download and install Python from the official website: [https://www.python.org/downloads/](https://www.python.org/downloads/)

Once Python is installed, you can verify the installation using the following command in your terminal:

```bash
python --version
```

