# Fetch EC2 Instance & Cost Information

This script fetches EC2 instance details across all accessible AWS regions in your account. It retrieves the hourly On-Demand price for each instance, calculates the estimated daily and monthly costs, and exports the data into an Excel sheet.

## Use Cases

* Identify how many EC2 instances are running in each AWS region and view their region-wise costs.
* Calculate estimated daily and monthly costs for EC2 instances based on hourly On-Demand pricing.

---

## Prerequisites

* **Python** & **pip**
* **Boto3** (AWS SDK for Python)
* **Excel Export Library** ( `openpyxl`)
* **AWS CLI** (for configuring credentials)

### 1. Check Python and pip Installation (Linux)

Verify that Python 3 and pip are installed on your system:

```bash
python3 --version
pip3 --version
```

### 2. Install Boto3 & Openpyxl

Install the `boto3` library using pip:

```bash
pip3 install boto3
pip3 install openpyxl
```

---

## AWS Credentials Setup

### 1. Install AWS CLI v2 (Linux/macOS)

Download and run the official AWS CLI v2 installation script:

```bash
curl -fsSL 'https://awscli.amazonaws.com/v2/install.sh' | sudo bash -s -- --system
```

### 2. Configure Credentials

Run the following command to set up your AWS credentials:

```bash
aws configure
```

Provide the following details when prompted:

* **AWS Access Key ID**
* **AWS Secret Access Key**
* **Default region name** (e.g., `us-east-1`)
* **Default output format** (e.g., `json`)

This command creates a credentials file at `~/.aws/credentials`. Boto3 automatically detects and uses these credentials when you run the script.
---
## Run python file
```
python3 main.py
```