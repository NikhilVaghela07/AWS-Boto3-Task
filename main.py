import boto3
import openpyxl
from openpyxl.styles import Font

ec2_client = boto3.client('ec2')
cost_client = boto3.client('ce')
region_response = ec2_client.describe_regions()
availability_zone_responce = ec2_client.describe_availability_zones()


def Create_sheet_and_header(sheet_title, header):
    wb = openpyxl.Workbook()
    sheet = wb.active
    sheet.title = sheet_title
    headers = header
    sheet.append(headers)

    for test_bold in sheet[1]:
        test_bold.font = Font(bold = True)

    return wb, sheet

def fetch_all_region(region_response):
    all_regions = []
    for region in region_response["Regions"]:
        all_regions.append(f"{region['RegionName']}")
    return all_regions

def get_hourly_price():
    hourly_prices = {
        "t2.micro": 0.0116,
        "t2.small": 0.0230,
        "t2.medium": 0.0464,
        "t3.micro": 0.0104,
        "t3.small": 0.0208,
        "t3.medium": 0.0416,
        "t3.large": 0.0832,
        "m5.large": 0.0960,
        "c5.large": 0.0850,
    }
    return hourly_prices

def calculate_cost_create_sheet(sheet, all_regions, hourly_prices):
    for region in all_regions:
        ec2_resource = boto3.resource("ec2", region_name=region)
        for instance in ec2_resource.instances.all():
            inst_state = instance.state['Name']
            inst_type = instance.instance_type
            az = instance.placement.get('AvailabilityZone', 'N/A')
            key_name = instance.key_name or "N/A"
            public_ip = instance.public_ip_address or "N/A"
    
            if inst_state == "running":
                hourly_rate = hourly_prices.get(inst_type, 0.05)
                daliy_cost = round(hourly_rate * 24, 2)
                monthly_cost = round(hourly_rate * 24 * 30, 2)
            else:
                hourly_rate = 0
                monthly_cost = 0
                
            sheet.append([region, instance.architecture, instance.state['Name'], instance.instance_type, az, instance.instance_id, key_name, public_ip, hourly_rate, monthly_cost])
            print(region, instance.architecture, instance.state['Name'], instance.instance_type, az, instance.instance_id, key_name, public_ip, hourly_rate, monthly_cost)

    
sheet_title = "EC2 Instance Data"
header = ["Region", "Architecture", "State", "Instance type", "AvailabilityZone","Instance-Id", "Key/Pair Name", "Public-Ip", "Daily Cost($)", "Monthly Cost($)"]
wb, sheet = Create_sheet_and_header(sheet_title, header)
all_regions = fetch_all_region(region_response)
hourly_prices = get_hourly_price()
calculate_cost_create_sheet(sheet, all_regions, hourly_prices)
wb.save("EC2_Informations.xlsx")
print("Done")

