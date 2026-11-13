import boto3
import json
import openpyxl
from openpyxl.styles import Font

ec2_client = boto3.client('ec2')
pricing_client = boto3.client('pricing', region_name='us-east-1')
print(pricing_client)
region_response = ec2_client.describe_regions()

# this fucntion will create empty sheet and append header in sheet
def create_sheet_and_header(sheet_title, header):
    wb = openpyxl.Workbook()
    sheet = wb.active
    sheet.title = sheet_title
    sheet.append(header)

    for cell in sheet[1]:
        cell.font = Font(bold=True)

    return wb, sheet

# fetch all regions that account have access
def fetch_all_region(region_response):
    all_regions = []
    for region in region_response["Regions"]:
        all_regions.append(f"{region['RegionName']}")
    return all_regions

# return os type 
def get_os_type(instance):
    platform_details = getattr(instance, "platform_details", None)
    if platform_details:
        if "Windows" in platform_details:
            return "Windows"
        elif "Red Hat" in platform_details or "RHEL" in platform_details:
            return "RHEL"
        elif "SUSE" in platform_details:
            return "SUSE"
        elif "Linux/UNIX" in platform_details:
            return "Linux"
    return "Linux"

# this function return instance name, if have no instance name then return N/A
def get_instance_name(instance):
    for tag in instance.tags or []:
        if tag["Key"] == "Name":
            return tag["Value"]
    return "N/A"

# region wise instance tags will be fetch from aws account
def get_instance_tags(instance):
    return ", ".join(
        f"{tag['Key']}={tag['Value']}" for tag in instance.tags or []
    ) or "N/A"

# base on region instance hourly price will be fetch from aws 
def get_hourly_price(region, instance_type, operating_system, price_cache):
    key = (region, instance_type, operating_system)
    if key in price_cache:
        return price_cache[key]

    result = pricing_client.get_products(
        ServiceCode="AmazonEC2",
        Filters=[
            {"Type": "TERM_MATCH", "Field": "regionCode", "Value": region},
            {"Type": "TERM_MATCH", "Field": "instanceType", "Value": instance_type},
            {"Type": "TERM_MATCH", "Field": "operatingSystem", "Value": operating_system},
            {"Type": "TERM_MATCH", "Field": "capacitystatus", "Value": "Used"},
            {"Type": "TERM_MATCH", "Field": "tenancy", "Value": "Shared"},
            {"Type": "TERM_MATCH", "Field": "preInstalledSw", "Value": "NA"},
            {"Type": "TERM_MATCH", "Field": "marketoption", "Value": "OnDemand"},
        ],
        MaxResults=1,
    )
    if not result["PriceList"]:
        price_cache[key] = None
        return None

    product = json.loads(result["PriceList"][0])
    term = next(iter(product["terms"]["OnDemand"].values()))
    dimension = next(iter(term["priceDimensions"].values()))
    price = float(dimension["pricePerUnit"]["USD"])
    price_cache[key] = price
    return price

# Function will calculate cost of ec2 instance and append data in sheet
def calculate_cost_and_append_data_in_sheet(sheet, all_regions):
    price_cache = {}
    for region in all_regions:
        ec2_resource = boto3.resource("ec2", region_name=region)
        for instance in ec2_resource.instances.all():
            # print(instance)
            inst_state = instance.state['Name']
            inst_type = instance.instance_type
            os_type = get_os_type(instance)
            instance_name = get_instance_name(instance)
            instance_tags = get_instance_tags(instance)

            if inst_state == "running":
                hourly_rate = get_hourly_price(
                    region,
                    inst_type,
                    os_type,
                    price_cache,
                )
                if hourly_rate is None:
                    print(
                        f"Price not found for {region}, {inst_type}, "
                        f"{os_type}"
                    )
                    daily_cost = None
                    monthly_cost = None
                else:
                    daily_cost = round(hourly_rate * 24, 2)
                    monthly_cost = round(hourly_rate * 24 * 30, 2)
            else:
                hourly_rate = 0
                daily_cost = 0
                monthly_cost = 0

            sheet.append([
                region,
                instance_name,
                instance_tags,
                inst_state,
                os_type,
                inst_type,
                instance.placement["AvailabilityZone"],
                hourly_rate,
                daily_cost,
                monthly_cost,
            ])
            print(region, instance_name , instance_tags, inst_state, os_type, inst_type, instance.placement["AvailabilityZone"], hourly_rate, daily_cost,monthly_cost)

sheet_title = "EC2 Instance Data"
header = [
    "Region",
    "Instance Name",
    "Tags",
    "State",
    "OS Type",
    "Instance type",
    "AvailabilityZone",
    "Hourly Cost($)",
    "Daily Cost($)",
    "Monthly Cost($)",
]

wb, sheet = create_sheet_and_header(sheet_title, header)
all_regions = fetch_all_region(region_response)
calculate_cost_and_append_data_in_sheet(sheet, all_regions)
# sheet will be save in current python file folder
wb.save("EC2_Informations.xlsx") 