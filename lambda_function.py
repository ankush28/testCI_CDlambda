import boto3
import openpyxl
from openpyxl.styles import Font
from datetime import datetime
import os


dynamodb_client = boto3.client('dynamodb')
cloudwatch_client = boto3.client('cloudwatch')
s3_client = boto3.client('s3')

def lambda_handler(event, context):
    tables = dynamodb_client.list_tables()['TableNames']
    
    # Create a workbook and a worksheet
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "DynamoDB Metrics"
    
    # Define the headers
    headers = ["Table Name", "Consumed Read Capacity Units", "Consumed Write Capacity Units", "Table Size (Bytes)"]
    ws.append(headers)
    
    # Style the headers
    for cell in ws["1:1"]:
        cell.font = Font(bold=True)
    
    # Collect metrics for each table
    for table_name in tables:
        # Fetch table description
        table_description = dynamodb_client.describe_table(TableName=table_name)['Table']
        
        # Fetch table size
        table_size = table_description['TableSizeBytes']
        
        # Fetch consumed capacity units
        read_capacity = get_consumed_capacity(table_name, 'ReadCapacityUnits')
        write_capacity = get_consumed_capacity(table_name, 'WriteCapacityUnits')
        
        # Append data to the worksheet
        ws.append([table_name, read_capacity, write_capacity, table_size])
    
    file_name = f"/tmp/dynamodb_metrics_{datetime.now().strftime('%Y-%m-%d')}.xlsx"
    wb.save(file_name)
    
    s3_bucket = 'testwss3333'
    s3_key = f"dynamodb_reports/{os.path.basename(file_name)}"
    s3_client.upload_file(file_name, s3_bucket, s3_key)
    
    return {
        'statusCode': 200,
        'body': f"Report generated and uploaded to S3: {s3_key}"
    }

def get_consumed_capacity(table_name, metric_name):
    response = cloudwatch_client.get_metric_statistics(
        Namespace='AWS/DynamoDB',
        MetricName=metric_name,
        Dimensions=[
            {'Name': 'TableName', 'Value': table_name}
        ],
        StartTime=datetime.utcnow().replace(hour=0, minute=0, second=0, microsecond=0),
        EndTime=datetime.utcnow(),
        Period=60,
        Statistics=['Sum']
    )
    return response['Datapoints'][0]['Sum'] if response['Datapoints'] else 0
