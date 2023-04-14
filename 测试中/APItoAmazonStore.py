import os
import requests
import boto3
import datetime
import schedule
import time

# 获取Amazon Advertising API访问令牌
def get_amazon_advertising_api_credentials():
    url = 'https://api.amazon.com/auth/o2/token'
    headers = {'Content-Type': 'application/x-www-form-urlencoded'}
    data = {
        'grant_type': 'client_credentials',
        'client_id': 'your_amazon_client_id',
        'client_secret': 'your_amazon_client_secret',
        'scope': 'cpc_advertising:campaign_management',
    }

    response = requests.post(url, headers=headers, data=data)
    if response.status_code == 200:
        return response.json()
    else:
        print(f'Error: {response.status_code}')
        return None

# 下载广告Bulk文件
def download_ad_bulk(start_date, end_date):
    token_data = get_amazon_advertising_api_credentials()
    access_token = token_data['access_token']
    profile_id = 'your_amazon_profile_id'

    url = f'https://advertising-api.amazon.com/v2/profiles/{profile_id}/download'
    headers = {'Authorization': f'Bearer {access_token}', 'Content-Type': 'application/json'}
    data = {'reportDate': start_date, 'campaignType': 'sponsoredProducts'}

    response = requests.post(url, headers=headers, json=data)
    if response.status_code == 200:
        return response.json()
    else:
        print(f'Error: {response.status_code}')
        return None

# 下载业务报告文件
def download_business_report(start_date, end_date):
    mws_auth_token = 'your_mws_auth_token'
    seller_id = 'your_seller_id'

    mws_client = boto3.client(
        'mws',
        region_name='your_region_name',
        aws_access_key_id='your_aws_access_key_id',
        aws_secret_access_key='your_aws_secret_access_key',
        mws_auth_token=mws_auth_token,
        seller_id=seller_id,
    )

    report_response = mws_client.request_report(
        ReportType='_GET_AMAZON_FULFILLED_SHIPMENTS_DATA_',
        StartDate=start_date,
        EndDate=end_date,
    )
    report_request_id = report_response['ReportRequestInfo']['ReportRequestId']

    report_ready = False
    while not report_ready:
        time.sleep(60)  # 等待60秒后再次检查报告状态
        report_request_list = mws_client.get_report_request_list(ReportRequestIdList=[report_request_id])
        report_request_info = report_request_list['ReportRequestInfo']
        report_status = report_request_info['ReportProcessingStatus']

        if report_status == '_DONE_':
            report_ready = True
            report_id = report_request_info['GeneratedReportId']
            report = mws_client.get_report(ReportId=report_id)
            return report['Report']
        elif report_status == '_CANCELLED_' or report_status == '_DONE_NO_DATA_':
            print("Report processing cancelled or no data available")
            return None

# 计划任务
def download_daily_reports():
    yesterday = datetime.datetime.now() - datetime.timedelta(days=1)
    ad_bulk_report = download_ad_bulk(yesterday, yesterday)
    business_report = download_business_report(yesterday, yesterday)

def download_weekly_reports():
    today = datetime.datetime.now()
    last_week_start = today - datetime.timedelta(days=today.weekday() + 7)
    last_week_end = today - datetime.timedelta(days=today.weekday() + 1)
    ad_bulk_report = download_ad_bulk(last_week_start, last_week_end)
    business_report = download_business_report(last_week_start, last_week_end)

# 每天下午5点下载前一天的报告
schedule.every().day.at("17:00").do(download_daily_reports)

# 每周一上午9点下载上一周的报告
schedule.every().monday.at("09:00").do(download_weekly_reports)

# 将整个程序打包到一个Python脚本中，并在你的服务器上运行
if __name__ == '__main__':
    while True:
        schedule.run_pending()
        time.sleep(60)  # 等待60秒后检查是否有待执行的任务
