
import pandas as pd
import datetime
from datetime import timedelta
from io import StringIO
#import win32com.client
from imap_tools import MailBox
from glob import glob
import os
import aiohttp
import asyncio
import numpy as np
from mytracker_export_api import MyTracker
from sshtunnel import SSHTunnelForwarder
from zipfile import ZipFile
from sqlalchemy import create_engine
import psycopg2
from termcolor import colored
import requests

with SSHTunnelForwarder(    
    ('49.12.77.59', 22), #Remote server IP and SSH port
    ssh_username = "analitytic",
    ssh_password = "BAXqwCZukkdPW6OFbUnW",
    remote_bind_address=('127.0.0.1', 5432)
    ) as server:

    server.start()

    local_port = str(server.local_bind_port)
    engine = create_engine('postgresql://analist:xl3KQugjnp@127.0.0.1:' + local_port +'/ad_stats_system_db') 

    octa_daily = pd.read_sql(
        f'''
        SELECT*
        FROM 
            actual_month_report_bi
        WHERE 
            date = '2024-09-01'       
        ''',
        con=engine
    )  
    print(octa_daily.head())
    octa_daily.to_excel('C:/Users/Artur/Downloads/InApp 26.08.2024.xlsx', sheet_name='Лист1')
print(octa_daily)
server.stop()

