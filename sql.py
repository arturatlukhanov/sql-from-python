
import pandas as pd
import datetime
from datetime import timedelta
from io import StringIO
from imap_tools import MailBox
from glob import glob
import os
from mytracker_export_api import MyTracker
from sshtunnel import SSHTunnelForwarder
from sqlalchemy import create_engine
import psycopg2
import requests

with SSHTunnelForwarder(    
    ('', ), #Remote server IP and SSH port
    ssh_username = "",
    ssh_password = "",
    remote_bind_address=(')
    ) as server:

    server.start()

    local_port = str(server.local_bind_port)
    engine = create_engine('postgresql://' + local_port +'/') 

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

