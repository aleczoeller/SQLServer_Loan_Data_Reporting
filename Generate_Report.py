# -*- coding: utf-8 -*-

"""
******************************************
*     Generate_Report.py                 *
*     ------------------                 *
*     Connect to SQL Database, retrieve  *
*     weekly loan data and email to      *
*     specified recipients as cleaned    *
*     Excel attachment. Also summarizes  *
*     loan data in other worksheets.     *
*                                        *
*     Author: Alec Zoeller, 2020         *
*                                        *
*                                        *
******************************************
"""

import os
import base64
import pyodbc
import pandas as pd
from pandas.io.json import json_normalize
import json
from datetime import datetime, timedelta
from getpass import getpass, getuser

import smtplib
import mimetypes
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from dotenv import load_dotenv
load_dotenv(verbose=True)


__author__ = 'Alec Zoeller'
__version__ = '1.0.1'


class DataSupport(object):
    """
    Support methods for data prep class.
    """
    @staticmethod
    def format_excel(writer, num_rows, df, sheetname, letters):
        """
        Format worksheet in Excel table for pandas export
        """
        print(sheetname)
        workbook = writer.book
        wsheet = writer.sheets[sheetname]
        #wsheet.set_column('A:Z', 16.0)
        headname = []
        for i in df.columns:
            headname.append(dict([('header', i)]))
        wsheet.add_table('A1:{0}{1}'.format(letters, str(num_rows)), 
                {'style': 'Table Style Medium 20',
                 'header_row':True, 'columns':headname})

class Send_Email(object):
    '''
    Support methods for Email object. 
    '''
    @staticmethod    
    def send_email(address, attachment, date):
        '''
        Method to send individual emails.
        '''
        #Connect to server. Specifically configured to Office365 login without token.
        #For advanced/tokenized O365 login, see shareplum library 
        s = smtplib.SMTP('smtp.office365.com', 587)
        s.ehlo()
        s.starttls()
        pwd = os.getenv('EMAIL')
        pwd = base64.b64decode(pwd).decode('ascii')
        from_address = base64.b64decode(os.getenv('FROM_EMAIL')).decode('ascii')
        s.login(from_address, pwd)
        #Prepare message
        from_email = from_address
        from_display = 'Weekly Reporting'
        date_display = datetime.strftime(date, '%m/%d/%Y')
        subject = 'Weekly Reporting - Week of {}'.format(date_display)
        mssg = f'<p>Hello,</p><p>This is an automatically generated weekly summary for loan '\
            f'volume and statistics. The data itemized in the attached table lists all pertinent '\
            f'information regarding loan, partner, channel and borrower. See additional worksheets '\
            f'in the Excel document for summary information on all close loans, as well as breakdowns for partners and '\
            f'channels.</p><p>To request that anyone else be added to this message, or to be removed'\
            f' from the mailing list feel free to reply to this message or email <a href="mailto'\
            f':{from_address}">{from_address}</a>. Thank you an have a great day.</p>'
        msg = MIMEMultipart()
        #Add Excel table attachment
        ctype, encoding = mimetypes.guess_type(attachment)
        if ctype is None or encoding is not None:
            ctype = 'application/octet-stream'
        maintype, subtype = ctype.split('/', 1)
        with open(attachment, 'rb') as fp:
            attach = MIMEBase(maintype, subtype)
            attach.set_payload(fp.read())
        encoders.encode_base64(attach)
        attach.add_header('Content-Disposition', 'attachment', filename=os.path.basename(attachment))
        msg.attach(attach)
        #Add subject and body to message, then send and quit SMTP object
        msg['Subject'] = subject
        msg['From'] = from_display
        msg['To'] = address
        body = MIMEText(mssg, 'html')
        msg.attach(body)
        s.sendmail(from_email, address, msg.as_string())
        s.quit()       
    


class Email(Send_Email):
    '''
    Class to set up and implement email alerts. 
    '''
    def __init__(self, email_list, daten, attachment):
        self.email_list = email_list
        self.daten = daten
        self.attachment = attachment
        
    def distribute_emails(self):
        '''
        Method to automatically email all recipients.
        '''
        for email in self.email_list:
            Send_Email.send_email(email, self.attachment, self.daten)

class DataGenerator(DataSupport):
    """
    Contains functionality for extracting data from database and 
    returning dataframe.
    """
    def __init__(self, query, conn, daten):
        self.query = query
        self.conn = conn
        self.daten = daten
        
    def generateTheD(self):
        """
        Ping database and pull weekly data, summarize and return
        full dataframe and all individual ones.
        """
        conn = self.conn
        query = self.query
        fulldf = pd.read_sql(con=conn, sql=query, parse_dates=True)
        #Extract json values from table/dataframe
        fulldf['json'] = fulldf.apply(lambda x: json.loads(x['jsondata']), axis=1)
        fulldf = json_normalize(fulldf['json'])
        fulldf.columns = [i.upper() for i in fulldf.columns]
        fulldf['CHANNEL'] = fulldf['CHANNEL'].fillna('no channel')
        fulldf.FINALIZED = pd.to_datetime(fulldf.FINALIZED)
        fulldf.sort_values(by='FINALIZED', inplace=True, ascending=True)
        outputdf = fulldf.copy()
        #Clean columns of master dataframe
        outputdf.drop(labels=['CORE_ID', 'PATH'], axis=1, inplace=True)
        ordered_cols = []
        for col in outputdf.columns:
            if col in ['EMP_YEARS', 'RATE', 'AGE', 'FICO', 'STATE', 'SCHOOL', 'OCCUPATION',
                       'DEGREE']:
                ordered_cols.append(1)
            elif col in ['STATUS', 'FINALIZED', 'VOLUME']:
                ordered_cols.append(0)
            else:
                ordered_cols.append(2)
        outputdf_columns = [x for _, x in sorted(zip(ordered_cols, outputdf.columns.tolist()))]
        outputdf = outputdf[outputdf_columns]
        #Prepare summary data columns and those for use in calculating mean, count and sum
        list_summ = ['CHANNEL', ['PROGRAM', 'TIER'], 'SCHOOL', 'STATE', 'OCCUPATION']
        summ_sheets = ['CHANNEL', 'PROGRAM&TIER', 'SCHOOL', 'STATE', 'OCCUPATION']
        stat_fields = ['VOLUME', 'EMP_YEARS', 'RATE', 'INCOME', 'AGE', 'FICO', 'LOAN_PAYMENT',
                        'TERM']
        fulldf['VOLUME_SUM'] = fulldf['VOLUME'].astype(float)
        for fld in stat_fields:
            fulldf[fld] = fulldf[fld].fillna(0.0)
            try:
                fulldf[fld] = fulldf.apply(lambda x: x[fld].replace('$', '').replace(',',''), axis=1)
            except:
                pass
            try:
                fulldf[fld] = fulldf[fld].astype(float)
            except:
                fulldf[fld] = fulldf.apply(lambda x: 0.0 if x[fld]=='' else x[fld], axis=1)
                fulldf[fld] = fulldf[fld].astype(float)
        #Create dictionary for applying statistics to summary tables
        dict_summ = {'APP_ID':'count','VOLUME_SUM':'sum'}
        for field in stat_fields:
            dict_summ[field] = 'mean'
        #Summarize data into supplemental dataframes
        channeldf = fulldf.loc[fulldf.STATUS=='Closed', :].groupby('CHANNEL').agg(dict_summ).reset_index()
        programdf = fulldf.loc[fulldf.STATUS=='Closed', :].groupby(['PROGRAM', 'TIER']).agg(dict_summ).reset_index()
        schooldf = fulldf.loc[fulldf.STATUS=='Closed', :].groupby('SCHOOL').agg(dict_summ).reset_index()
        statedf = fulldf.loc[fulldf.STATUS=='Closed', :].groupby('STATE').agg(dict_summ).reset_index()
        occupationdf = fulldf.loc[fulldf.STATUS=='Closed', :].groupby('OCCUPATION').agg(dict_summ).reset_index()
        fulldf.drop(labels='VOLUME_SUM', axis=1, inplace=True)
        #Get column lettes for Excel formatting
        first_letter = chr(int(len(fulldf.columns.tolist())/26)+64)
        second_letter = chr((int(len(fulldf.columns.tolist()))%26)+64) if \
                            int(len(fulldf.columns.tolist()))%26 > 0 else 'A'
        letters = first_letter + second_letter
        #Write full data to main excel table
        daten = datetime.strftime(self.daten, '%m%d%Y')
        report_path = os.path.join(os.path.dirname(__file__), 'Reports', 'Loan_Report_Week_of_{}.xlsx'.format(daten))
        writer = pd.ExcelWriter(report_path, engine='xlsxwriter')
        outputdf.to_excel(writer, 'WEEKLY_LOANS', index=False, header=outputdf.columns)
        DataSupport.format_excel(writer, len(outputdf) + 1, outputdf, 'WEEKLY_LOANS', letters)
        #Add worksheets for all the summay statistics
        letters = chr(len(channeldf.columns.tolist())%26 + 64)
        summ_dfs = [channeldf, programdf, schooldf, statedf, occupationdf]
        col_rename = {'APP_ID':'COUNT', 'VOLUME':'AVG_VOLUME', 'EMP_YEARS':'AVG_EMP',
                      'RATE':'AVG_RATE', 'INCOME':'AVG_INCOME', 'AGE':'AVG_AGE', 
                      'FICO':'AVG_FICO', 'LOAN_PAYMENT':'AVG_PAYMENT',
                      'TERM':'AVG_TERM'
                      }
        for i in range(len(summ_dfs)):
            summ_dfs[i].rename(col_rename, axis=1, inplace=True)
            summ_dfs[i]['AVG_VOLUME'] = summ_dfs[i].apply(lambda x: '${:,}'.format(round(float(x['AVG_VOLUME']),
                                            2)), axis=1)
            summ_dfs[i]['VOLUME_SUM'] = summ_dfs[i].apply(lambda x: '${:,}'.format(round(float(x['VOLUME_SUM']),
                                            2)), axis=1)
            summ_dfs[i]['AVG_INCOME'] = summ_dfs[i].apply(lambda x: '${:,}'.format(round(float(x['AVG_INCOME']),
                                            2)), axis=1)                                
            if summ_sheets[i] == 'PROGRAM&TIER':
                summ_dfs[i].to_excel(writer, summ_sheets[i], index=False, header=summ_dfs[i].columns)
                DataSupport.format_excel(writer, len(summ_dfs[i])+1, summ_dfs[i], 
                    summ_sheets[i], chr(len(summ_dfs[i].columns.tolist())%26 + 64))
            summ_dfs[i].to_excel(writer, summ_sheets[i], index=False, header=summ_dfs[i].columns)
            DataSupport.format_excel(writer, len(summ_dfs[i])+1, summ_dfs[i], summ_sheets[i], letters)
        #Save and close Excel table    
        writer.save()
        writer.close()        
        return report_path
        
        
        
def main():   
    """"
    Main method - create objects for pulling data and sending emails.
    """
    #Create SQL connection object and define query   
    pwd = os.getenv('DBUSER')
    pwd = base64.b64decode(pwd).decode('ascii')
    server = base64.b64decode(os.getenv('SERVER')).decode('ascii')
    database = base64.b64decode(os.getenv('DATABASE')).decode('ascii')
    user = base64.b64decode(os.getenv('USERNAME')).decode('ascii')
    conn = pyodbc.connect(user=user, password=pwd, 
                          driver='{SQL Server}', #Choose correct, installed driver for server
                         server=server, database=database)
    query = '''
    select jsondata as jsondata,
           modified as modified
    from dbo.loan loan
    where loan.finalized >= '{}'
    '''.format(datetime.strftime(datetime.now() - timedelta(days=7), '%m-%d-%Y'))
    datenow = datetime.now()

    #Generate dataframes and Excel table with SQL pull
    datagen = DataGenerator(query, conn, datenow)
    report_path = datagen.generateTheD()
    #Assume email list is in textfile in same repository, named emails.txt
    with open('emails.txt', 'r') as f:
        email_list = f.read().splitlines()
    #Email each recipient the Excel table
    email = Email(email_list, datenow, report_path)
    email.distribute_emails()

if __name__ == '__main__':
    main()

