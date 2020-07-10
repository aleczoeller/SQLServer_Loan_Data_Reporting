# SQLServer_Loan_Data_Reporting
Python script that pulls SQL data in pandas, cleans, summarizes and emails weekly reports to selected recipients.

Main method creates an odbc connection object and establishes a link with an Azure-hosted SQL database. This can be readily modified to accomodate any source.

Data is transformmed from json format with pandas and summarized to list out the previous week's activity, along with a table summarizing.  This occurs as a support class object. 

An email-functionality class creates an object that connects to an Office365 SMTP server.  Non-O365 servers can be accomodated easily be removing specific references to a port and changing hostname. As mentioned in the comments, check out the Shareplum library if you need to authenticate with Cookies with O365.

Assuming a list of email recipients as separate lines in a local text file, this script will neatly export the pandas DataFrames into properly formatted Excel documents and email to each one, with a specified message.  

Any one of the elements of this process (connecting to Azure db, ETL, formatting Excel, and emailing with O365 account) can be helpful in itself.  
