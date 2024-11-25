from bs4 import BeautifulSoup
import sys
import datetime
from zipfile import ZipFile
import html
import os
import shutil
from selenium.webdriver.edge.service import Service as EdgeService
from webdriver_manager.microsoft import EdgeChromiumDriverManager
from Diff import diff_match_patch
from selenium.webdriver.common.by import By
from webdriver_manager.microsoft import EdgeChromiumDriverManager
from selenium.webdriver.edge.service import Service as EdgeService
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
import requests
import pandas as pd
import sqlalchemy
from sqlalchemy import text
from sqlalchemy import insert
from sqlalchemy import MetaData, Table, Column
import re
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import formatdate
from email import encoders


print("Start Checking")
# from PIL import Image, ImageChops
collect_sites = []

collect_data = []
import pandas as pd
from html_sanitizer import Sanitizer

sanitizer = Sanitizer()

main_dir = "D:\\A-ABORDOAL\\PG_Court"
states_folder = main_dir + "\\States\\"


engine_stmt = (
    "mssql+pyodbc://ContentTeamAPP:U98=k9M23H&Cgw5MW6tDBgAp9@RETDA2SQLD150/ContentTeamOPEX?driver=SQL Server")
engine = sqlalchemy.create_engine(engine_stmt)
with engine.connect() as conn1:
        sql = "SELECT * FROM PG_Court_Rules"
        df = pd.read_sql(sql, conn1)
        df.to_excel(main_dir + "\\data.xlsx", index=False)



df = pd.read_excel(main_dir + "\\data.xlsx")

options = Options()
try:
    options.add_experimental_option("excludeSwitches", ["enable-logging"])
except:
    pass

edge_options=webdriver.EdgeOptions()
try:
    edge_options.add_experimental_option("excludeSwitches", ["enable-logging"])
    edge_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/118.0.0.0 Safari/537.36")
    # edge_options.add_argument('--headless')
except:
    pass
# edge_options.add_argument("--headless=new")
edge_options.add_argument("start-maximized")
edge_options.add_argument("disable-gpu")
edge_options.add_argument("disable-dev-shm-usage")

driver = webdriver.Edge(service=EdgeService(EdgeChromiumDriverManager().install()), options=edge_options)


def clean_output(rawtext):
    # Parse the HTML content
    rawtext = rawtext.replace('&para;<br>','')
    soup = BeautifulSoup(rawtext, 'html.parser')

# Iterate over all tags except for the img tag
    for tag in soup.find_all(lambda tag: tag.name != 'a'):
        # Remove all attributes
        tag.attrs = {}
        # Print the modified HTML
    # print(str(soup.prettify()))
    return str(soup.prettify())


# send_mail("anthony.abordo@reedelsevier.com", "anthony.abordo@reedelsevier.com", "Court Website Alert", 'Starting checking sites.', 'lngdayappp007.legal.regn.net', 25)
# ob = Screenshot.Screenshot()
events_count = len(df)
print(events_count)
i = 1
for index, row in df.iterrows():
    finished = 100 * (i / events_count)
    i = i + 1
    print(str(i))
    with open(main_dir + "\\Data.txt", "w", encoding="utf-8") as f:
        f.write(str(int(finished)))

    # orig_height = 660
    # orig_width = 1050
    # court_id = row["ID"]
    juris = row['Jurisdiction']
    websites = row['Website']
    type = row['Type']
    # collect_sites.append(juris)
    db_last_modified = row['Datemodified']
    db_last_checked = row['Datelastcheked']
    db_XPATH = row["XPATH"]
    Status = "Error Opening site"
    Last_modified = "No date available"
    FolderName = row["FolderName"]
    current_status = row["Status"]
    # if current_status == "Current" or current_status == "Newly Added":
    #     pass
    # else:
    #     continue
    if type=="PDF":
        try:
            my_header = {'User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/118.0.0.0 Safari/537.36'}
            header = requests.get(websites, headers=my_header )
            if 'Last-Modified' in header.headers:
                Last_modified = header.headers['Last-Modified']
            else:
                try:
                    my_header = {'User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/118.0.0.0 Safari/537.36'}
                    header = requests.get(websites, headers=my_header )
                    if 'Last-Modified' in header.headers:
                        Last_modified = header.headers['Last-Modified']
                        print(Last_modified)
                    if 'Last-Modified' in header.headers:
                        Last_modified = header.headers['last-modified']
                except:
                    collect_data.append([juris, websites, Status, "", db_last_modified, datetime.datetime.now().strftime("%b %d %Y %I:%M %p"), FolderName])
                    continue   
        except:
            try:
                my_header = {'User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/118.0.0.0 Safari/537.36'}
                header = requests.get(websites, headers=my_header )
                if 'Last-Modified' in header.headers:
                    Last_modified = header.headers['Last-Modified']
                    print(Last_modified)
                if 'Last-Modified' in header.headers:
                    Last_modified = header.headers['last-modified']
            except:
                collect_data.append([juris, websites, Status, "", db_last_modified, datetime.datetime.now().strftime("%b %d %Y %I:%M %p"), FolderName])
                continue
        if Last_modified == "No date available":
            collect_data.append([juris, websites, Status, "", db_last_modified, datetime.datetime.now().strftime("%b %d %Y %I:%M %p"), FolderName])
            continue
        else:
            if Last_modified == db_last_modified:
                collect_data.append([juris, websites, "Current", "", Last_modified,  datetime.datetime.now().strftime("%b %d %Y %I:%M %p"), FolderName])
                continue
            else:
                collect_data.append([juris, websites, "With Changes", "", Last_modified,  datetime.datetime.now().strftime("%b %d %Y %I:%M %p"), FolderName])       
                continue

    else:
        try:
            my_header= {'User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/118.0.0.0 Safari/537.36'}
            header = requests.get(websites, headers=my_header )
            if 'Last-Modified' in header.headers:
                Last_modified = header.headers['Last-Modified']
        except:
            pass
        try:
            driver.get(websites)
        except:
            collect_data.append([juris, websites, Status, "", Last_modified,  datetime.datetime.now().strftime("%b %d %Y %I:%M %p"), FolderName])
            continue
        # create folder name
        folder_name = ""
        folder_name = states_folder + FolderName
        try:
            os.makedirs(states_folder + FolderName)
            folder_name = states_folder + FolderName
        except:
            pass
        try:
            text = driver.find_element(By.XPATH, db_XPATH).get_attribute("innerHTML")
            text = "<html>" + text + "</html>"
            print(folder_name)
            if os.path.exists(folder_name + "\\old_" + FolderName + ".html") and os.path.exists(folder_name + "\\new_" + FolderName + ".html"):
                os.remove(folder_name + "\\old_" + FolderName + ".html")
                os.rename(folder_name + "\\new_" + FolderName + ".html", folder_name + "\\old_" + FolderName + ".html")
            if os.path.exists(folder_name + "\\old_" + FolderName + ".html"):
                with open(folder_name + "\\new_" + FolderName + ".html", 'w', encoding="utf-8") as file:
                    file.write(text)
                text1 = open(folder_name + "\\old_" + FolderName + ".html", "r", encoding="utf-8").read()
                text2 = open(folder_name + "\\new_" + FolderName + ".html", "r", encoding="utf-8").read()
                text1 =  re.sub('\s+',' ',text1)
                text2 =  re.sub('\s+',' ',text2)
                text1 = str(clean_output(text1))
                text2 = str(clean_output(text2))
                soup_text1 = BeautifulSoup(text1, 'html.parser')
                a_element = soup_text1.findAll('a')
                for href_elem in a_element:
                    try:
                        if str(href_elem['href']).startswith("#"):
                            href_elem.attrs = {}
                        else:
                            # Remove all attributes except 'href'
                            href_elem.attrs = {'href': href_elem['href']}
                    except:
                        pass

                text1 = str(soup_text1)

                soup_text2 = BeautifulSoup(text2, 'html.parser')
                a_element = soup_text2.findAll('a')

                for href_elem in a_element:
                    
                    try:
                        if str(href_elem['href']).startswith("#"):
                            href_elem.attrs = {}
                        else:
                            # Remove all attributes except 'href'
                            href_elem.attrs = {'href': href_elem['href']}
                    except:
                        pass

                text2 = str(soup_text2)
                dmp = diff_match_patch()
                # Depending on the kind of text you work with, in terms of overall length
                # and complexity, you may want to extend (or here suppress) the
                # time_out feature
                dmp.Diff_Timeout = 0  # or some other value, default is 1.0 seconds
                # All 'diff' jobs start with invoking diff_main()
                diffs = dmp.diff_main(text1, text2)
                # diff_cleanupSemantic() is used to make the diffs array more "human" readable
                dmp.diff_cleanupSemantic(diffs)
                # and if you want the results as some ready to display HTML snippet
                htmlSnippet = dmp.diff_prettyHtml(diffs)
                compare_path = folder_name + "\\" + FolderName + "_" + "Compare.html"
                htmlSnippet = htmlSnippet.replace("&para;<br>", "")

                with open(compare_path, "w", encoding='utf-8') as compare_file:
                    compare_file.write(html.unescape(htmlSnippet.replace("&lt;", "<").replace("&gt;", ">")))
                    compare_file.close()
                if "<mark>" in htmlSnippet:
                    Status = "With Changes"
                else:
                    Status = "Current"
                if Status == "With Changes":
                    collect_sites.append(FolderName)
                    collect_data.append([juris, websites, Status, juris + "_" + "Compare.html", Last_modified,  datetime.datetime.now().strftime("%b %d %Y %I:%M %p"), FolderName])
                else:
                    collect_data.append([juris, websites, Status, '', Last_modified,  datetime.datetime.now().strftime("%b %d %Y %I:%M %p"), FolderName])

            else:
                with open(folder_name + "\\old_" + juris + ".html", 'w', encoding="utf-8") as file:
                    file.write(text)
                collect_data.append([juris, websites, 'Cannot Compare new website added', '', Last_modified,  datetime.datetime.now().strftime("%b %d %Y %I:%M %p"), FolderName])
        except:
            collect_data.append([juris, websites, 'Unable to check website. Please check if the website is available', '', Last_modified,  datetime.datetime.now().strftime("%b %d %Y %I:%M %p"), FolderName])

   
# Create object of ZipFile
current_date_now = datetime.datetime.now()
current_date_now = current_date_now.strftime("%b_%d_%Y_%I_%M_%p")

df = pd.DataFrame(collect_data, columns=['Jurisdiction', 'Rule websites', 'Status', 'Compare file', 'Last modified', 'Last checked', 'FolderName'])
df.to_excel(main_dir + "\\Report\\Report_" + current_date_now +".xlsx", index=False)


with ZipFile(main_dir + '\\ZIP\\Report_'+current_date_now+'.zip', 'w') as zip_object:
# Traverse all files in directory
    for folder_name, sub_folders, file_names in os.walk(main_dir + "\\States"):
        for sites in collect_sites:
            if (os.path.basename(folder_name)) == sites:
                print((os.path.basename(folder_name)))
                for filename in file_names:
                        # Create filepath of files in directory
                        file_path = os.path.join(folder_name, filename)
                        # Add files to zip file
                        print(file_path, os.path.basename(file_path))
                        zip_object.write(file_path, os.path.basename(file_path))
    try:
        zip_object.write(main_dir + "\\Report\\" + "Report_" + current_date_now +".xlsx", "Report_" + current_date_now +".xlsx")
    except Exception as e:
        pass


try:
    driver.quit()
except:
    pass

from sqlalchemy import text
conn_str = ("Driver={SQL Server};"
            "Server=RETDA2SQLD150;"
            "Database=ContentTeamOPEX;"
            "UID=ContentTeamAPP;"
            "PWD=U98=k9M23H&Cgw5MW6tDBgAp9;")
engine_stmt = (
    "mssql+pyodbc://ContentTeamAPP:U98=k9M23H&Cgw5MW6tDBgAp9@RETDA2SQLD150/ContentTeamOPEX?driver=SQL Server")
engine = sqlalchemy.create_engine(engine_stmt)

with engine.connect() as conn1:
    for index, row  in df.iterrows():
        juris =str( row["Jurisdiction"])
        status = str(row["Status"])
        compare_file = str(row["Compare file"])
        date_last_check = str(row["Last checked"])
        date_modified_check = str(row["Last modified"])
        s = "UPDATE PG_Court_Rules SET Status = '%s', Links = '%s', Datemodified = '%s', Datelastcheked = '%s' WHERE Jurisdiction ='%s'" % ("Current", compare_file,  date_modified_check, date_last_check, juris)
        conn1.execute(text(s))
        conn1.commit()


text = ""

try:
    os.remove(main_dir + "\\running.txt")
except:
    pass
try:
    os.remove(main_dir + "\\Data.txt")
except:
    pass

def dispatch(app_name: str):
    try:
        from win32com import client
        app = client.DispatchEx(app_name)
        return app
    except AttributeError:
        try:
            # Remove cache and try again.
            module_list = [m.__name__ for m in sys.modules.values()]
            for module in module_list:
                if re.match(r'win32com\.gen_py\..+', module):
                    del sys.modules[module]
            shutil.rmtree(os.path.join(os.environ.get('LOCALAPPDATA'), 'Temp', 'gen_py'))
            from win32com import client
            app = client.DispatchEx(app_name)
            return app
        except:
            pass

outlook_app = dispatch('Outlook.Application')


def send_mail(send_to, attachment_path,table_df , date_report):
    outlook = dispatch('Outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = send_to
    mail.Subject = "SiteScout Update Report " + date_report

    try:
        statinfo = os.stat(attachment_path)
        if statinfo.st_size <= 26214400:
            mail.Attachments.Add(attachment_path)
        else:
            pass
    except:
        pass
    data = table_df[table_df['Status'] != 'Current']
    data.reset_index(drop=True, inplace=True)
    data.index += 1
    Table = data  # Table variable goes into the space {1} in the HTML Code Below
    html = """\
                <html>
                <head> 
                </head>
                <body>
                        Hi Everyone,
                    <br>
                    <br>
                        You may now download and check the latest report in the <a href ="https://lngdayappp007.legal.regn.net/PGCourt/">https://lngdayappp007.legal.regn.net/PGCourt/</a>.
                    <br>
                    <br>
                        The table below shows which states have had revisions to their websites.
                    <br>
                    <br>
                        {0}
                    <br>
                    <br>
                        Thank You.
                    <br>
                    <br>
                </body>
                </html>
                """.format(Table.to_html())
    
    mail.HTMLbody = html
    mail.Send()

df = pd.read_excel(main_dir + "\\Report\\Report_" + current_date_now +".xlsx")
with open(main_dir + "\\email.txt", "r", encoding="utf-8") as email_txt:
    list_email = str(email_txt.read())
send_mail(list_email, main_dir + '\\ZIP\\Report_'+current_date_now+'.zip', df, "Report_" + current_date_now +".xlsx")
# send_mail("anthony.abordo@reedelsevier.com", "anthony.abordo@reedelsevier.com", "Court Website Alert", 'You may now download and check the latest report in the http://lngdayappp007.lexisnexis.com:90/.', 'lngdayappp007.legal.regn.net', 25, main_dir + '\\ZIP\\Report_'+ current_date_now + '.zip', df)

sys.exit()
