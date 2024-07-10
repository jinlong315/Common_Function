import shutil
import pyodbc
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.header import Header
from email.utils import formataddr
import logging
from pathlib import Path
import math
import sqlalchemy
import pandas as pd
import base64
import datetime
import pymysql

# create connection to My SQL
mysql_user = "root"
mysql_password = "hjl28041108"
mysql_host = "localhost"
mysql_port = 3306
mysql_database = "cleaned_data"

# create connection to EPAPER_GC
epaper_sqlserver_user = "epaper_sca"
epaper_sqlserver_password = "9xp9WJrA"
epaper_sqlserver_host = "WS007238"
epaper_sqlserver_port = 1433
epaper_sqlserver_database = "EPAPER_SCA_GC"
epaper_sqlserver_charset = "utf8"

# create connection to sca_digital
sca_digital_sqlserver_user = "SCA_Admin"
sca_digital_sqlserver_password = "JAdmin!2309!"
sca_digital_sqlserver_host = "WS007238"
sca_digital_sqlserver_port = 1433
sca_digital_sqlserver_database = "SCA_Digital"
sca_digital_sqlserver_charset = "utf8"

# create connection to sca_digital_dev
sca_dev_sqlserver_user = "SCA_Admin"
sca_dev_sqlserver_password = "JAdmin!2309!"
sca_dev_sqlserver_host = "WS007238"
sca_dev_sqlserver_port = 1433
sca_dev_sqlserver_database = "SCA_Digital_Dev"
sca_dev_sqlserver_charset = "utf8"

# create connection to MES SQL Server
mes_sqlserver_user = "factetk"
mes_sqlserver_password = "factetk"
mes_sqlserver_host = "sqlmep22.schaeffler.com"
mes_sqlserver_port = 1433
mes_sqlserver_database = "MEP22"
mes_sqlserver_charset = "utf8"

# create connection to Azure SQL Server for S7 dashboard
s7_azure_sql_server = "consqlp0013.database.windows.net"
s7_azure_sql_user = "data_reader"
s7_azure_sql_password = "Jv?yDNQK&g!)U7Q#kxY="
s7_azure_sql_database = "Machine_Dashboards_BHL_P"
s7_azure_sql_port = 1433
s7_azure_sql_charset = "utf8"

# create connection to HR SQL Server
sap_hr_sql_server = "DE010781.de.ina.com"
sap_hr_sql_user = "OLForm_UserData_Read"
sap_hr_sql_password = "!OLFudread"
sap_hr_sql_database = "ITSGInfoDB"
sap_hr_sql_port = 1433
sap_hr_sql_charset = "utf8"

# create connection to Azure SQL Server for SDP
azure_sdp_sql_server = "sdp-s-d-sqls.database.windows.net"
azure_sdp_sql_user = "p3pq@schaeffler.com"
azure_sdp_sql_password = "Pq123456"
azure_sdp_sql_database = "sdp-s-gcscapbi"
azure_sdp_sql_port = 1433
azure_sdp_sql_charset = "utf8"

# create connection to SQL Server (SCA_Digital / SCA_Digital_Dev / EPAPER_SCA_GC)
common_user = "SCA_IE_Dig_R"
common_password = "!25@Sca"

# server account
server_pc = "P01226819"
server_user = "p3pq"
server_password = "Pq123456"
server_email = "P3pq@schaeffler.com"

# define one dictionary to save email address
dict_email = {"HuangJinlong": "huangjon@schaeffler.com",
              "QuXiaoli": "quxol@schaeffler.com",
              "ZhuBaoli": "zhubol@schaeffler.com",
              "ShengYan": "shengyn@schaeffler.com",
              "ZhaiXiaotong": "zhaixia@schaeffler.com",
              "ZhuZijing": "zhuzji@schaeffler.com",
              "ZhaoJing": "zhaojn4@schaeffler.com",
              "BaiHuasong": "baihas@schaeffler.com",
              "ZhangYuanyuan": "zhanyyn@schaeffler.com",
              "ZhangMengfei": "zhangmgf@schaeffler.com",
              "ChengChun": "chengchu@schaeffler.com",
              "ShiKaiteng": "shikit@schaeffler.com",
              "HuJun": "hujn5@schaeffler.com",
              "MinCui": "minci@schaeffler.com",
              "DingLing": "dinglig@schaeffler.com",
              "LiFengjia": "lifej@schaeffler.com",
              "IE_PM": "OR-Taicang-Plant3-IE-and-PM@schaeffler.com",
              "CaoYichen": "caoych@schaeffler.com",
              "SuYun": "suyun@schaeffler.com",
              "YaoXiaowei": "yaoxow@schaeffler.com",
              "ZhouPengrui": "zhoupgr@schaeffler.com",
              "ZhuYihong": "zhuyih@schaeffler.com",
              "Xionghao": "xiongho@schaeffler.com"
              }

# define list to save year
list_year_yyyy = [str(i) for i in range(2023, 2124, 1)]

# define list to save month
list_english_month = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
list_numerical_month = [str(i).rjust(2, "0") for i in range(1, 13, 1)]
dict_english_numerical_month = dict(zip(list_english_month, list_numerical_month))

# define list to save 26 english letters
list_letters_upper = [chr(i) for i in range(65, 91, 1)]
list_letters_lower = [chr(i).lower() for i in range(65, 91, 1)]

# define list to save year month
list_year_month = []
for year in list_year_yyyy:
    for month in list_numerical_month:
        year_month = "".join([year, month])
        list_year_month.append(year_month)

# suffix of files
excel_suffix = [".xlsm", ".xlsx", ".xlsb", ".xls"]
picture_suffix = [".png", ".jpeg", "", ".jpg", ".gif", ".svg"]
pdf_suffix = [".pdf"]


# define one object to send email
class SendEmail:
    """
    One object to send email
    """

    def __init__(self, sender_name, sender_address, receiver, cc, subject, content):
        """
        :param sender_name:  str | sender_name,
        :param sender_address:  str | sender_address,
        :param receiver: list | email address of receiver,
        :param cc: list | email address of Cc,
        :param subject: str | email title
        :param content: str | content of email, can only be normal text
        """
        self.sender_name = sender_name
        self.sender_address = sender_address
        self.receiver = receiver
        self.cc = cc
        self.subject = subject
        self.content = content

    def send_email_with_text(self):
        # create connection to Email Serve
        email_server = smtplib.SMTP(host="mail-de-hza.schaeffler.com", port=25)

        # create email object
        msg = MIMEMultipart()

        # create subject
        title = Header(s=self.subject, charset="utf-8").encode()
        msg["Subject"] = title

        # set sender
        msg["From"] = formataddr((self.sender_name, self.sender_address))

        # set receiver
        msg["To"] = ",".join(self.receiver)

        # set Cc
        msg["Cc"] = ",".join(self.cc)

        # add content
        text = MIMEText(_text=self.content, _subtype="plain", _charset="utf-8")
        msg.attach(text)

        # extend receiver list
        to_list = msg["To"].split(",")
        cc_list = msg["Cc"].split(",")
        to_list.extend(cc_list)

        # send email
        email_server.sendmail(from_addr=msg["From"], to_addrs=to_list, msg=msg.as_string())
        email_server.quit()

    def send_email_with_html(self):
        # create connection to Email Serve
        email_server = smtplib.SMTP(host="mail-de-hza.schaeffler.com", port=25)

        # create email object
        msg = MIMEMultipart()

        # create subject
        title = Header(s=self.subject, charset="utf-8").encode()
        msg["Subject"] = title

        # set sender
        msg["From"] = formataddr((self.sender_name, self.sender_address))

        # set receiver
        msg["To"] = ",".join(self.receiver)

        # set Cc
        msg["Cc"] = ",".join(self.cc)

        # add content
        html = MIMEText(_text=self.content, _subtype="html", _charset="utf-8")
        msg.attach(html)

        # extend receiver list
        to_list = msg["To"].split(",")
        cc_list = msg["Cc"].split(",")
        to_list.extend(cc_list)

        # send email
        email_server.sendmail(from_addr=msg["From"], to_addrs=to_list, msg=msg.as_string())
        email_server.quit()


class Logger:
    """
    define one logger for common usage
    """

    def __init__(self, level, file_name):
        """
        :param level: logging level
        :param file_name: absolute directory of log file
        """

        self.level = level
        self.file_name = file_name

    def basic_configuration(self):
        """
        :return: basic configuration for logging
        """
        return logging.basicConfig(level=self.level,
                                   filename=self.file_name,
                                   filemode="w",
                                   format="%(asctime)s - %(filename)s[line:%(lineno)d] - %(levelname)s: %(message)s")


class MSSQL:
    """
    define one object to get connection with MS SQL Server
    """

    def __init__(self, server, user, password, database):
        """
        :param server: str | host name of MS SQL Server
        :param user: str | user to log in MS SQL Server
        :param password: str | password to log in MS SQL Server
        :param database: str | database name in MS SQL Server
        """
        self.server = server
        self.user = user
        self.database = database
        self.password = password

    def pyodbc_connection(self):
        """
        :return: pyodbc connection
        """
        if not self.database:
            raise (NameError, "Incorrect configuration of MS SQL Server !")

        # create connection to SQL Server
        connection_string = f'DRIVER={{SQL Server}};SERVER={self.server};DATABASE={self.database};UID={self.user};PWD={self.password}'
        try:
            pyodbc_con = pyodbc.connect(
                connection_string,
                fast_executemany=True
            )

            return pyodbc_con

        except Exception:
            raise Exception("Failed to connect to MS SQL Server !")

    def sqlalchemy_connection(self):
        """
        :return: SQLAlchemy + pyodbc connection
        """

        # create connection to SQL Server
        try:
            engine = sqlalchemy.create_engine(
                'mssql+pyodbc://{}:{}@{}/{}?driver=ODBC+Driver+17+for+SQL+Server'.format(self.user,
                                                                                         self.password,
                                                                                         self.server,
                                                                                         self.database),
                fast_executemany=True)

            sqlalchemy_con = engine.connect()

            return sqlalchemy_con

        except Exception:
            raise Exception("Failed to connect to MS SQL Server !")

    def add_table_property(self, table_name, table_desc):
        """
        :param table_name: table name in MS SQL Server
        :return:
        """

        # create sql string
        sql = f"""
        EXEC sp_addextendedproperty   
                @name = N'MS_Description',
                @value = N'{table_desc}',
                @level0type = N'Schema',
                @level0name = N'dbo',
                @level1type = N'Table',
                @level1name = N'{table_name}';
        """

        # get cursor
        con = self.pyodbc_connection()
        cursor = con.cursor()

        # execute sql string
        cursor.execute(sql)

        # submit and close
        con.commit()
        con.close()

    def update_table_property(self, table_name, table_desc):
        """
        :param table_name: table name in MS SQL Server
        :return:
        """

        # create sql string
        sql = f"""
        EXEC sp_updateextendedproperty   
                @name = N'MS_Description',
                @value = N'{table_desc}',
                @level0type = N'Schema',
                @level0name = N'dbo',
                @level1type = N'Table',
                @level1name = N'{table_name}';
        """

        # get cursor
        con = self.pyodbc_connection()
        cursor = con.cursor()

        # execute sql string
        cursor.execute(sql)

        # submit and close
        con.commit()
        con.close()

    def execute_sql_query(self, sql):
        """
        :param sql: sql query string
        :return: None
        """

        # get cursor
        con = self.pyodbc_connection()
        cursor = con.cursor()

        # execute sql query
        cursor.execute(sql)

        # commit and close
        con.commit()
        con.close()


class MySQL:
    """
    define one object to get connection with MS SQL Server
    """

    def __init__(self, server, user, password, database):
        """
        :param server: str | host name of MS SQL Server
        :param user: str | user to log in MS SQL Server
        :param password: str | password to log in MS SQL Server
        :param database: str | database name in MS SQL Server
        """
        self.server = server
        self.user = user
        self.database = database
        self.password = password

    def sqlalchemy_connection(self):
        """
        :return: Connect to MySQL
        """

        # create connection to MySQL
        con = sqlalchemy.create_engine(
            f"mysql+pymysql://{mysql_user}:{mysql_password}@{mysql_host}:{mysql_port}/{mysql_database}")

        return con


class PDFData:
    """
    Deal with PDF data
    """

    def __init__(self, dir_pdf):
        """
        :param dir_pdf: absolute directory of PDF file
        """

        self.dir_pdf = dir_pdf

    def convert_to_base64(self):
        """
        :return: DataFrame with spil string
        """

        # open PDF file
        with open(file=self.dir_pdf, mode="rb") as file:
            content = file.read()

            # convert to base64
            base64_content = base64.b64encode(content).decode("utf-8")

            # split string
            str_length = len(base64_content)
            step = 100
            loop_count = math.ceil(str_length / step)

            dict_split = {}

            for i in range(0, loop_count, 1):
                str_split = base64_content[i * step: (i + 1) * step]
                dict_split[i] = str_split

            # save data into pandas
            df_pdf = pd.DataFrame.from_dict(data=dict_split, orient="index", columns=["base64"], dtype="str")

            # reset index
            df_pdf.index.name = "base64_order"
            df_pdf.reset_index(drop=False, inplace=True)

            # create DataFrame
            file_name = Path(self.dir_pdf).name
            df_file = pd.DataFrame([[file_name]], index=[i for i in range(0, loop_count, 1)], columns=["file_name"])

            # concatenate DataFrame
            df_merged = pd.concat([df_file, df_pdf], axis=1)

            return df_merged


class GetFileList:
    """
    get list of directory for specified type of file in the given folder
    """

    def __init__(self, folder_directory, suffix):
        """
        :param folder_directory: str | absolute directory of raw data
        :param suffix: list | list of suffixes
        """
        self.folder_directory = folder_directory
        self.suffix = suffix

    def get_abs_dir(self):
        """
        :return: list, which saved the full directory of files which end with specified file type
        """
        list_file = []
        for i in Path(self.folder_directory).iterdir():
            if (Path(i).is_file()) and (Path(i).suffix.lower() in self.suffix) and ("$" not in Path(i).stem):
                list_file.append(Path(i))
        # return list_xlsm
        return list_file

    def get_ANA_report(self):
        """
        :return: list of absolute directory of raw data
        """

        # get "PCR" path
        path_pcr = Path(self.folder_directory)
        list_ana_report = []

        for i in path_pcr.iterdir():
            # get "year" path
            if (Path(i).is_dir()) and (Path(i).stem in list_year_yyyy):
                path_year = Path(i)
                year = Path(path_year).stem

                for j in path_year.iterdir():
                    # get "month" path
                    if (Path(j).is_dir()) and (Path(j).stem[-3:] in list_english_month):
                        path_month = Path(j)

                        for k in path_month.iterdir():
                            # get "ANA official file" path
                            if (Path(k).is_dir()) and ("ana_official_file" in Path(k).stem.lower()):
                                path_ana = Path(k)

                                for m in path_ana.iterdir():
                                    # get "submit" path
                                    if (Path(m).is_dir()) and ("submit" in Path(m).stem.lower()):
                                        path_submit = Path(m)

                                        for n in path_submit.iterdir():
                                            # get "ANA report file" path
                                            if (Path(n).is_file()) and (Path(n).suffix.lower() in excel_suffix) and (
                                                    "ana_kzb" in Path(n).stem.lower()) and (
                                                    "$" not in Path(n).stem):
                                                path_file = Path(n)

                                                list_ana_report.append(path_file)

        return list_ana_report

    def get_assy_daily_credit(self):
        """
        :return: list of absolute directory of EXCEL file
        """

        # define variable to save absolute directory of raw data
        list_excel = []

        # filter EXCEL file which meet specified conditions
        path_daily_mirror = Path(self.folder_directory)
        for year in path_daily_mirror.iterdir():

            if (Path(year).is_dir()) and (Path(year).stem in list_year_yyyy):
                path_year = Path(year)

                for month in path_year.iterdir():

                    if (Path(month).is_dir()) and (Path(month).stem in list_year_month):
                        path_month = Path(month)

                        for file in path_month.iterdir():

                            if (Path(file).is_file()) and (Path(file).suffix.lower() in self.suffix) and (
                                    "Credit_Auto".lower() in Path(file).stem.lower()) and ("$" not in Path(file).stem):
                                path_file = Path(file)
                                list_excel.append(path_file)

        return list_excel

    def get_muc_daily_credit(self):
        """
        :return: list of absolute directory of EXCEL file
        """

        # define variable to save absolute directory of raw data
        list_excel = []

        # filter EXCEL file which meet specified conditions
        path_daily_mirror = Path(self.folder_directory)
        for year in path_daily_mirror.iterdir():

            if (Path(year).is_dir()) and (Path(year).stem in list_year_yyyy):
                path_year = Path(year)

                for file in path_year.iterdir():

                    if (Path(file).is_file()) and (Path(file).suffix.lower() in self.suffix) and (
                            "MUC Daily Credit".lower() in Path(file).stem.lower()) and ("$" not in Path(file).stem):
                        path_file = Path(file)
                        list_excel.append(path_file)

        return list_excel

    def get_zprl(self):
        """
        :return: list of absolute directory for raw data
        """

        path_0 = Path(self.folder_directory)

        list_zprl = []
        for file in path_0.iterdir():
            if (Path(file).is_file()) and (Path(file).suffix.lower() in excel_suffix) and ("$" not in Path(file).stem):
                path_file = Path(file)
                list_zprl.append(path_file)

        return list_zprl

    def get_delivery_auto(self):
        """
        :return: list of absolute directory of EXCEL file
        """

        # define variable to save absolute directory of raw data
        list_excel = []

        # filter EXCEL file which meet specified conditions
        path_daily_mirror = Path(self.folder_directory)
        for year in path_daily_mirror.iterdir():

            if (Path(year).is_dir()) and (Path(year).stem in list_year_yyyy):
                path_year = Path(year)

                for month in path_year.iterdir():

                    if (Path(month).is_dir()) and (Path(month).stem in list_year_month):
                        path_month = Path(month)

                        for file in path_month.iterdir():

                            if (Path(file).is_file()) and (Path(file).suffix.lower() in self.suffix) and (
                                    "Delivery_Auto".lower() == Path(file).stem.lower()) and (
                                    "$" not in Path(file).stem):
                                path_file = Path(file)
                                list_excel.append(path_file)

        return list_excel
