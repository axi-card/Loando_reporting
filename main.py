from download_files import Downloads
from mailing import Mailing
from report_preparing import Reporting
import time
import datetime
import glob
import pandas as pd
import os
import sys

def download_LOANDO(code):

    downloads = Downloads()

    downloads.download_Proposals_DC1(code)

    downloads.download_Proposals_DC2(code)

    downloads.concat_Proposals()

    downloads.download_Credit_Cards_DC1(code)

    downloads.download_Credit_Cards_DC2(code)

    downloads.concat_Credit_Cards()

    downloads.download_Reports_Cards()

    downloads.download_Processing()

    downloads.download_Raport_do_CC()

def prepare_report_LOANDO(code):

    report = Reporting()

    report.card_proposals_preparation()

    report.credit_cards_preparation(code)

    report.report_to_file(code)




def remove_old():

    try:
        os.remove("DOWNLOADS/Compare/old.xlsx")
    except:
        print("COULD NOT FIND OLD.XLSX")

def change_new_to_old(starttime):

    try:
        os.rename("DOWNLOADS/Compare/new.xlsx", "DOWNLOADS/Compare/old.xlsx")
    except Exception as e:

        elapsedtime = time.time() - starttime

        email.send_error_message(elapsedtime, e)

        print("COULD NOT FIND NEW.XLSX\nPYTHON SCRIP HAS TO STOP")
        sys.exit()



if __name__ == "__main__":

    if datetime.datetime.today().weekday() in [5,6]:
        sys.exit()

    starttime = time.time()

    email = Mailing()



    codes = ['48060009003006','Odnaol','Ondaol']



    for code in codes:

        try:

            download_LOANDO(code)

        except Exception as e:

            print("Report could not be created due to downloading failure")

            elapsedtime = time.time() - starttime

            #email.send_error_message(elapsedtime,e)

            sys.exit()



        try:

            prepare_report_LOANDO(code)

        except Exception as e:

            print("Report could not be created due to report preparation failure")

            elapsedtime = time.time() - starttime

            #email.send_error_message(elapsedtime,e)

            sys.exit()


    path = "DOWNLOADS/Concat Reports"

    all_files = glob.glob(path + "/*.xlsx")
    df_list = []

    for file in all_files:
        df = pd.read_excel(file)
        df_list.append(df)

    main_df = pd.concat(df_list, axis=0, ignore_index=True)


    main_df = main_df.sort_values(by=['Proposal date'], ascending=False)

    main_df = main_df.drop_duplicates(subset=['PESEL'])



            # The last stage of reporting

    remove_old()

    change_new_to_old(starttime)

    report = Reporting()

    report.save_report(main_df,"DOWNLOADS/Compare/","new",".xlsx")

    report.save_report(main_df,"J:/Public/tymczasowe/Raporty LOANDO/Full/",report.get_report_name(),".csv")

    report.compare_files()

    elapsedtime = time.time() - starttime

    #email.send_success_message(elapsedtime)

