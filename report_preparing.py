import pandas as pd
pd.set_option('display.max_columns',500)
from numpy import where
import sys
import ezodf
import os
import datetime

class Reporting:

    def __init__(self):
        self.template = pd.DataFrame(columns=['Proposal date', 'Decision date', 'Post date', 'Activation date', 'PESEL','Customer', 'Phone', 'Limit', 'Status', 'Substatus', 'Comments', 'Comment date'])


    def card_proposals_preparation(self):

        self.card_proposals_df = pd.read_excel("DOWNLOADS/Proposals/gvProposals.xls")


        self.card_proposals_df['Stauts'] = self.card_proposals_df['Status'].str.strip()


        self.card_proposals_df =  self.card_proposals_df[-self.card_proposals_df["Status"].isin(["Approved"])]

        self.card_proposals_df = self.card_proposals_df.rename(columns={"Date of Proposal ":"Proposal date", "Date Credit Dept. ":"Decision date", "Names":"Customer",})

        self.card_proposals_df['Post date'] = ""

        self.card_proposals_df['Activation date'] = ""

        self.card_proposals_df['Substatus'] = ""

        self.card_proposals_df['Comments'] = ""

        self.card_proposals_df['Comment date'] = ""

        # Order


        self.card_proposals_df = self.card_proposals_df[['Proposal date', 'Decision date', 'Post date', 'Activation date', 'PESEL','Customer', 'Phone', 'Limit', 'Status', 'Substatus',
                                                         'Comments', 'Comment date']]

        self.card_proposals_df = self.card_proposals_df.sort_values(by=['Proposal date'], ascending=False)

        self.card_proposals_df = self.card_proposals_df.drop_duplicates(subset=['PESEL'])



    def credit_cards_preparation(self, code):

        # Get appropriate files into pandas dataframes

        self.credit_cards_df = pd.read_excel("DOWNLOADS/CreditCards/ASPxGridViewCreditCards.xls")

        self.reports_cards_df = pd.read_excel("DOWNLOADS/ReportsCards/ASPxGridViewCards.xls")

        self.processing_prep()

        self.raport_do_cc_prep()

        self.card_proposals_temp_df = pd.read_excel("DOWNLOADS/Proposals/gvProposals.xls")


        self.credit_cards_df['Status'] = self.credit_cards_df['Status'].str.strip()


        self.credit_cards_df = self.credit_cards_df[self.credit_cards_df['Status'].isin(['With signed contract', 'Approved'])]



        #PESEL matching for Credit Cards from both Reports Cards and Proposals

        self.credit_cards_df = self.credit_cards_df.merge(self.reports_cards_df[['PESEL', 'Date of Activation', 'CID']], on='CID', how='left')

        self.credit_cards_df = self.credit_cards_df.merge(self.card_proposals_temp_df[['Phone', 'PESEL']], on='Phone', how='left')

        self.credit_cards_df['PESEL'] = where(self.credit_cards_df['PESEL_x'].isnull(), self.credit_cards_df['PESEL_y'],self.credit_cards_df['PESEL_x'])

        #Merge with Raport do CC


        self.credit_cards_df = self.credit_cards_df.merge(self.raport_do_cc_df[['CID', 'Post date', 'Comments','Date of return']], on='CID', how='left')

        self.credit_cards_df['PESEL'] = self.credit_cards_df['PESEL'].fillna("0")

        self.credit_cards_df['PESEL'] = self.credit_cards_df['PESEL'].astype('int64')



        #self.credit_cards_df['Date of return'] = where(self.credit_cards_df['Comments'].isnull(),"", self.credit_cards_df['Date of return'])

        self.credit_cards_df['Substatus'] = ""

        self.credit_cards_df['Comment date'] = self.credit_cards_df['Date of return']



        self.credit_cards_df = self.credit_cards_df.rename(columns={"Date of Proposal": "Proposal date", "Approval Date": "Decision date", "Date of Activation": "Activation date"})

        self.credit_cards_df = self.credit_cards_df[['Proposal date', 'Decision date', 'Post date', 'Activation date', 'PESEL', 'Customer', 'Phone', 'Limit', 'Status',
                                                     'Substatus', 'Comments', "Comment date",'CID']]


        #Concatenate Proposals and Credit Cards

        self.concatenated_df = pd.concat([self.card_proposals_df, self.credit_cards_df], ignore_index=True, sort=False)




        #Merge with processing

        self.concatenated_df = self.concatenated_df.drop(columns=['Substatus','CID'])

        self.concatenated_df = self.concatenated_df.merge(self.processing_df[['Komentarz', 'PESEL', 'Substatus', 'Data komentarza']], on='PESEL', how='left')


        # Combine Komentarz and Comments columns

        self.concatenated_df['Comments'] = self.concatenated_df['Comments'].fillna("")

        self.concatenated_df['Komentarz'] = self.concatenated_df['Komentarz'].fillna("")

        self.concatenated_df['Komentarz'] = where(self.concatenated_df['Komentarz'] == "", self.concatenated_df['Comments'],
                                                  self.concatenated_df['Komentarz'] + " " + self.concatenated_df['Comments'])

        self.concatenated_df['Comments'] = self.concatenated_df['Komentarz'].str.strip()


        # Change to datetime format

        self.concatenated_df['Post date'] = self.concatenated_df['Post date'].astype('datetime64[ns]')

        self.concatenated_df['Comment date'] = self.concatenated_df['Comment date'].astype('datetime64[ns]')


        self.concatenated_df['Comment date'] = where(self.concatenated_df['Comment date'].isnull(),self.concatenated_df['Data komentarza'],self.concatenated_df['Comment date'])

        # Order, limit and other manipulations

        self.concatenated_df = self.concatenated_df.rename(columns={"Date of Proposal":"Proposal date","Approval Date": "Decision date", "Date of Activation": "Activation date"})

        self.concatenated_df['Comments'] = self.concatenated_df['Comments'].str.replace("\n"," ")


        self.concatenated_df['additional'] = where(self.concatenated_df['Comments'].str.contains("ODSTĄPIENIE"),self.concatenated_df['Comments'],None)

        self.concatenated_df['additional_date'] = where(self.concatenated_df['Comments'].str.contains("ODSTĄPIENIE"), self.concatenated_df['Comment date'], None)

        self.concatenated_df['Comments'] = where(self.concatenated_df['Status'] == 'With signed contract', None,self.concatenated_df['Comments'])


        self.concatenated_df['Comment date'] = where(self.concatenated_df['Comments'].isnull(), None, self.concatenated_df['Comment date'])

        self.concatenated_df['Comments'] = where(self.concatenated_df['Comments'].isnull(), self.concatenated_df['additional'], self.concatenated_df['Comments'])

        self.concatenated_df['Comment date'] = where(self.concatenated_df['Comment date'].isnull(), self.concatenated_df['additional_date'], self.concatenated_df['Comment date'])


        self.concatenated_df['Comment date'] =  self.concatenated_df['Comment date'].astype('datetime64[ns]')


        self.concatenated_df = self.concatenated_df[['Proposal date', 'Decision date', 'Post date', 'Activation date', 'PESEL', 'Customer', 'Phone', 'Limit', 'Status',
                                                     'Substatus', 'Comments', "Comment date"]]



    def processing_prep(self):

        try:
            self.processing_df = pd.read_excel("DOWNLOADS/Processing/processing LOANDO.xlsx")

            if not pd.Series(['Komentarz', 'PESEL', 'Substatus', 'Data komentarza']).isin(self.processing_df.columns).all():
                raise Exception

        except FileNotFoundError:
            msg = "There is no DOWNLOADS/Processing/processing.xlsx directory.\nReport preparation has stopped."
            #self.emsg.send_critical_message(msg)
            print(msg)
            sys.exit()

        except Exception:
            msg = "processing.xlsx file does not contain all necessary columns.\nReport preparation has stopped."
            #self.emsg.send_critical_message(msg)
            print(msg)
            sys.exit()

    def raport_do_cc_prep(self):

        try:
            self.raport_do_cc_df = self.read_ods("DOWNLOADS/Raport do CC/Raport do CC NEW.ods", 0)

            if not pd.Series(['CID','Post date','Comments']).isin(self.raport_do_cc_df.columns).all():
                raise Exception

        except FileNotFoundError:
            msg = "There is no DOWNLOADS/Raport do CC/Raport do CC NEW.ods directory.\nReport preparation has stopped."
            #self.emsg.send_critical_message(msg)
            print(msg)
            sys.exit()

        except Exception:
            msg = "Raport do CC NEW.ods file does not contain all necessary columns.\nReport preparation has stopped."
            #self.emsg.send_critical_message(msg)
            print(msg)
            sys.exit()

        self.raport_do_cc_df['CID'] = self.raport_do_cc_df['CID'].fillna(0)

        self.raport_do_cc_df['CID'] = self.raport_do_cc_df['CID'].astype('int')


    def read_ods(self,filename, sheet_no=0, header=0):
        tab = ezodf.opendoc(filename=filename).sheets[sheet_no]
        return pd.DataFrame({col[header].value: [x.value for x in col[header + 1:]] for col in tab.columns()})


    def report_to_file(self, code):


        try:
            os.remove("DOWNLOADS/Concat Reports/{}.xlsx".format(code))

        except:

            pass

        finally:

            self.concatenated_df.to_excel("DOWNLOADS/Concat Reports/{}.xlsx".format(code), index=False)


    def compare_files(self):
        """Compare both files and return differences report"""

        old_df = pd.read_excel("DOWNLOADS/Compare/old.xlsx")

        new_df = pd.read_excel("DOWNLOADS/Compare/new.xlsx")

        new_df = new_df.merge(old_df[['Status','Post date','Comments','PESEL']], on="PESEL", how='left', suffixes=('', '_previous',))


        # Difference zone

        new_df['Post date'] = new_df['Post date'].fillna("")

        new_df['Post date_previous'] = new_df['Post date_previous'].fillna("")

        new_df['Comments'] = new_df['Comments'].fillna("")

        new_df['Comments_previous'] = new_df['Comments_previous'].fillna("")

        new_df['Difference'] = where((new_df['Status'] != new_df['Status_previous']) | (new_df['Post date'] != new_df['Post date_previous']) | (new_df['Comments'] != new_df['Comments_previous']), "Yes", "No")

        new_df = new_df[new_df['Difference'] == "Yes"]

        new_df = new_df.drop(columns='Difference')

        del new_df['Comments_previous']

        del new_df['Post date_previous']

        new_df['Substatus'] = where(new_df['Status'] != "Processing", "", new_df['Substatus'])

        # in case if with signed contract was erased due to lack of id

        new_df['Status_previous'] = where(new_df['Status'] == 'With signed contract', 'Approved', new_df['Status_previous'])

        new_df.to_excel("OUTPUT/{0}.xlsx".format(self.get_report_name()), index=False)

        new_df.to_excel("J:/Public/tymczasowe/Raporty LOANDO/Changes/{0}.xlsx".format(self.get_report_name()), index=False)




    def get_report_name(self):

        n = datetime.datetime.now()

        report_name = "LOANDO " + str(n.year) + "_"

        if len(str(n.month)) == 1:
            report_name += "0" + str(n.month) + "_"
        else:
            report_name += str(n.month) + "_"


        if len(str(n.day)) == 1:
            report_name += "0" + str(n.day) + "_"
        else:
            report_name += str(n.day) + "_"


        if len(str(n.hour)) == 1:
            report_name += "0" + str(n.hour) + "_"
        else:
            report_name += str(n.hour) + "_"


        if len(str(n.minute)) == 1:
            report_name += "0" + str(n.minute)
        else:
            report_name += str(n.minute)

        return report_name

    def save_report(self,df,location,report_name,extension):

        if extension == ".xlsx":

            df.to_excel(location + report_name + extension, index=False)

        elif extension == ".csv":

            df.to_csv(location + report_name + extension, index=False, encoding='windows-1250')