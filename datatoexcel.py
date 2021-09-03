# Code written by Steven Au
# Purpose: Converts all subdirectories of CSV and TSV files into one Excel file.
# Required modules installed from PIP: pandas, xlsxwriter
import os

import pandas as pd


class DataToExcel:
    """ Converts all subdirectory data into an unified Excel file."""
    
    def __init__(self):
        """
            self.partners is populated within the script due to events where a certain subdirectory is to be omitted. 
        """
        self.source = None
        self.partners = self.the_partners()  # Subdirectories (folder names) must be prepopulated before executing the script.
        self.path_list = []
        self.df_list = []
        self.aggregated_data = None
    
    def the_partners(self):
        """ Update partners here.
            The purpose of using the dictionary is for scalable subdirectories with varying number of skipped heading rows.
            Variables:
                    partner_dictionary = {STRING - parner name (Must match the folder name): INTEGER - skip heading rows (Enter 0, otherwise any integer greater than 0)}
                                    SAMPLE:   {"partnerName": 1,
                                                "folder2": 0,
                                                "OtHeR", 5
                                             }
            Returns:
                    partner_dictionary: dictionary data per the subdirectory name and the number of rows skipped.
        """
        # PartnerName (Has to match the folder name), skipHeadingRows
        partner_dictionary = {
            "": 0
        }
        return partner_dictionary
    
    def set_source(self):
        """ User input of the parent "root" directory path to get the corresponding partner subdirectory files.
            To add a path, simply copy and paste or type the main directory.
            Example: Parent directory is, for Windows, where the parent directory: +, subdirectory: *, and file: ^.
                +C:\downloads
                |          |
                *folder1   *partnerName
                |          |
                ^file.csv  | ^ partnerfile.tsv  # Please have only one filetype per folder if the data's structure differs (Such as headings, columns, etc.)
                           | ^ Partner2.csv
            What you need to enter is the main parent directory in this example.
                    Enter folder source path: C:\downloads
        """
        self.source = input("Enter folder source path: ")
    
    # Aggregators
    def data_gather(self):
        """ Gather all data from sosurce directory folder and aggregate them into a single list."""
        if self.source is None:  # Initiate the source if it is not yet set.
            self.set_source()
            
        for root, dirs, files in os.walk(self.source):
            for name in files:
                if name.endswith('.csv') or name.endswith('.tsv'):  # Further branching or additional conditions can be added to scale up all file types.
                    self.path_list.append(os.path.join(root, name))
    
    def frame_gen(self, file_list, frame_type, skipping_rows):
        """ Generate individual pandas dataframes for each file stored in the list."""
        if self.source is None:  # Prevent execution if there are no files in the list.
            return
        
        for each_file in file_list:
            if each_file.endswith('.csv'):
                df = pd.read_csv(each_file, skiprows=skipping_rows)  
            elif i.endswith('.tsv'):
                df = pd.read_csv(each_file, sep='\t', skiprows=skipping_rows)  # TSV files uses a tab as the seperator. 
            else:  # Note: Due to the aggregation from the data_gather method, please add additional elif branching as appropriate to scale. 
                continue
            
            df.insert(0, 'filename', each_file) # The filename is used as the reference association
            frame_type.append(df)  # Each dataframe will be added to an alternate list
                
    def search_crit(self, partner_name, skipping_rows):
        """ Criteria for search partner dataframes.
            The process uses the partner names as the search criteria.
            Returns:
                    search_fr: the searched criteria dataframes.
        """
        search_ls = []
        search_fr = []  # dataframes
        
        for each_path in self.path_list:
            if each_path.find(partner_name) != -1:
                search_ls.append(i)
                
        self.frame_gen(search_ls, search_fr, skipping_rows)
        
        return search_fr
    
    def data_concat(self, partner, skips=0):
        """ Concatenates the dataframes per the partners into one dataset.
            Returns:
                    the_data: The unified data of each subdirectory partner files.
        """
        the_frames = self.search_crit(partner, skips)
        the_data = pd.concat(the_frames, ignore_index=True)
        
        return the_data
    
    def data_merge(self):
        """ Merge all frames into one dictionary per the amount of partners on record.
            Example: self.partners [Partner1, Something2] (In the respective dictionary: {"Partner1": 0,"Something2":3})
            All their dataframes are then allocated accordingly. Partner1: PANDAS_DATASET, Soemthing2: PANDAS_DATASET2
        """
        file_dictionary = {}
        
        for each_partner in self.partners:
            file_dictionary[each_partner] = self.data_concat(each_partner, self.partners[each_partner])

        self.aggregated_data = file_dictionary
    
    def excel_export(self):
        """ Export the data into an Excel file where the tabs are the dictionary names and the data is the pandas dataframes."""
        writer = pd.ExcelWriter(self.source+"_data_aggregate.xlsx", engine='xlsxwriter')  # Must add .xlsx
        for sheet, frame in self.aggregated_data.items():
            frame.to_excel(writer, index=None, sheet_name=sheet)
        writer.save()
        writer.close()
    
    # Show
    def show_list(self):
        return self.path_list
    
    def show_df_list(self):
        return self.df_list

    def show_aggregated_data(self):
        return self.aggregated_data

    def show_partners(self):
        return self.partners

    def show_source(self):
        return self.source
    
    # Process all
    def process_data(self):
        """ Aggregates all the important methods and executes accordingly."""
        self.data_gather()
        self.data_merge()
        self.excel_export()


if __name__ == "__main__":
    DTE = DataToExcel()
    DTE.process_data()
