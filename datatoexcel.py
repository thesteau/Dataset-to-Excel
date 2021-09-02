# coding: utf-8
# Code written by Steven Au for data automation.
# Required modules installed from PIP: pandas, xlsxwriter
import os

import pandas as pd


class DataToExcel:
    
    def __init__(self):
        self.source = None
        self.partners = self.the_partners()
        self.path_list = []
        self.df_list = []
        self.aggregated_data = None
    
    def the_partners(self):
        """ Update partners here,"""
        # PartnerName (Has to match the folder name), skipHeadingRows
        partner_dictionary = {
            "": 0
        }
        return partner_dictionary
    
    def set_source(self):
        self.source = input("Enter folder source path: ")
    
    # Aggregators
    def data_gather(self):
        """ Gather all data from sosurce directory folder"""
        if self.source is None:
            self.set_source()
            
        for root, dirs, files in os.walk(self.source):
            for name in files:
                if name.endswith('.csv') or name.endswith('.tsv'): 
                    self.path_list.append(os.path.join(root, name))
    
    def frame_gen(self, file_list, frame_type, skipping_rows):
        """Genereate pandas dataframes"""
        if self.source is None:
            return
        
        for i in file_list:
            if i.endswith('.csv'):
                df = pd.read_csv(i, skiprows=skipping_rows)  

            elif i.endswith('.tsv'):
                df = pd.read_csv(i, sep='\t', skiprows=skipping_rows)
                
            df.insert(0, 'filename', i) 
            frame_type.append(df)
                
    def search_crit(self, partner_name, skipping_rows):
        """Criteria for search partner"""
        search_ls = []
        search_fr = []  # dataframes
        
        for i in self.path_list:
            if i.find(partner_name) != -1:
                search_ls.append(i)
                
        self.frame_gen(search_ls, search_fr, skipping_rows)
        
        return search_fr
    
    def data_concat(self, partner, skips=0):
        """Concatenates the data per the partners"""
        the_frames = self.search_crit(partner, skips)
        
        the_data = pd.concat(the_frames, ignore_index=True)
        
        return the_data
    
    def data_merge(self):
        """ Merge all frames into one dictionary"""
        file_dictionary = {}
        for i in self.partners:
            file_dictionary[i] = self.data_concat(i, self.partners[i])

        self.aggregated_data = file_dictionary
    
    def excel_export(self):
        writer = pd.ExcelWriter(self.source+"_data_aggregate.xlsx", engine='xlsxwriter')
        for sheet, frame in self.aggregated_data.items():
            frame.to_excel(writer, index=None, sheet_name=sheet)
        writer.save()
        writer.close()
    
    # Show
    def show_list(self):
        return self.path_list
    
    def show_df_list(self):
        return self.df_list
    
    # Process all
    def process_data(self):
        self.data_gather()
        self.data_merge()
        self.excel_export()


if __name__ == "__main__":
    DTE = DataToExcel()
    DTE.process_data()





