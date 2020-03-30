#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Friday Mar 27 04:08:45 2020
@author: Alireza Ameri
alirezaameri.ca
Licensed under GNU Affero General Public License v3.0
"""


import xlrd


# add the excel file here 
loc =("adil.xlsx")

#set the index for the excel file
wb = xlrd.open_workbook(loc) 
sheet = wb.sheet_by_index(0) 
sheet.cell_value(0, 0)

#ask for the company
company = (input("Please enter the name of the company: "))

#going over the rows 
for i in range (sheet.nrows):

    
    if (company == 'Link' and sheet.cell_value(i+1,7) == 'Y'):
        x = (sheet.cell_value(i+1,5)) 
        sep = x.split('Added:')
        dates = sep[0]
        rest = sep[1]
        rest2 = rest.split("Removed:")
        added = rest2[0]
        rest3 = rest2[1]
        rest4 = rest3.split("Tier changes:")
        removed = rest4[0]
        rest5 = rest4[1]
        rest6 = rest5.split("Other changes:")
        tier = rest6[0]
        other = rest6[1]
                    
        print ('----')
        print ('Insurer: ', sheet.cell_value(i+1,2))
        print ('Formulary: ', sheet.cell_value(i+1,3))
        print ('To view the full formulary, please click on the link below.')
        print (sheet.cell_value(i+1,6))
        print (sep[0])
        
        
        
              
                
        #finish the print    
        print ()
        print ('Drugs Added: ', added)
        print ('Drugs Removed: ', removed)
        print ('Tier Changes: ', tier)
        print ('Other Changes: ', other)
        print ('----')
    
        
    elif (company == 'Neurocrine' and sheet.cell_value(i+1,8) == 'Y'):
        
        x = (sheet.cell_value(i+1,5)) 
        sep = x.split('Added:')
        dates = sep[0]
        rest = sep[1]
        rest2 = rest.split("Removed:")
        added = rest2[0]
        rest3 = rest2[1]
        rest4 = rest3.split("Tier changes:")
        removed = rest4[0]
        rest5 = rest4[1]
        rest6 = rest5.split("Other changes:")
        tier = rest6[0]
        other = rest6[1]
                    
        print ('----')
        print ('Insurer: ', sheet.cell_value(i+1,2))
        print ('Formulary: ', sheet.cell_value(i+1,3))
        print ('To view the full formulary, please click on the link below.')
        print (sheet.cell_value(i+1,6))
        print (sep[0])
        
        
        
              
                
        #finish the print    
        print ()
        print ('Drugs Added: ', added)
        print ('Drugs Removed: ', removed)
        print ('Tier Changes: ', tier)
        print ('Other Changes: ', other)
        print ('----')
        
    elif (company == 'Alexion' and sheet.cell_value(i+1,9) == 'Y'):
        
        x = (sheet.cell_value(i+1,5)) 
        sep = x.split('Added:')
        dates = sep[0]
        rest = sep[1]
        rest2 = rest.split("Removed:")
        added = rest2[0]
        rest3 = rest2[1]
        rest4 = rest3.split("Tier changes:")
        removed = rest4[0]
        rest5 = rest4[1]
        rest6 = rest5.split("Other changes:")
        tier = rest6[0]
        other = rest6[1]
                    
        print ('----')
        print ('Insurer: ', sheet.cell_value(i+1,2))
        print ('Formulary: ', sheet.cell_value(i+1,3))
        print ('To view the full formulary, please click on the link below.')
        print (sheet.cell_value(i+1,6))
        print (sep[0])
        
        
        
              
                
        #finish the print    
        print ()
        print ('Drugs Added: ', added)
        print ('Drugs Removed: ', removed)
        print ('Tier Changes: ', tier)
        print ('Other Changes: ', other)
        print ('----')
        
    elif (company == 'Endo' and sheet.cell_value(i+1,10) == 'Y'):
        
        x = (sheet.cell_value(i+1,5)) 
        sep = x.split('Added:')
        dates = sep[0]
        rest = sep[1]
        rest2 = rest.split("Removed:")
        added = rest2[0]
        rest3 = rest2[1]
        rest4 = rest3.split("Tier changes:")
        removed = rest4[0]
        rest5 = rest4[1]
        rest6 = rest5.split("Other changes:")
        tier = rest6[0]
        other = rest6[1]
                    
        print ('----')
        print ('Insurer: ', sheet.cell_value(i+1,2))
        print ('Formulary: ', sheet.cell_value(i+1,3))
        print ('To view the full formulary, please click on the link below.')
        print (sheet.cell_value(i+1,6))
        print (sep[0])
        
        
        
              
                
        #finish the print    
        print ()
        print ('Drugs Added: ', added)
        print ('Drugs Removed: ', removed)
        print ('Tier Changes: ', tier)
        print ('Other Changes: ', other)
        print ('----')
        
    elif (company == 'Novartis' and sheet.cell_value(i+1,11) == 'Y'):
        
        x = (sheet.cell_value(i+1,5)) 
        sep = x.split('Added:')
        dates = sep[0]
        rest = sep[1]
        rest2 = rest.split("Removed:")
        added = rest2[0]
        rest3 = rest2[1]
        rest4 = rest3.split("Tier changes:")
        removed = rest4[0]
        rest5 = rest4[1]
        rest6 = rest5.split("Other changes:")
        tier = rest6[0]
        other = rest6[1]
                    
        print ('----')
        print ('Insurer: ', sheet.cell_value(i+1,2))
        print ('Formulary: ', sheet.cell_value(i+1,3))
        print ('To view the full formulary, please click on the link below.')
        print (sheet.cell_value(i+1,6))
        print (sep[0])
        
        
        
              
                
        #finish the print    
        print ()
        print ('Drugs Added: ', added)
        print ('Drugs Removed: ', removed)
        print ('Tier Changes: ', tier)
        print ('Other Changes: ', other)
        print ('----')
        
    elif (company == 'Millennium Pharma' and sheet.cell_value(i+1,12) == 'Y'):
        
        x = (sheet.cell_value(i+1,5)) 
        sep = x.split('Added:')
        dates = sep[0]
        rest = sep[1]
        rest2 = rest.split("Removed:")
        added = rest2[0]
        rest3 = rest2[1]
        rest4 = rest3.split("Tier changes:")
        removed = rest4[0]
        rest5 = rest4[1]
        rest6 = rest5.split("Other changes:")
        tier = rest6[0]
        other = rest6[1]
                    
        print ('----')
        print ('Insurer: ', sheet.cell_value(i+1,2))
        print ('Formulary: ', sheet.cell_value(i+1,3))
        print ('To view the full formulary, please click on the link below.')
        print (sheet.cell_value(i+1,6))
        print (sep[0])
        
        
        
              
                
        #finish the print    
        print ()
        print ('Drugs Added: ', added)
        print ('Drugs Removed: ', removed)
        print ('Tier Changes: ', tier)
        print ('Other Changes: ', other)
        print ('----')
        
    elif (company == 'Retrophin' and sheet.cell_value(i+1,13) == 'Y'):
        
        x = (sheet.cell_value(i+1,5)) 
        sep = x.split('Added:')
        dates = sep[0]
        rest = sep[1]
        rest2 = rest.split("Removed:")
        added = rest2[0]
        rest3 = rest2[1]
        rest4 = rest3.split("Tier changes:")
        removed = rest4[0]
        rest5 = rest4[1]
        rest6 = rest5.split("Other changes:")
        tier = rest6[0]
        other = rest6[1]
                    
        print ('----')
        print ('Insurer: ', sheet.cell_value(i+1,2))
        print ('Formulary: ', sheet.cell_value(i+1,3))
        print ('To view the full formulary, please click on the link below.')
        print (sheet.cell_value(i+1,6))
        print (sep[0])
        
        
        
              
                
        #finish the print    
        print ()
        print ('Drugs Added: ', added)
        print ('Drugs Removed: ', removed)
        print ('Tier Changes: ', tier)
        print ('Other Changes: ', other)
        print ('----')
        
    elif (company == 'Pharmacyclics' and sheet.cell_value(i+1,14) == 'Y'):
        
        x = (sheet.cell_value(i+1,5)) 
        sep = x.split('Added:')
        dates = sep[0]
        rest = sep[1]
        rest2 = rest.split("Removed:")
        added = rest2[0]
        rest3 = rest2[1]
        rest4 = rest3.split("Tier changes:")
        removed = rest4[0]
        rest5 = rest4[1]
        rest6 = rest5.split("Other changes:")
        tier = rest6[0]
        other = rest6[1]
                    
        print ('----')
        print ('Insurer: ', sheet.cell_value(i+1,2))
        print ('Formulary: ', sheet.cell_value(i+1,3))
        print ('To view the full formulary, please click on the link below.')
        print (sheet.cell_value(i+1,6))
        print (sep[0])
        
        
        
              
                
        #finish the print    
        print ()
        print ('Drugs Added: ', added)
        print ('Drugs Removed: ', removed)
        print ('Tier Changes: ', tier)
        print ('Other Changes: ', other)
        print ('----')
        
    elif (company == 'Pacira' and sheet.cell_value(i+1,15) == 'Y'):
        
        x = (sheet.cell_value(i+1,5)) 
        sep = x.split('Added:')
        dates = sep[0]
        rest = sep[1]
        rest2 = rest.split("Removed:")
        added = rest2[0]
        rest3 = rest2[1]
        rest4 = rest3.split("Tier changes:")
        removed = rest4[0]
        rest5 = rest4[1]
        rest6 = rest5.split("Other changes:")
        tier = rest6[0]
        other = rest6[1]
                    
        print ('----')
        print ('Insurer: ', sheet.cell_value(i+1,2))
        print ('Formulary: ', sheet.cell_value(i+1,3))
        print ('To view the full formulary, please click on the link below.')
        print (sheet.cell_value(i+1,6))
        print (sep[0])
        
        
        
              
                
        #finish the print    
        print ()
        print ('Drugs Added: ', added)
        print ('Drugs Removed: ', removed)
        print ('Tier Changes: ', tier)
        print ('Other Changes: ', other)
        print ('----')
        
    elif (company == 'Smith & Nephew' and sheet.cell_value(i+1,16) == 'Y'):
        
        x = (sheet.cell_value(i+1,5)) 
        sep = x.split('Added:')
        dates = sep[0]
        rest = sep[1]
        rest2 = rest.split("Removed:")
        added = rest2[0]
        rest3 = rest2[1]
        rest4 = rest3.split("Tier changes:")
        removed = rest4[0]
        rest5 = rest4[1]
        rest6 = rest5.split("Other changes:")
        tier = rest6[0]
        other = rest6[1]
                    
        print ('----')
        print ('Insurer: ', sheet.cell_value(i+1,2))
        print ('Formulary: ', sheet.cell_value(i+1,3))
        print ('To view the full formulary, please click on the link below.')
        print (sheet.cell_value(i+1,6))
        print (sep[0])
        
        
        
              
                
        #finish the print    
        print ()
        print ('Drugs Added: ', added)
        print ('Drugs Removed: ', removed)
        print ('Tier Changes: ', tier)
        print ('Other Changes: ', other)
        print ('----')
        
    elif (company == 'Mylan' and sheet.cell_value(i+1,17) == 'Y'):
        
        x = (sheet.cell_value(i+1,5)) 
        sep = x.split('Added:')
        dates = sep[0]
        rest = sep[1]
        rest2 = rest.split("Removed:")
        added = rest2[0]
        rest3 = rest2[1]
        rest4 = rest3.split("Tier changes:")
        removed = rest4[0]
        rest5 = rest4[1]
        rest6 = rest5.split("Other changes:")
        tier = rest6[0]
        other = rest6[1]
                    
        print ('----')
        print ('Insurer: ', sheet.cell_value(i+1,2))
        print ('Formulary: ', sheet.cell_value(i+1,3))
        print ('To view the full formulary, please click on the link below.')
        print (sheet.cell_value(i+1,6))
        print (sep[0])
        
        
        
              
                
        #finish the print    
        print ()
        print ('Drugs Added: ', added)
        print ('Drugs Removed: ', removed)
        print ('Tier Changes: ', tier)
        print ('Other Changes: ', other)
        print ('----')
        
    elif (company == 'NXDC' and sheet.cell_value(i+1,18) == 'Y'):
        
        x = (sheet.cell_value(i+1,5)) 
        sep = x.split('Added:')
        dates = sep[0]
        rest = sep[1]
        rest2 = rest.split("Removed:")
        added = rest2[0]
        rest3 = rest2[1]
        rest4 = rest3.split("Tier changes:")
        removed = rest4[0]
        rest5 = rest4[1]
        rest6 = rest5.split("Other changes:")
        tier = rest6[0]
        other = rest6[1]
                    
        print ('----')
        print ('Insurer: ', sheet.cell_value(i+1,2))
        print ('Formulary: ', sheet.cell_value(i+1,3))
        print ('To view the full formulary, please click on the link below.')
        print (sheet.cell_value(i+1,6))
        print (sep[0])
        
        
        
              
                
        #finish the print    
        print ()
        print ('Drugs Added: ', added)
        print ('Drugs Removed: ', removed)
        print ('Tier Changes: ', tier)
        print ('Other Changes: ', other)
        print ('----')
        

      
        






"""

* master set of ifs 

    if (company == 'Link' and sheet.cell_value(i+1,7) == 'Y'):
        
    if (company == 'Neurocrine' and sheet.cell_value(i+1,8) == 'Y'):
        
    if (company == 'Alexion' and sheet.cell_value(i+1,9) == 'Y'):
        
    if (company == 'Endo' and sheet.cell_value(i+1,10) == 'Y'):
        
    if (company == 'Novartis' and sheet.cell_value(i+1,11) == 'Y'):
        
    if (company == 'Millennium Pharma' and sheet.cell_value(i+1,12) == 'Y'):
        
    if (company == 'Retrophin' and sheet.cell_value(i+1,13) == 'Y'):
        
    if (company == 'Pharmacyclics' and sheet.cell_value(i+1,14) == 'Y'):
        
    if (company == 'Pacira' and sheet.cell_value(i+1,15) == 'Y'):
        
    if (company == 'Smith & Nephew' and sheet.cell_value(i+1,16) == 'Y'):
        
    if (company == 'Mylan' and sheet.cell_value(i+1,17) == 'Y'):
        
    if (company == 'NXDC' and sheet.cell_value(i+1,18) == 'Y'):
        
        
"""



"""


#  master set of code

       #searh within the cell to figure out the notes
       x = (sheet.cell_value(i+1,5)) 
       sep = x.split('Added:')
       dates = sep[0]
       rest = sep[1]
       rest2 = rest.split("Removed:")
       added = rest2[0]
       rest3 = rest2[1]
       rest4 = rest3.split("Tier changes:")
       removed = rest4[0]
       rest5 = rest4[1]
       rest6 = rest5.split("Other changes:")
       tier = rest6[0]
       other = rest6[1]
            
       print ('----')
       print ('Insurer: ', sheet.cell_value(i+1,2))
       print ('Formulary: ', sheet.cell_value(i+1,3))
       print ('To view the full formulary, please click on the link below.')
       print (sheet.cell_value(i+1,6))
       print (sep[0])
        
        
      
      
        
        #finish the print    
        print ()
        print ('Drugs Added: ', added)
        print ('Drugs Removed: ', removed)
        print ('Tier Changes: ', tier)
        print ('Other Changes: ', other)
        print ('----')
"""
    

    
"""
Licensed under GNU Affero General Public License v3.0
"""
