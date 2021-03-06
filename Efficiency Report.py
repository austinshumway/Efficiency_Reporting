import pandas as pd
import xlsxwriter
import openpyxl
import os
from datetime import date
import numpy as np

#### Reads in the Excel files for the Proposal Form and Resource Allocation.
### We need to transition these onto the company portal so we don't have to keep switching the directories depending on who runs the program
proposal_form = pd.read_excel(r'C:\Users\Austin\OneDrive - The Cicero Group\Cicero\Proposal Form.xlsx')
resource_allocation = r'C:\Users\Austin\The Cicero Group\Company Portal - Resource Allocation\2020 Resource Allocation (Cicero).xlsx'
billable_projects = pd.read_excel(resource_allocation, 'Billable Projects', skiprows=1)
billable_projects_completed = pd.read_excel(resource_allocation, 'Billable Projects (Completed)', skiprows=1)
ed_direction = pd.read_excel(resource_allocation, 'Ed Direction', skiprows=1)
ra = pd.concat([billable_projects, billable_projects_completed, ed_direction])

## I noticed a function like this in Isaac's code so I just quickly wrote it out in case it needed to be used later
##roles_renamed = ra.rename(columns={'Partner Hours':'PART', 'Principal Hours':'PRIN', 'Engagement Manager Hours':'EM', 'Associate Hours':'ASSOC', 'Analyst Hours':'BA'}, inplace=True)

###insert melt function here
##currently not working. I don't know why it won't select data in a range
ra_combined = ra.melt(id_vars='date', col_level=['1/1/2020']:['4/26/2021'])


#this is good
today = date.today()
Month = today.month

### Filtering and Aggregation
ra_sorted = ra[ra['Status'].isin(['Live'])]
grouped = ra.groupby(['Project Number', 'Role']).sum()

sort = proposal_form[proposal_form['Project Status'].isin(['Live'])]
pf_reordered = sort.reindex(
    columns=['Project Number', 'Project Name', 'Client Name', 'Practice Area', 'Project Lead', 'Executive Sales Lead',
             'Executive Delivery Lead'])

