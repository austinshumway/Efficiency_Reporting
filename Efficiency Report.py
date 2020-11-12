import pandas as pd
import xlsxwriter
import openpyxl
import os
from datetime import date
import numpy as np

PF = r'C:\Users\Austin\OneDrive - The Cicero Group\Cicero\Proposal Form.xlsx'
RA = r'C:\Users\Austin\The Cicero Group\Company Portal - Resource Allocation\2020 Resource Allocation (Cicero).xlsx'
test = r'C:\Users\Austin\OneDrive - The Cicero Group\Cicero\Project Efficiency Reporting\Test.xlsx'

pf = pd.read_excel(PF)
bp = pd.read_excel(RA, 'Billable Projects', skiprows=1)
bpc = pd.read_excel(RA, 'Billable Projects (Completed)', skiprows=1)
ed = pd.read_excel(RA, 'Ed Direction', skiprows=1)

if os.path.exists(test):
    os.remove(test)

ra = pd.concat([bp, bpc, ed])
ra_sorted = ra[ra['Status'].isin(['Live'])]
grouped = ra_sorted.groupby(['Project Number']).sum()
rat=grouped.T

today = date.today()
Month = today.month
rate = rat.drop([])

rat.to_excel(test)

sort = pf[pf['Project Status'].isin(['Live'])]
pf_reordered = sort.reindex(
    columns=['Project Number', 'Project Name', 'Client Name', 'Practice Area', 'Project Lead', 'Executive Sales Lead',
             'Executive Delivery Lead'])
# ra_fixed = ra.reindex(columns=['Project Number', 'Role', 'Hours Used', '11/2/2020', '11/9/2020', '11/16/2020', '11/23/2020', '11/30/2020'])

# merge=pd.merge(pf_reordered, ra_fixed, how='left', on=['Project Number'])
# ran=ra.T
# merge.insert(14, 'Summed', '=sum(k2:o2)', allow_duplicates=False)

# merge.to_excel(test)


# a = r.reindex(columns = ['Created', 'Project Number', 'Employee', 'Employee Email', 'Mentor Email', 'Reviewer', 'Reviewer Email', 'Cicero Way', 'Comments (Cicero Way Implementation)', 'Communication and Client Interaction (Professional)', 'Communication and Client Interaction (Efficient and Strategic)', 'Comments (Communication and Client Interaction)', 'Data Analysis (Technically Skilled)', 'Data Analysis (Accurate)', 'Comments (Data Analysis)', 'Deliverable Development & Insight (Detailed)', 'Deliverable Development & Insight (Insight-driven)', 'Deliverable Development & Insight (Strategic Thinker)', 'Comments (Deliverable Development & Insight)', 'Engagement Methodology (Structured)', 'Engagement Methodology (Identifies Client Needs)', 'Comments (Engagement Methodology)', 'Project Management (Meets Deadlines)', 'Project Management (Returns and Reports)', 'Project Management (Utilizes Worksteps)', 'Project Management (Long-term Vision)', 'Comments (Project Management)', 'Intangibles (Implements Feedback)', 'Intangibles (Shows Improvement)', 'Intangibles (Eases Burdens)', 'Comments (Intangibles)', 'Areas of Strength', 'Areas of Improvement', 'Score Count', 'Score Sum', 'Score Avg', 'Item Type', 'Path '])
# ra['Year'] = pd.DatetimeIndex(ra['Created']).year
# ra['Month'] = pd.DatetimeIndex(ra['Created']).month
# year_filter = a[a['Year'] == 2020]
# month_filter = year_filter[year_filter['Month'] == Month]
# print(month_filter)
