#!/usr/bin/env python
# coding: utf-8

# Begin with the standard imports.

import numpy as np
import pandas as pd
import openpyxl
#from fpdf import FPDF

# Take user input for the keyword to be searched. 
keyword = input('Enter keyword: ')

# Create a pandas dataframe from the CORDIS dataset.
projects = pd.read_excel("data/xlsx/project.xlsx")

objectives = pd.Series(projects['objective'], dtype="string").str.lower()
# reg = r"\.|\s[the]\.\s"
counts = objectives.str.count(" " + keyword + " ")
counts = counts[counts>0]
# current implementation ignores cases such as .{term} or {term}. Add regex to overcome these cases. 
sortedCounts = counts.sort_values(ascending=False)

df = projects.iloc[sortedCounts.index]
projectList = df[['id', 'title']]

orgs = pd.read_excel("data/xlsx/organization.xlsx")

t = orgs[orgs['projectID'].isin(projectList['id'])]
companyList = t[['organisationID', 'name']]
#companyList.reset_index(drop=True)
companyList = companyList.drop_duplicates()
topProjects = projectList.iloc[0:5]

with pd.ExcelWriter("keywordResults.xlsx") as writer:
    projectList.to_excel(writer, sheet_name="projects", index=False, float_format="%f", header=["projectID", "projectName"])
    companyList.to_excel(writer, sheet_name="organizations", index=False, float_format="%f")
    topProjects.to_excel(writer, sheet_name="top_projects", index=False, float_format="%f")
