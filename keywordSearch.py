#!/usr/bin/env python
# coding: utf-8

# Begin with the standard imports.

import numpy as np
import pandas as pd
import openpyxl
from PyPDF2 import PdfFileMerger 
from pathlib import Path
import urllib.request
import os
#from fpdf import FPDF

# Take user input for the keyword to be searched. 
keyword = input('Enter keyword: ')

# Create a pandas dataframe from the CORDIS dataset.
projects = pd.read_excel("data/xlsx/project.xlsx")

objectives = pd.Series(projects['objective'], dtype="string").str.lower()
counts = objectives.str.count(" " + keyword + " ")
counts = counts[counts>0]
# current implementation ignores cases such as .{term} or {term}. Add regex to overcome these cases. 
sortedCounts = counts.sort_values(ascending=False)

df = projects.iloc[sortedCounts.index]
projectList = df[['id', 'title', 'objective']]

orgs = pd.read_excel("data/xlsx/organization.xlsx")

t = orgs[orgs['projectID'].isin(projectList['id'])]
companyList = t[['organisationID', 'name', 'country', 'totalCost', 'projectID']]
companyList = companyList.sort_values(by=['name'])
#companyList = companyList.drop_duplicates()
nRow = projectList.shape[0]
maxSlice = 5
if nRow < 5: 
    maxSlice = nRow
topProjects = projectList.iloc[0:maxSlice]

with pd.ExcelWriter("keywordResults.xlsx") as writer:
    projectList.to_excel(writer, sheet_name="projects", index=False, float_format="%f", header=["projectID", "projectName", "objective"])
    companyList.to_excel(writer, sheet_name="organizations", index=False, float_format="%f")
    topProjects.to_excel(writer, sheet_name="top_projects", index=False, float_format="%f")


def download_file(download_url, filename):
    response = urllib.request.urlopen(download_url)    
    file = open(filename + ".pdf", 'wb')
    file.write(response.read())
    file.close()
 
pdf_urls = []
for id in topProjects['id']:
    pdf_urls.append("https://cordis.europa.eu/project/id/" + str(id) + "/en?format=pdf")

pdf_list = []
for index, pdf_path in enumerate(pdf_urls):
    download_file(pdf_path, str(index))
    pdf_list.append(str(index) + ".pdf")

pdfMerger = PdfFileMerger()
for fileName in pdf_list:
    pdfMerger.append(fileName)
    os.remove(fileName)

with Path("report.pdf").open(mode="wb") as output_file:
    pdfMerger.write(output_file)
