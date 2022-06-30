import numpy as np
import pandas as pd
import openpyxl
from pathlib import Path
import urllib.request
import os
import spacy
import pytextrank
import plotly.graph_objects as go
import json

isUnion = True
# Take user input for the keyword to be searched. 
print('Welcome to keyword search. Default search mode will look for keyword1 OR keyword2.')
keywordList = []
keyword = input('Enter first keyword, or enter 1 if you want to change mode to intersection search: ')
if keyword == '1':
  isUnion = False
  print('You are now doing a search based on keyword1 AND keywordB...')
while keyword != "":
  if keyword != '1':
    keywordList.append(keyword.lower())
    keyword = input('Enter next keyword (enter nothing if done): ')
  else:
    keyword = input('Enter first keyword: ')
print("Searching for keyword(s)...")

if not keywordList:
  print("No keywords given")
else:
  projects = pd.read_csv("data/csv/project.csv", delimiter=";")

  useDF = projects
  if not isUnion:
    #drops all the rows where one of the keywords don't exist.
    for key in keywordList:
      useDF = useDF[~(~useDF['objective'].str.contains(key) & ~useDF['title'].str.contains(key))]

  objectives = pd.Series(useDF['objective'], dtype="string").str.lower()
  titles = pd.Series(useDF['title'], dtype="string").str.lower()

  counts = objectives.str.count(keywordList[0]) + titles.str.count(keywordList[0])
  for key in keywordList[1:]:
    counts = counts + objectives.str.count(key) + titles.str.count(key)
  counts = counts[counts>0]
  if counts.size == 0:
    print("Keyword(s) not found")
  else:
    sortedCounts = counts.sort_values(ascending=False)
    projectList = projects.iloc[sortedCounts.index][['id', 'title', 'startDate', 'endDate', 'totalCost', 'ecMaxContribution', 'legalBasis', 'masterCall', 'subCall', 'objective']]

    # creates the company(organizations) slices
    orgs = pd.read_csv("data/csv/organization.csv", delimiter=";")
    t = orgs[orgs['projectID'].isin(projectList['id'])]
    companyList = t[['organisationID', 'name', 'country', 'totalCost', 'projectID', 'role']]
    companyList = companyList.sort_values(by=['name'])

    # to limit data slice if filter results in fewer than 5 rows
    nRow = projectList.shape[0]
    maxSlice = 5
    if nRow < 5: 
        maxSlice = nRow
    topRelevanceProjects = projectList.iloc[0:maxSlice]
    topCostProjects = projectList.sort_values(by=['totalCost'],ascending=False).iloc[0:maxSlice]

    # writes results into the keywordResults.xlsx spreadsheet
    with pd.ExcelWriter("keywordResults.xlsx") as writer:
        projectList.to_excel(writer, sheet_name="projects", index=False, float_format="%f")
        companyList.to_excel(writer, sheet_name="organizations", index=False, float_format="%f")
        topRelevanceProjects.to_excel(writer, sheet_name="top projects (relevance)", index=False, float_format="%f")
        topCostProjects.to_excel(writer, sheet_name="top projects (cost)", index=False, float_format="%f")

    print("Spreadsheet complete!")

    # Uses textrank NLP to generate other keywords from the selected projects
    print("Generating other keywords...")
    numPrint = 20
    text = ' '.join(projects['objective'].iloc[projectList.index])
    #TODO: investigate whether we want to have keywords be case-sensitive
    #text = text.lower() 
    nlp = spacy.load("en_core_web_sm")
    nlp.add_pipe("textrank")
    nlp_words = []
    if (len(text) < 1000000):
      doc = nlp(text)
      for n, phrase in enumerate(doc._.phrases):
          nlp_words.append(phrase.text)
          if n == numPrint: break
      print("Other keywords generated!")
    else:
      print("Keyword results too large to generate other keywords.")

    def add_names(df):
      nameList = []
      for ids in df['projectIDs']:
        nameList.append(projectList.loc[projectList['id'].isin(ids)]['title'].tolist())
      nameList = pd.Series(nameList)
      nameList.name = 'projects'
      return nameList

    def find_project_IDs(df):
      moreTopData = companyList.loc[companyList['name'].isin(df['name'])]
      listOfIdLists = moreTopData.groupby(['name']).agg(tuple)['projectID'].map(list)
      listOfIdLists.name = 'projectIDs'
      return listOfIdLists

    def generate_counts_df(df):
      countsDF = df['name'].value_counts().to_frame().reset_index(level=0)
      countsDF.columns = ['name', 'nProjects']
      countsDF = countsDF.merge(find_project_IDs(countsDF), how="left", on="name")
      countsDF = countsDF.join(add_names(countsDF))
      return countsDF

    coordinatorList = companyList[companyList['role']=="coordinator"]
    #coordinatorCounts = coordinatorList['name'].value_counts().to_frame().reset_index(level=0)
    #coordinatorCounts.columns = ['name', 'nProjects']
    #coordinatorCounts = coordinatorCounts.join(find_project_IDs(coordinatorCounts))
    #coordinatorCounts = coordinatorCounts.join(add_names(coordinatorCounts))
    coordinatorCounts = generate_counts_df(coordinatorList)

    otherList = companyList[companyList['role']!="coordinator"]
    otherCounts = generate_counts_df(otherList)

    sumCost = projectList['totalCost'].sum()
    sumEC = projectList['ecMaxContribution'].sum()
    
    labels = ['EC Contribution', 'Other Contribution']
    values = [sumEC, sumCost-sumEC]
    figCosts = go.Figure(data=[go.Pie(labels=labels, values=values)], layout=go.Layout(
            title=go.layout.Title(text="EU Contributions Relative to Other Contributions")
        ))

    companyCosts = companyList.groupby(['name']).sum().sort_values(by=['totalCost'],ascending=False)
    companyCosts = companyCosts.reset_index(level=0)
    companyCosts = companyCosts.drop(['organisationID', 'projectID'], axis=1)
    companyCosts = companyCosts.merge(companyList.drop_duplicates(subset=['name']), how="left", on='name')
    ccounts = companyList['name'].value_counts().to_frame().reset_index(level=0)
    ccounts.columns = ['name', 'nProjects']
    companyCosts = companyCosts.merge(ccounts, how='left', on=['name'])
    companyCosts = companyCosts.drop(['totalCost_y', 'projectID','organisationID', 'role'], axis=1)
    topC = companyCosts.head()
    topC = topC.merge(find_project_IDs(topC), how="left", on="name")
    topC = topC.join(add_names(topC))

    callCounts = projectList['masterCall'].value_counts()
    #callCounts = pd.concat([projectList['masterCall'],projectList['subCall']]).value_counts()

    # Creates html report
    print("Generating html report...")
    page_title = "Keyword Report"
    if isUnion:
      title_text="Results for '" + " OR ".join(keywordList) + "' in CORDIS Horizon 2020"
    else:
      title_text="Results for '" + " AND ".join(keywordList) + "' in CORDIS Horizon 2020"
    text="Total number of results: " + str(len(projectList.index))
    projects_text='Most relevant projects by frequency of keyword(s)'
    projects_subtitle='These projects used the keywords most frequently in their title and objective statements.'
    projects_text_cost='Most costly projects with keyword(s)'
    projects_text_cost_subtitle='These projects were the most costly out of the projects that contained the keyword(s).'
    companies_text1='All companies involved as coordinators'
    companies_text2='All participant companies'
    links_text='Links to projects'
    total_text='Total cost of all projects: &euro;'
    ec_text='Max contributions by the EC: &euro;'
    call_counts_text='The number of each masterCall type'
    keywords_text ='Other prominent keywords from the selected projects'
    commonCompanies_text = 'Top companies funding multiple projects with keyword(s)'
    commonCompanies_subtext = 'The project IDs and their titles are pulled from all projects a company is involved in, either as a coordinator or participant.'

    urls = []
    for id in topRelevanceProjects['id']:
      urls.append("https://cordis.europa.eu/project/id/" + str(id))

    html = f'''
        <html>
          <head>
              <title>{page_title}</title>
              <link rel="stylesheet" href="reportStyle.css" type="text/css" media="all">
              <meta charset="UTF-8">
          </head>
          <body>
            <div id="title">
              <h1>{title_text}</h1>
                <p>{text}</p>
            </div>
              <h2>{projects_text}</h2>
                <p>{projects_subtitle}</p>
                {topRelevanceProjects.to_html(index=False, justify="center", table_id="topProjects", columns=['id', 'title', 'startDate', 'endDate', 'totalCost', 'ecMaxContribution',
       'legalBasis', 'masterCall', 'subCall'])}
              <h2>{links_text}</h2>
              <script>
                const urlList = {json.dumps(urls)};
                const projects = {json.dumps(topRelevanceProjects["title"].tolist())};
                for (let i = 0; i < {len(urls)}; i++):
                  var a = document.createElement('a');
                  a.textContent = {json.dumps(topRelevanceProjects["title"].tolist())}[i];
                  a.setAttribute('href', {json.dumps(urls)}[i]);
                  document.body.appendChild(a);
              </script>
                <a href={urls[0]}>{topRelevanceProjects["title"].iloc[0]}</a><br>
                <a href={urls[1]}>{topRelevanceProjects["title"].iloc[1]}</a><br>
                <a href={urls[2]}>{topRelevanceProjects["title"].iloc[2]}</a><br>
                <a href={urls[3]}>{topRelevanceProjects["title"].iloc[3]}</a><br>
                <a href={urls[4]}>{topRelevanceProjects["title"].iloc[4]}</a><br>
              <h2>{projects_text_cost}</h2>
                <p>{projects_text_cost_subtitle}</p>
                {topCostProjects.to_html(index=False, justify="center", table_id="costProjects", columns=['id', 'title', 'startDate', 'endDate', 'totalCost', 'ecMaxContribution',
       'legalBasis', 'masterCall', 'subCall'])}
              <span id="cost">
                <h2>{total_text}{sumCost}</h2>
                <h2>{ec_text}{sumEC}</h2>
                {figCosts.to_html(full_html=False)}
              </span>
              <span id ="callCounts">
                <h2>{call_counts_text}</h2>
                {callCounts.to_frame().to_html(header=False, col_space="10cm")}
              </span>
              <span id ="keywords">
                <h2>{keywords_text}</h2>
                {", ".join(nlp_words)}
              </span>
              <span id="commonCompanies">
                <h2>{commonCompanies_text}</h2>
                  <p>{commonCompanies_subtext}</p>
                  <h3>Coordinators</h3>
                    {coordinatorCounts.head(5).to_html(index=False)}
                  <h3>Non-Coordinators</h3>
                    {otherCounts.head(5).to_html(index=False)}
              </span>
              <h2>Top 5 companies by total cost</h2>
                {topC.to_html(index=False)}
              <h2>{companies_text1}<h2>
                {coordinatorList.to_html(index=False)}
              <h2>{companies_text2}<h2>
                {otherList.to_html(index=False)}
          </body>
        </html>
    '''

    with open('html_report.html', 'w') as f:
        f.write(html)
    print("Html report complete!")