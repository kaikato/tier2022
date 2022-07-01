import numpy as np
import pandas as pd
import openpyxl
import os
import spacy
import pytextrank
import plotly.graph_objects as go
import json
import xlsxwriter
import plotly.express as px
pd.options.mode.chained_assignment = None

isPillar2 = False
isUnion = True
# Take user input for the keyword to be searched. 
print('Welcome to keyword search. Default search mode will look for keyword1 OR keyword2.')
pillarType = input("Are you searching for Pillar 2? (y/n): ")
if pillarType == 'y':
  isPillar2 = True
  print('Searching through ONLY pillar 2')
else:
  print('Searching through all pillars.')
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
  pillar2 = 'BBI|CS2|CSA|ECSEL|EuroHPC|FCH2|IA|IMI2|RIA|SESAR|Shift2Rail|SME'
  if isPillar2:
    projects = projects[projects['fundingScheme'].str.contains(pillar2)]

  useDF = projects
  if not isUnion:
    #drops all the rows where one of the keywords don't exist.
    for key in keywordList:
      useDF = useDF[~(~useDF['objective'].str.contains(key) & ~useDF['title'].str.contains(key))]

  objectives = pd.Series(useDF['objective'], dtype="string").str.lower()
  titles = pd.Series(useDF['title'], dtype="string").str.lower()

  #counts the number of keyword occurences.
  counts = objectives.str.count(keywordList[0]) + titles.str.count(keywordList[0])
  for key in keywordList[1:]:
    counts = counts + objectives.str.count(key) + titles.str.count(key)
  counts = counts[counts>0]
  if counts.size == 0:
    print("Keyword(s) not found")
  else:
    sortedCounts = counts.sort_values(ascending=False)
    projectList = projects.loc[sortedCounts.index][['id', 'acronym', 'title', 'startDate', 'endDate', 'totalCost', 'ecMaxContribution', 'legalBasis', 'masterCall', 'subCall', 'fundingScheme', 'objective']]
    projectList['percentEUFunded'] = (projectList['ecMaxContribution']/projectList['totalCost']).apply(lambda x: f"{x:.0%}")
    print(str(len(projectList.index)) + ' projects found')

    # creates the company(organizations) slices
    orgs = pd.read_csv("data/csv/organization.csv", delimiter=";")
    t = orgs[orgs['projectID'].isin(projectList['id'])]
    companyList = t[['organisationID', 'name', 'shortName', 'country', 'totalCost', 'projectID', 'role']]
    companyList = companyList.sort_values(by=['name'])
    topRelevanceProjects = projectList.head()
    topCostProjects = projectList.sort_values(by=['totalCost'],ascending=False).head()

    # Uses textrank NLP to generate other keywords from the selected projects
    print("Generating other keywords...")
    numPrint = 20
    text = ' '.join(projects['objective'].loc[projectList.index])
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

    def generate_counts_df(df):
      countsDF = df['name'].value_counts().to_frame().reset_index(level=0)
      countsDF.columns = ['name', 'nProjects']
      countsDF = countsDF.merge(df.drop_duplicates(subset=['name']), how="left", on='name')
      countsDF = countsDF.drop(['totalCost','projectID','organisationID', 'role'], axis=1)
      return countsDF

    def generate_costs_df(df1, df2):
      costDF = df1.groupby(['name']).sum().sort_values(by=['totalCost'],ascending=False)
      costDF = costDF.reset_index(level=0)
      costDF = costDF.drop(['organisationID', 'projectID'], axis=1)
      costDF = costDF.merge(df1.drop_duplicates(subset=['name']), how="left", on='name')
      costDF = costDF.merge(df2[['name','nProjects']], how='left', on='name')
      costDF = costDF.drop(['totalCost_y','projectID','organisationID', 'role'], axis=1)
      return costDF

    # coordinatorCounts is the 'organisation' database filtered for coordinators and ranked by # of projects
    coordinatorList = companyList[companyList['role']=="coordinator"]
    coordinatorCounts = generate_counts_df(coordinatorList)

    # allCounts is the 'organisation' database with all coordinators + participants and ranked by # of projects
    allCounts = generate_counts_df(companyList)

    sumCost = projectList['totalCost'].sum()
    sumEC = projectList['ecMaxContribution'].sum()
    
    labels = ['EC Max Contribution', 'Other Contribution']
    values = [sumEC, sumCost-sumEC]
    figCosts = go.Figure(data=[go.Pie(labels=labels, values=values)], layout=go.Layout(
            title=go.layout.Title(text="EU Contributions Relative to Other Contributions")
        ))
    
    sumCost = "{:,.0f}".format(sumCost)
    sumEC = "{:,.0f}".format(sumEC)

    # coordinatorCosts is the organisations database filtered for coordinators and ranked by total costs
    coordinatorCosts = generate_costs_df(coordinatorList, coordinatorCounts)

    # allCosts is the organisations database sorted by total costs for each org
    allCosts = generate_costs_df(companyList, allCounts)

    callCounts = projectList['masterCall'].value_counts()
    #callCounts = pd.concat([projectList['masterCall'],projectList['subCall']]).value_counts()

    # writes results into the keywordResults.xlsx spreadsheet
    with pd.ExcelWriter("keywordResults.xlsx", engine='xlsxwriter') as writer:
        projectList.to_excel(writer, sheet_name="projects", index=False, float_format="%f")
        companyList.to_excel(writer, sheet_name="organizations", index=False, float_format="%f")
        topRelevanceProjects.to_excel(writer, sheet_name="top projects (relevance)", index=False, float_format="%f")
        topCostProjects.to_excel(writer, sheet_name="top projects (cost)", index=False, float_format="%f")
        callCounts.to_frame().to_excel(writer, sheet_name='number of calls')
        worksheet = writer.sheets['number of calls']
        worksheet.set_column('A:A', 30)
        coordinatorCounts.to_excel(writer, sheet_name='nProjects (coordinators)', index=False)
        allCounts.to_excel(writer, sheet_name='nProjects (all)', index=False)
        coordinatorCosts.to_excel(writer, sheet_name='by cost (coordinators)', index=False)
        allCosts.to_excel(writer, sheet_name='by cost (all)', index=False)

    print("Spreadsheet complete!")

    # generate graphs
    topRelevanceProjects['otherContribution'] = topRelevanceProjects['totalCost']-topRelevanceProjects['ecMaxContribution']
    topCostProjects['otherContribution'] = topCostProjects['totalCost']-topCostProjects['ecMaxContribution']
    relevantProjectCounts = px.bar(topRelevanceProjects, title="Most Relevant Projects", x='acronym', y=[topRelevanceProjects['ecMaxContribution'],topRelevanceProjects['otherContribution']], labels={"acronym":"Organization", "value":"Total Cost (Euro)"})
    costlyProjectCounts = px.bar(topCostProjects, title="Most Costly Projects", x='acronym', y=[topCostProjects['ecMaxContribution'],topCostProjects['otherContribution']], labels={"acronym":"Organization", "value":"Total Cost (Euro)"})
    coordinatorCountsGraph = px.bar(coordinatorCounts.head(), title="Coordinator No. of Projects", x='shortName', y='nProjects', labels={'shortName':'Organization','nProjects':'Number of Projects'})
    coordinatorCountsGraph.update_traces(marker_color='green')
    allCountsGraph = px.bar(allCounts.head(), title="All Companies No. of Projects", x='shortName', y='nProjects', labels={'shortName':'Organization','nProjects':'Number of Projects'})
    coordinatorCostsGraph = px.bar(coordinatorCosts.head(), title="Coordinator Total Cost of Projects", x='shortName', y='totalCost_x', text='nProjects', labels={'shortName':'Organization','totalCost_x':'Total Cost of Projects'})
    coordinatorCostsGraph.update_traces(marker_color='green')
    allCostsGraph = px.bar(allCosts.head(), title="All Companies Total Cost of Projects", x='shortName', y='totalCost_x', text='nProjects', labels={'shortName':'Organization','totalCost_x':'Total Cost of Projects'})

    def reformat_projects(df):
      df['totalCost'] = df['totalCost'].apply(lambda x: f"{x:,.0f}")
      df['id'] = df['id'].apply(lambda x: f'<a href="https://cordis.europa.eu/project/id/{x}">{x}</a>')
      df.rename(columns = {'id':'ID', 'acronym':'Short Title', 'title': 'Project Title', 'totalCost':'Total Cost (Euro)', 'percentEUFunded': 'EU Funding Ratio'}, inplace = True)
      return df

    # format dataframes for html report
    topRelevanceProjects = reformat_projects(topRelevanceProjects)
    topCostProjects = reformat_projects(topCostProjects)
    coordinatorCounts.rename(columns = {'name': 'Coordinator Name', 'nProjects': 'Number of Projects Coordinated', 'shortName': 'Short Name', 'country':'Country'}, inplace = True)
    allCounts.rename(columns = {'name': 'Organisation Name', 'nProjects': 'Number of Projects Involved', 'shortName': 'Short Name', 'country':'Country'}, inplace = True)
    coordinatorCosts['totalCost_x'] = coordinatorCosts['totalCost_x'].apply(lambda x: f"{x:,.0f}")
    coordinatorCosts.rename(columns = {'name': 'Coordinator Name', 'nProjects': 'Number of Projects Coordinated', 'shortName': 'Short Name', 'country':'Country', 'totalCost_x':'Total  (Euro)'}, inplace = True)
    allCosts['totalCost_x'] = allCosts['totalCost_x'].apply(lambda x: f"{x:,.0f}")
    allCosts.rename(columns = {'name': 'Organisation Name', 'nProjects': 'Number of Projects Involved', 'shortName': 'Short Name', 'country':'Country', 'totalCost_x':'Total Cost (Euro)'}, inplace = True)

    # Creates html report
    print("Generating html report...")
    page_title = "Keyword Report"
    if isUnion:
      title_text="Results for '" + " OR ".join(keywordList) + "' in CORDIS Horizon 2020"
    else:
      title_text="Results for '" + " AND ".join(keywordList) + "' in CORDIS Horizon 2020"
    if isPillar2:
      title_text=title_text+" (Pillar 2)"
    else:
      title_text=title_text+" (All Pillars)"
    text="Total number of results: " + str(len(projectList.index))
    projects_text='Most relevant projects by frequency of keyword(s)'
    projects_subtitle='These projects used the keywords most frequently in their title and objective statements.'
    projects_text_cost='Most costly projects with keyword(s)'
    projects_text_cost_subtitle='These projects were the most costly out of the projects that contained the keyword(s).'
    total_text='Total cost of all projects: &euro;'
    ec_text='Max contributions by the EC: &euro;'
    keywords_text ='Other prominent keywords from the selected projects'
    commonCompanies_text = 'Most involved companies by number of projects with keyword(s)'
    costlyCompanies_text = 'Most involved companies by total cost of projects with keyword(s)'
    coordinatorCounts_subtext = 'Full data can be found in the "nProject (coordinators)" sheet.'
    allCounts_subtext = 'Full data can be found in the "nProject (all)" sheet.'
    coordinatorCosts_subtext = 'Full data can be found in the "by cost (coordinators)" sheet.'
    allCosts_subtext = 'Full data can be found in the "by cost (all)" sheet.'

    html = f'''
        <html>
          <head>
              <title>{page_title}</title>
              <link rel="stylesheet" href="reportStyle.css" type="text/css" media="all">
              <meta charset="UTF-8">
              <script src="https://cdn.plot.ly/plotly-2.12.1.min.js"></script>
          </head>
          <body>
            <div id="title">
              <h1>{title_text}</h1>
                <p>{text}</p>
            </div>
            <div id="topRelevance">
              <h2>{projects_text}</h2>
                <p>{projects_subtitle}</p>
                {topRelevanceProjects.to_html(index=False, escape=False, justify="center", classes='table', table_id="topProjects", columns=['ID', 'Short Title', 'Project Title', 'Total Cost (Euro)', 'EU Funding Ratio'])}
                {relevantProjectCounts.to_html(full_html=False, include_plotlyjs=False, default_width='50%', default_height='65%', div_id='relevantProjGraph')}
            </div>
            <div id="topCost">
              <h2>{projects_text_cost}</h2>
                <p>{projects_text_cost_subtitle}</p>
                {topCostProjects.to_html(index=False, escape=False, justify="center", classes='table', table_id="costProjects", columns=['ID', 'Short Title', 'Project Title', 'Total Cost (Euro)', 'EU Funding Ratio'])}
                {costlyProjectCounts.to_html(full_html=False, include_plotlyjs=False, default_width='50%', default_height='65%', div_id='costProjGraph')}
            </div>
            <div id ="parent" class="clearfix">
              <div id="cost">
                <h2>{total_text}{sumCost}</h2>
                <h2>{ec_text}{sumEC}</h2>
                {figCosts.to_html(full_html=False, include_plotlyjs=False, default_width='70%', default_height='65%', div_id='costsPieGraph')}
              </div>
              <div id ="keywords">
                <h2>{keywords_text}</h2>
                {", ".join(nlp_words)}
              </div>
            </div>
            <div id="commonCompanies">
              <h2>{commonCompanies_text}</h2>
                <h3>As Coordinators</h3>
                  {coordinatorCounts.head(5).to_html(index=False, justify="center", classes='table', table_id='coordinatorCounts')}
                  <p class="referText">{coordinatorCounts_subtext}</p>
                  {coordinatorCountsGraph.to_html(full_html=False, include_plotlyjs=False, default_width='50%', default_height='65%', div_id='coordinatorCountsGraph')}
                  
                <h3>All</h3>
                  {allCounts.head(5).to_html(index=False, justify="center", classes='table', table_id='allCounts')}
                  <p class="referText">{allCounts_subtext}</p>
                  {allCountsGraph.to_html(full_html=False, include_plotlyjs=False, default_width='50%', default_height='65%', div_id='allCountsGraph')}
                  
            </div>
            <div id="companyCosts">
              <h2>{costlyCompanies_text}</h2>
                <div id="coordinatorCostsContainer">
                  <h3>As Coordinators</h3>
                    {coordinatorCosts.head().to_html(index=False, justify="center", classes='table', table_id='coordinatorCosts')}
                    <p class="referText">{coordinatorCosts_subtext}</p>
                    {coordinatorCostsGraph.to_html(full_html=False, include_plotlyjs=False, default_width='50%', default_height='65%', div_id='coordinatorCostsGraph')}
                </div>
                <div id="allCostsContainer"> 
                  <h3>All</h3>
                  {allCosts.head().to_html(index=False, justify="center", classes='table', table_id='allCosts')}
                  <p class="referText">{allCosts_subtext}</p>
                  {allCostsGraph.to_html(full_html=False, include_plotlyjs=False, default_width='50%', default_height='65%', div_id='allCostsGraph')}
                  
                </div>
            </div>
          </body>
        </html>
    '''

    with open('html_report.html', 'w') as f:
        f.write(html)
    print("Html report complete!")