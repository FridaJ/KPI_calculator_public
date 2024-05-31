#!/usr/bin/env python

def get_kpis():

    import fmrest
    import pandas as pd
    import datetime
    import numpy as np
    #import logging # for creating log file
    #from contextlib import redirect_stdout
    #import io
    
    ##### Using REST-api to get data from the FileMaker database
    
    fms_application = fmrest.Server('https://fmapp19-p.gu.gu.se',
                       user = 'api',
                       password = 'pythonR3st',
                       database = 'pps_data',
                       layout = 'application',
                       api_version = 'vLatest')
    
    fms_project = fmrest.Server('https://fmapp19-p.gu.gu.se',
                       user = 'api',
                       password = 'pythonR3st',
                       database = 'pps_data',
                       layout = 'project',
                       api_version = 'vLatest')
    
    fms_application.login()
    fms_project.login()
    
    foundset_appl = fms_application.get_records(limit=10000)
    foundset_proj = fms_project.get_records(limit=20000)
    
    fms_application.logout()
    fms_project.logout()
    
    ##### Preprocessing
    
    appl_df = foundset_appl.to_df()
    proj_df = foundset_proj.to_df()
    proj_df.columns = ['project::' + colname for colname in proj_df.columns ]
    
    proj_df['ApplicationID'] = [ id[:-1] for id in proj_df['project::ProjectID'] ]
    pps_data = pd.merge(appl_df, proj_df, how='outer', on='ApplicationID')
    
    pps_data.replace({'\r': '\n'}, regex=True, inplace=True)
    pps_data.replace(r'^\s*$', np.nan, regex=True, inplace=True)
    pps_data_raw = pps_data
    
    # change type for batchDates, application_Year and application_Date
    pps_data_raw = pps_data_raw.convert_dtypes()
    pps_data_raw = pps_data_raw.infer_objects()
    
    convert_dict = {'ApplicationYear': 'int', 'project::LastBatch': 'datetime64[ns]', 'ApplicationDate': 'datetime64[ns]', 'project::projectFinishedDate': 'datetime64[ns]'}
    pps_data_raw = pps_data_raw.astype(convert_dict)
    
    # Fill ApplicationID for subprojects
    # cols = all columns not starting with "project::"
    cols = [colname for colname in pps_data_raw.columns if not colname.startswith("project::")]
    for id in range(len(pps_data_raw.ApplicationID)):
            if pd.isna(pps_data_raw.ApplicationID[id]):
                  pps_data_raw.loc[id, cols] = pps_data_raw.loc[id-1, cols]
    
    pps_data = pps_data_raw.sort_values(['PPS_lab', 'ApplicationID']).reset_index()
    pps_data = pps_data.replace(r"^ +| +$", r"", regex=True) # to remove all spaces from beginning and end of strings
    
    affiliations = {'gu.se': 'Gothenburg University', 'lu.se': 'Lund University', 'liu.se': 'Linköping University', 'ki.se': 'KI', 'kth.se':'KTH', 'su.se':'Stockholm University', 'uu.se':'Uppsala University', 'umu.se':'Umeå University', 'chalmers.se':'Chalmers University', 'slu.se':'SLU'}
    
    # From here, log to save output to log file
    # Set up an in-memory string buffer
    #f = io.StringIO()
    
    # Redirect stdout to the buffer
    #with redirect_stdout(f):
    
    # Correct wrong data and fill missing data
    
    affiliations = {'gu.se': 'Gothenburg University', 'lu.se': 'Lund University', 'liu.se': 'Linköping University', 'ki.se': 'KI', 'kth.se':'KTH', 'su.se':'Stockholm University', 'uu.se':'Uppsala University', 'umu.se':'Umeå University', 'chalmers.se':'Chalmers University', 'slu.se':'SLU'}
    
    # Change gender F to female and M to male
    # (Check names for spaces, missing letters, first + last not separated)
    
    # Find gender values to change:
    
    for id in range(len(pps_data['Applicant_Gender'])):
        if pd.notna(pps_data.loc[id, 'Applicant_Gender']):
            if pps_data.loc[id, 'Applicant_Gender'] == 'F':
                    # Set data values with df.loc
                    pps_data.loc[id, 'Applicant_Gender'] = 'Female'
            elif pps_data.loc[id, 'Applicant_Gender'] == 'M':
                    pps_data.loc[id, 'Applicant_Gender'] = 'Male'
    
    for id in range(len(pps_data['Investigator_Gender'])):
        if pd.notna(pps_data.loc[id, 'Investigator_Gender']):
            if pps_data.loc[id, 'Investigator_Gender'] == 'F':
                    # Set data values with df.loc
                    pps_data.loc[id, 'Investigator_Gender'] = 'Female'
            elif pps_data.loc[id, 'Investigator_Gender'] == 'M':
                    pps_data.loc[id, 'Investigator_Gender'] = 'Male'
    
    # Find Country values to change:
    for id in range(len(pps_data['Investigator_Country'])):
        if pd.notna(pps_data.loc[id, 'Investigator_Country']):
            if pps_data.loc[id, 'Investigator_Country'] == 'Sverige':
                    # Set data values with df.loc
                    pps_data.loc[id, 'Investigator_Country'] = 'Sweden'
        if pd.notna(pps_data.loc[id, 'Applicant_Country']):
            if pps_data.loc[id, 'Applicant_Country'] == 'Sverige':
                    # Set data values with df.loc
                    pps_data.loc[id, 'Investigator_Country'] = 'Sweden'
    
    # If there are any null values here, fill them or make lists for user to set value in database
    
    no_data_lab = pps_data[pd.isna(pps_data['PPS_lab'])]
    no_data_aff = pps_data[pd.isna(pps_data['Investigator_PrimaryAffiliation'])]
    no_data_aff_name = pps_data[pd.isna(pps_data['Investigator_University_Company'])]
    no_data_country = pps_data[pd.isna(pps_data['Investigator_Country'])]
    no_data_gender = pps_data[pd.isna(pps_data['Investigator_Gender'])]
    no_data_status = pps_data[pd.isna(pps_data['TotalStatusText'])]
    
    if not no_data_lab.empty:
        print('The following projects have not yet been assigned to a PPS lab, and will not be counted:\n')
        no_data_lab = no_data_lab.drop_duplicates(['ApplicationID'])
        print(no_data_lab.loc[:,['ApplicationID', 'project::PPScontact']].to_string(index=False)) 
        print("\n---------------------------------------------------------------------------\n")
    else:
        print('All projects have been assigned to a PPS lab.')
        print("\n---------------------------------------------------------------------------\n")
    pps_data.dropna(subset=['PPS_lab'], inplace=True)
    
    if not no_data_aff.empty:
        print("The following labs have projects missing the affiliation type for the PI:\n")
        no_data_aff_grouped = no_data_aff.groupby(['PPS_lab'])['ApplicationID'].count()
        print(no_data_aff_grouped.to_string())
    
        # check if academic affiliation, if so set the field
        print('\nChecking for academic affiliations...')
        print('\nSetting affiliation type to \'Academic\'...')
        for id in range(len(no_data_aff['ApplicationID'])):
            aff_name = ''
            if not pd.isna(no_data_aff.iloc[id,]['Investigator_University_Company']):
                aff_name = no_data_aff.iloc[id,]['Investigator_University_Company']
            elif not pd.isna(no_data_aff.iloc[id,]['Applicant_University_Company']):
                if no_data_aff.iloc[id,]['Applicant_Name'] == no_data_aff.iloc[id,]['Investigator_Name']:
                    aff_name = no_data_aff.iloc[id,]['Applicant_University_Company']
            if aff_name in affiliations.values():
                app_id = no_data_aff.iloc[id,]['ApplicationID']
                #print(f"Setting affiliation type to Academic for {app_id} because affiliation is {aff_name}.")
                pps_data.loc[pps_data.ApplicationID == app_id, 'Investigator_PrimaryAffiliation'] = 'Academic'
            #if not pd.isna(no_data_aff.iloc[id,]['Investigator_Email']):
            #    if any(key in no_data_aff.iloc[id,]['Investigator_Email'] for key in affiliations.keys()):
            #        print('Found email in dict.')
    
        no_data_aff_after_fill = pps_data[pd.isna(pps_data['Investigator_PrimaryAffiliation'])]
        if not no_data_aff_after_fill.empty:
            print("\nHere is information for projects that still have missing affiliation type:\n")
            for aff in range(len(no_data_aff_after_fill['ApplicationID'])):
                print(no_data_aff_after_fill.iloc[aff,][['ApplicationID','Applicant_Email', 'Investigator_Email', 'Investigator_University_Company', 'Applicant_University_Company']])    
                print("--------------------------------------")
            print('These projects will not be counted, please correct the affiliation type in the database.')
            pps_data.dropna(subset=['Investigator_PrimaryAffiliation'], inplace=True)
        else: 
            print('\nAll projects with missing affiliation type have now been fixed.')
    else:
        print('All projects have data for affiliation type.')
    
    print("\n---------------------------------------------------------------------------\n")
    
    if not no_data_aff_name.empty:
        print("\nThe following labs have projects missing the affiliation name of the PI:")
        no_data_aff_name_grouped = no_data_aff_name.groupby(['PPS_lab'])['ApplicationID'].count()
        print(no_data_aff_name_grouped.to_string())
    
        # printing info about projects with missing affiliation name:
        print("\nHere is information for projects with missing affiliation name:\n")
        for aff in range(len(no_data_aff_name['ApplicationID'])):
            print(no_data_aff_name.iloc[aff,][['ApplicationID', 'Applicant_Name', 'Applicant_Email', 'Investigator_Name', 'Investigator_Email', 'project::PPScontact']])    
            print("----------------------------------------")
        print('These projects will not be counted, please correct the affiliation name in the database.')
        pps_data.dropna(subset=['Investigator_University_Company'], inplace=True)
    else:
        print('All projects have data for affiliation name.')
    
    print("\n---------------------------------------------------------------------------\n")
    
    if not no_data_country.empty:
            print("\nThe following labs have projects missing the country of the PI:\n")
            no_data_country_grouped = no_data_country.groupby(['PPS_lab'])['ApplicationID'].count()
            print(no_data_country_grouped.to_string())
            #print(no_data_country[['PPS_lab', 'ApplicationID']])
    
            print("\nFilling missing data based on name...")
            for id in range(len(no_data_country['ApplicationID'])):
                if no_data_country.iloc[id,]['Applicant_Name'] == no_data_country.iloc[id,]['Investigator_Name']:
                    app_id = no_data_country.iloc[id,]['ApplicationID']
                    pps_data.loc[pps_data.ApplicationID == app_id, 'Investigator_Country'] = \
                    no_data_country.iloc[id,]['Applicant_Country']
    
            # Check which entries still have missing data for Country and fill by email:
            no_data_country_after_fill = pps_data[pd.isna(pps_data['Investigator_Country'])]
            print("\nFilling missing data based on email...")
            for id in range(len(pps_data['ApplicationID'])):
                if pd.isna(pps_data.Investigator_Country[id]):
                    pi_email = pps_data.Investigator_Email[id]
                    if isinstance(pi_email, str) and pi_email.endswith(".se"):
                        app_id = pps_data.ApplicationID[id]
                        #print(f"Changing Investigator Country for {app_id} to \"Sweden\".")
                        pps_data.loc[pps_data['ApplicationID'] == app_id, 'Investigator_Country'] = "Sweden"
    
            no_data_country_after_fill = pps_data[pd.isna(pps_data['Investigator_Country'])]
            if not no_data_country_after_fill.empty:
                print(f"\nThere are still {len(no_data_country_after_fill.ApplicationID)} subprojects with missing country:")
                print('These projects will not be counted, please correct the country in the database.')
                pps_data.dropna(subset=['Investigator_Country'], inplace=True)
                print(no_data_country_after_fill[['ApplicationID', 'PPS_lab', 'Investigator_Name', 'Investigator_Email']].to_string(index=False))
            else:
                print('\nAll projects with missing PI country have now been fixed.')
    else:
        print('All projects have data for PI country.')
    
    print("\n---------------------------------------------------------------------------\n")
    
    if not no_data_gender.empty:
            print("\nThe following labs have projects missing the gender of the PI:\n")
            no_data_gender_grouped = no_data_gender.groupby(['PPS_lab'])['ApplicationID'].count()
            print(no_data_gender_grouped.to_string())
    
            # fill PI gender with Appl. gender if names are the same:
            print("\nSetting PI gender to the value of applicant gender if names are the same...")
            for id in range(len(no_data_gender['ApplicationID'])):
                if not pd.isna(no_data_gender.iloc[id,]['Applicant_Gender']):
                    if no_data_gender.iloc[id,]['Applicant_Name'] == no_data_gender.iloc[id,]['Investigator_Name']:
                            app_id = no_data_gender.iloc[id,]['ApplicationID']
                            pps_data.loc[pps_data.ApplicationID == app_id, 'Investigator_Gender'] = no_data_gender.iloc[id,]['Applicant_Gender']
    
            no_data_gender_after_fill = pps_data[pd.isna(pps_data['Investigator_Gender'])]
            if not no_data_gender_after_fill.empty:
                print(f"\nThere are still {len(no_data_gender_after_fill.ApplicationID)} subprojects with missing PI gender:")
                print('These projects will not be counted, please correct the gender in the database.\n')
                pps_data.dropna(subset=['Investigator_Country'], inplace=True)
                print(no_data_gender_after_fill[['ApplicationID', 'Applicant_Name', 'PPS_lab']].to_string(index=False))
            else:
                print('\nAll projects with missing PI gender have now been fixed.')
    else:
        print('All projects have data for PI gender!')
    
    print("\n---------------------------------------------------------------------------\n")
    
    if not no_data_status.empty:
            #print("\nThe following contacts have projects missing a project status:\n")
            #no_data_status_grouped = no_data_status.groupby(['project::PPScontact', 'ApplicationID'])['ApplicationID'].count()
            #print(no_data_status_grouped)
    
            # Set project status to 'Not started' for projects (subprojects) with missing status:
            print(f"Setting project status to \"Not started\" for the following projects with missing status:\n")
            for id in range(len(no_data_status['project::ProjectID'])):
                    app_id = no_data_status.iloc[id,]['project::ProjectID']
                    print(app_id + ', with PPS contact: ' + str(no_data_status.iloc[id,]['project::PPScontact']))
                    pps_data.loc[pps_data.ApplicationID == app_id, 'project::Status'] = "Not started"
                    #print("----------------------------------------")
            print('\nAll projects with missing status have now been set to \'Not started\'.')
    else:
        print('All projects have data for project status.')
    
    # Get the captured output from the buffer
    #out = f.getvalue()
    # Write the captured output to the log file
    #logging.info(out)
    
    #   Assign new category Academic/Non-academic
    
    pps_data['Academic'] = pps_data['Investigator_PrimaryAffiliation']
    
    for id in range(len(pps_data['ApplicationID'])):
        aff = pps_data.iloc[id,]['Investigator_PrimaryAffiliation']
        if aff == 'Academic':
            pps_data.iat[id, pps_data.columns.get_loc('Academic')] = 'Academic'
        else:
            pps_data.iat[id, pps_data.columns.get_loc('Academic')] = 'Non-Academic'
    
    #  Assign new category Swedish/Non-Swedish
    
    pps_data['Swedish'] = pps_data['Investigator_Country']
    
    for id in range(len(pps_data['ApplicationID'])):
        country = pps_data.iloc[id,]['Investigator_Country'].lower().strip()
        if country == 'sweden':
            pps_data.iat[id, pps_data.columns.get_loc('Swedish')]= 'Swedish'
        else:
            pps_data.iat[id, pps_data.columns.get_loc('Swedish')] = 'Non-Swedish'
       
    
    ##### Functions
    
    # function to get the number of applied projects
    
    def get_applied(year, df=pps_data):
        applied_projects = df[df['ApplicationYear'] == year].reset_index(drop=True)
        return applied_projects
    
    def get_applied_count(year, df=pps_data):
        applied_projects = df[df['ApplicationYear'] == year].reset_index(drop=True)
        count = len(applied_projects.drop_duplicates(subset='ApplicationID'))
        return count
    
    #Functions to get df and count for finished projects:
    
    def get_finished(year, df=pps_data):
        finished_projects = df.loc[df['project::projectFinishedDate'].dt.year == year,].reset_index(drop=True)
        finished_projects = finished_projects.loc[finished_projects['TotalStatusText'] == 'Finished',].reset_index(drop=True)
        #if finished_projects.empty:
        #    print(f"There are no projects finished for in {year}.")
        return finished_projects
    
    def get_finished_count(year, df=pps_data):
        finished_projects = get_finished(year, df)
        count = len(finished_projects.drop_duplicates(subset='ApplicationID'))
        #if finished_projects.empty:
        #    print(f"There are no projects finished in {year}.")
        return count
    
    # Function to get the finished date from the latest finished project in an application:
    
    def get_latest(app_id, df):
        project_df = df.loc[df['ApplicationID'] == app_id,]
        date_list = []
        for i, row in project_df.iterrows():
            status = row['project::Status']
            if status == 'Finished':      
            #if status in ['Not started', 'Ongoing', 'Finished']:
                date = row['project::projectFinishedDate'].date()
                if not pd.isnull(date):
                    date_list.append(date)
                    if date_list:
                        latest = max(date_list)
                        return latest
                    else:
                        return 'No date'
    
    def find_with_finished_date(status, show_without_date, df):
        df = df[~pd.isna(df['project::projectFinishedDate'])].reset_index(drop=True) 
        # ~ means exclude rows, like "not"! use drop=True to avoid level 0 valueerror because of two index columns
        finished_list = []
        last_id = ""
        for i in df.index:
            total_status = df.iloc[i,]['TotalStatusText']
            app_id = df.iloc[i,]['ApplicationID']
            #print(app_id, total_status)
            if total_status == status:
                pps_lab = df.iloc[i,]['PPS_lab']
                finished_date = get_latest(app_id, df)
                if app_id == last_id:
                    continue
                else:
                    #print(f'Total status: {total_status}')
                    #print(finished_date)
                    finished_list.append((app_id, pps_lab, finished_date))
                    #print(f'The project belongs to {pps_lab}.\n')
            last_id = app_id
        finished_df = pd.DataFrame(finished_list, columns =['ApplicationID', 'PPS_lab', 'FinishedDate'])
        finished_df = finished_df.drop_duplicates(subset='ApplicationID').reset_index(drop=True)
        #print(finished_df)
        return finished_df.sort_values('FinishedDate')
        
    def print_with_finished_date(status, show_without_date, df):
        finished_projects = find_with_finished_date(status, show_without_date, df)
        print(f'\nThere are {len(finished_projects)} projects marked as {status}:\n')
        print(finished_projects.to_markdown(index=False))
    
    def get_female_count(df=pps_data):
        df = df.drop_duplicates(subset=['ApplicationID'])
        all_count = len(df)
        df_females = df.loc[df['Investigator_Gender'] == 'Female', ]
        females_count = len(df_females)
        #print(f'There are {females_count} female PIs in a total of {all_count} projects.')
        return females_count
        
    
    #Function to get the affiliation type count on any df
    
    def get_affiliation_type_count(df=pps_data):
        df = df.drop_duplicates(subset=['ApplicationID'])
        aca = 0
        public = 0
        comm = 0
        other = 0
        for aff in df['Investigator_PrimaryAffiliation']:
            if 'commercial entity' in aff.lower():
                comm += 1
            elif 'public sector' in aff.lower():
                public += 1
            elif aff == 'Academic':
                aca += 1
            elif aff == 'Other':
                other += 1
        if len(df) != aca+public+comm+other:
            print('Something is wrong, the numbers do not add up correctly.')
        return aca, public, comm, other
    
    # New function to get the host-type of an academic project on any df, after changes from VR 2024
    
    # Make a new column for type of host: 'Own', 'PPS', 'Swedish', "International"
    # Definition of 'Own' is if the PI university is the same as the PPS lab UNIVERSITY
    
    def add_host_column_new(df=pps_data):
        hosts_labs = {'Gothenburg University': ['MPE', 'SNC', 'YPP'], 'Lund University': ['LP3 3.3', 'LP3 4.1', 'LP3 4.2'], 'KI': ['PSF'], 'KTH': ['KTH'], 'Umeå University': ['PEP 3.1', 'PEP 3.5']}
        df = df.reset_index(drop=True)
        df.loc[:,'Host_type'] = df['Investigator_University_Company']
    
        for id in range(len(df['Host_type'])):
            aff = df['Host_type'][id].strip()
            swe = df.loc[id, 'Swedish']
            if aff in hosts_labs.keys() and df.loc[id, 'PPS_lab'] in hosts_labs[aff]:
                df.iat[id, df.columns.get_loc('Host_type')]= 'Own'
            elif aff in hosts_labs.keys():
                df.iat[id, df.columns.get_loc('Host_type')]= 'PPS'
            elif swe == 'Non-Swedish':
                df.iat[id, df.columns.get_loc('Host_type')]= 'International'
            else:
                df.iat[id, df.columns.get_loc('Host_type')]= 'Swedish'
        return df
    
    # Function to get the host type count on any df, only academic records
    
    def get_host_type_count(df=pps_data): # changed 1/2 2024
        df = df.drop_duplicates(subset=['ApplicationID'])
        own = 0
        pps = 0
        swe = 0
        inter = 0
        for i in range(len(df['Host_type'])):
            if df.iloc[i, ]['Academic'] == 'Academic':
                host = df.iloc[i, ]['Host_type']
                if host == 'Own':
                    own += 1
                elif host == 'PPS':
                    pps += 1
                elif host == 'Swedish':
                    swe += 1
                elif host == 'International':
                    inter += 1
            #if len(df) != aca+public+comm+other:
            #    print('Something is wrong, the numbers do not add up correctly.')
        return own, pps, swe, inter
    
    # make dataframes from the lists of lists that are the tables in get_kpis. 
    # then use concat to add them together to a df that looks like in the kpi report.
    # finally, export these dfs to excel, each year on a separate sheet
    
    def kpis_to_df(year, pps_data=pps_data, finished=False, uniquePI=False):
        
        pps_data_year = pps_data[pps_data['ApplicationYear'] == year]
        
        #labs = list(set(pps_data_year['PPS_lab']))
        #labs.sort()
        labs = ['LP3 3.3', 'LP3 4.1', 'LP3 4.2', 'MPE', 'SNC', 'YPP', 'PSF', 'KTH', 'PEP 3.1', 'PEP 3.5']
        lab_dfs = {}
        
        if finished:
            pps_data_year = get_finished(year, pps_data) #changed pps_data_year to pps_data 23/1-24
            
        pps_data_year = add_host_column_new(pps_data_year) # New function 1/2 2024
        pps_data_year_unique = pps_data_year.drop_duplicates(subset=['ApplicationID'])
          
        for lab in labs:
            lab_df = pps_data_year_unique.loc[pps_data_year_unique.PPS_lab == lab,]
            if uniquePI:
                lab_df = lab_df.drop_duplicates(subset='Investigator_Name')
            lab_dfs[lab] = lab_df
        
        # Make df that shows PI gender count, per lab
        
        print_list = []
    
        for lab, lab_df in lab_dfs.items():
            women = get_female_count(lab_df)
            men = len(lab_df) - women
            print_list.append([lab, len(lab_df), men, women])
    
        if uniquePI:
            pps_data_year_unique = pps_data_year_unique.drop_duplicates(subset=['Investigator_Name'])
            total_pis = len(pps_data_year_unique)
            women = get_female_count(pps_data_year_unique)
            men = total_pis - women
            print_list.append(['All', total_pis, men, women])
        else:
            # get totals from adding the element lists together
            zip_list = list(zip(*print_list))
            total = [sum(zip_tuple) for zip_tuple in zip_list[1:]] # the three numbers needed for total
            all_proj = total[0]
            men = total[1]
            women = total[2]
            print_list.append(['All', all_proj, men, women])
        
        df_gender = pd.DataFrame(print_list, columns=['PPS Lab', 'Total', 'Men', 'Women']).set_index('PPS Lab')
        
        # Make df that shows PI affiliation type count, per lab
    
        print_list = []
    
        for lab, lab_df in lab_dfs.items():
            a, p, c, o = get_affiliation_type_count(lab_df)
            print_list.append([lab, a, c, p, o]) # changed the order
    
        aca, pub, comm, oth = get_affiliation_type_count(pps_data_year_unique)
        print_list.append(['All', aca, comm, pub, oth]) # same order as for lab_dfs
    
        df_afftype = pd.DataFrame(print_list, columns=['PPS Lab', 'Academic', 'Commercial', 'Public', 'Other']).set_index('PPS Lab')
        # use all but the academic column for concat
        
        # Academic affiliation should be divided into genders
    
        print_list = []
    
        for lab, lab_df in lab_dfs.items():
            #print(lab, lab_df.ApplicationID, lab_df.Host_type)
            df_academic = lab_df.loc[lab_df.Academic == 'Academic',].reset_index(drop=True)
            women = get_female_count(df_academic)
            men = len(df_academic) - women
            print_list.append([lab, len(df_academic), men, women])
    
        pps_data_academic = pps_data_year_unique.loc[pps_data_year_unique.Academic == 'Academic',].reset_index(drop=True)
        women = get_female_count(pps_data_academic)
        men = len(pps_data_academic) - women
        print_list.append(['All', len(pps_data_academic), men, women])
    
        df_aca_gender = pd.DataFrame(print_list, columns=['PPS Lab', 'Academic', 'Men Academic', 'Women Academic']).set_index('PPS Lab')
    
    
        # Make a table that shows PI host type count, per lab
    
        print_list = []
    
        for lab, lab_df in lab_dfs.items():
            lab_df_academic = lab_df[lab_df['Academic'] == 'Academic']
            own, pps, swe, inter = get_host_type_count(lab_df_academic)
            print_list.append([lab, own, pps, swe, inter])
    
        own = pps_data_academic.loc[pps_data_academic.Host_type == 'Own',]['ApplicationID'].count()
        pps = pps_data_academic.loc[pps_data_academic.Host_type == 'PPS',]['ApplicationID'].count()
        swe = pps_data_academic.loc[pps_data_academic.Host_type == 'Swedish',]['ApplicationID'].count()
        inter = pps_data_academic.loc[pps_data_academic.Host_type == 'International',]['ApplicationID'].count()
        print_list.append(['All', own, pps, swe, inter])
    
        df_host = pd.DataFrame(print_list, columns=['PPS Lab', 'Own', 'PPS', 'Other Swedish uni', 'International']).set_index('PPS Lab')
        
        # use concat to merge the four dataframes into one:
        
        df_all = pd.concat([df_gender, df_aca_gender, df_afftype.iloc[:,1:], df_host], axis=1).reset_index()
        
        df_all['PPS Lab'] = ['3.3 LP3', '4.1 LP3', '4.2 LP3', '3.4.A MP3', '3.6 SNC', '3.2 YPP', '3.1.A PSF', '3.4.B KTH', '3.1.B PEP', '3.5 PEP', '0 All']
        df_all = df_all.sort_values('PPS Lab')
        
        return df_all
    
    # Function to calculate number of batches delivered during a specific year, per lab
    
    def get_batches(year, df=pps_data):
        # Select all subprojects with a batch from the specified year:
        pps_data_year = df.loc[df['project::BatchDates'].str.contains(f'{year}', na=False),].reset_index(drop=True)
    
        #labs = list(set(pps_data_year['PPS_lab']))
        #labs.sort()
        labs = ['PSF', 'PEP 3.1', 'YPP', 'LP3 3.3', 'MPE', 'KTH', 'PEP 3.5', 'SNC', 'LP3 4.1', 'LP3 4.2']
        lab_dfs = {}
        for lab in labs:
            lab_df = pps_data_year.loc[pps_data_year.PPS_lab == lab,].reset_index()
            # Count the number of batches for each subproject for the specified year. 
            # Save as new column 'Batches_Year'
            lab_df['Batches_Year'] = lab_df['project::BatchDates'].str.count(f'{year}')
            lab_dfs[lab] = lab_df
    
        batches = []
        for lab, df in lab_dfs.items():
            batches_tot = df['Batches_Year'].sum()
            female_df = df[df['Investigator_Gender'] == 'Female']
            female_batches = female_df['Batches_Year'].sum()
            male_batches =  batches_tot - female_batches
            batches.append((lab, batches_tot, male_batches, female_batches))
    
        df_to_print = pd.DataFrame(batches, columns=['PPS_lab', 'Total', 'Male', 'Female'])
        return df_to_print
    
    def kpis_to_excel(filename='kpis_VR.xlsx'):
        
        import xlsxwriter
        
        years = sorted([int(year) for year in list(set(pps_data.ApplicationYear))]) # unique years ascending
        
        # make excel file with first year on sheet1:
        with pd.ExcelWriter(filename, engine='xlsxwriter') as writer:  
    
            for year in years:
                
                sheet = str(year)
                row = 0
                
                df1 = kpis_to_df(year, pps_data, finished=False, uniquePI=False)
                df2 = kpis_to_df(year, pps_data, finished=True, uniquePI=False)
                df3 = kpis_to_df(year, pps_data, finished=False, uniquePI=True)
                df4 = kpis_to_df(year, pps_data, finished=True, uniquePI=True)
                df5 = get_batches(year, pps_data)
    
                df1.to_excel(writer, sheet_name=sheet, startrow=1, index=False)
                workbook  = writer.book
                worksheet = writer.sheets[sheet]
                worksheet.write(row, 0, f'Applied PPS projects {year}:')
                df2.to_excel(writer, sheet_name=sheet, startrow=row+15, index=False)
                worksheet.write(row+14, 0, f'Finished PPS projects {year}:')
                df3.to_excel(writer, sheet_name=sheet, startrow=row+30, index=False)
                worksheet.write(row+29, 0, f'Applied PPS projects {year}, unique PI\'s (Note that the total number for PPS may differ from the sum of the lab projects.):')
                df4.to_excel(writer, sheet_name=sheet, startrow=row+45, index=False)
                worksheet.write(row+44, 0, f'Finished PPS projects {year}, unique PI\'s (Note that the total number for PPS may differ from the sum of the lab projects.):')
                df5.to_excel(writer, sheet_name=sheet, startrow=row+60, index=False)
                worksheet.write(row+59, 0, f'Delivered batches in {year}:')
    
                for lab in ['PSF', 'PEP 3.1', 'YPP', 'LP3 3.3', 'MPE', 'KTH', 'PEP 3.5', 'SNC', 'LP3 4.1', 'LP3 4.2']:
    
                    row += 75
                    
                    pps_data_lab = pps_data[pps_data['PPS_lab'] == lab]
                    
                    df1 = kpis_to_df(year, pps_data_lab, finished=False, uniquePI=False)
                    df2 = kpis_to_df(year, pps_data_lab, finished=True, uniquePI=False)
                    df3 = kpis_to_df(year, pps_data_lab, finished=False, uniquePI=True)
                    df4 = kpis_to_df(year, pps_data_lab, finished=True, uniquePI=True)
                    df5 = get_batches(year, pps_data_lab)
                
                    df1.to_excel(writer, sheet_name=sheet, startrow=row+1, index=False)
                    worksheet.write(row-1, 0, f'{lab} {year}:')
                    worksheet.write(row, 0, f'Applied PPS projects {year}:')
                    df2.to_excel(writer, sheet_name=sheet, startrow=row+15, index=False)
                    worksheet.write(row+14, 0, f'Finished PPS projects {year}:')
                    df3.to_excel(writer, sheet_name=sheet, startrow=row+30, index=False)
                    worksheet.write(row+29, 0, f'Applied PPS projects {year}, unique PI\'s (Note that the total number for PPS may differ from the sum of the lab projects.):')
                    df4.to_excel(writer, sheet_name=sheet, startrow=row+45, index=False)
                    worksheet.write(row+44, 0, f'Finished PPS projects {year}, unique PI\'s (Note that the total number for PPS may differ from the sum of the lab projects.):')
                    df5.to_excel(writer, sheet_name=sheet, startrow=row+60, index=False)
                    worksheet.write(row+59, 0, f'Delivered batches in {year}:')
    
    kpis_to_excel()
    
    
    
    
        
    
    
    
    
    
    
    
    
