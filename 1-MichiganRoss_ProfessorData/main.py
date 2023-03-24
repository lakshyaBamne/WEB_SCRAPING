# Main URL to scrape => https://michiganross.umich.edu/faculty-research/directory
# this website can be scraped by using a very specific query sent to the server

# => root_url + '?view=list' + f'&last={letter}' + '&status=Active'

#######################################################################################
# Importing required libraries

import pandas as pd
import requests
from bs4 import BeautifulSoup

import time

#######################################################################################

root_url = "https://michiganross.umich.edu/faculty-research/directory"

# we need to send request and get data for professors Alphabetically
ALPHABETS = [
    'A',
    'B',
    'C',
    'D',
    'E',
    'F',
    'G',
    'H',
    'I',
    'J',
    'K',
    'L',
    'M',
    'N',
    'O',
    'P',
    'Q',
    'R',
    'S',
    'T',
    'U',
    'V',
    'W',
    'X',
    'Y',
    'Z'
]


# let us create an empty data frame which will contain the required information
df = pd.DataFrame()

# these are the fields that we need to collect data for
PROFESSOR_NAME = []
ACADEMIC_AREA = []
POSITION = []
EMAIL = []

# list of tuples which show any exceptions that occured during the scraping process
EXCEPTION_LIST = []

for letter in ALPHABETS:
    url = root_url + '?view=list' + f'&last={letter}' + '&status=Active'

    try:
        # sleeping the requests for some time to avoid getting timed out
        time.sleep(1)


        # get the request from the web site alphabetically
        r = requests.get(url)

        # creating a soup object with the obtained response
        soup = BeautifulSoup(r.content, 'html.parser')

        required_table = soup.find_all("div", class_="view-content")[0].contents[1]

        # every <tr> in the table represents one single professor
        proff_list = required_table.find_all('tr')

        for one_proff in proff_list:
            PROFESSOR_NAME.append(one_proff.contents[3].contents[1].contents[0].text)
            POSITION.append(one_proff.contents[3].contents[2].text[:-1])
            ACADEMIC_AREA.append(one_proff.contents[5].text[13:-10])
            EMAIL.append(one_proff.contents[11].text[13:-10])

            print(f"...[LOG]... added {one_proff.contents[3].contents[1].contents[0].text} ... ")


    except:
        EXCEPTION_LIST.append((letter, url))
        pass

# adding the fields to the data frame
df['NAME'] = PROFESSOR_NAME
df['POSITION'] = POSITION
df['ACADEMIC_AREA'] = ACADEMIC_AREA
df['EMAIL'] = EMAIL

# we should find the professors who do not have the email available
# list to store the professors whose email is not present in the website
EMAIL_NOT_FOUND = []

# removing the professors who dont have an email 
for i in range(len(df)):
    if len(str(df.loc[i,['EMAIL']][0])) < 2:
        EMAIL_NOT_FOUND.append(str(df.loc[i,['NAME']][0]))
        df = df.drop(i)

# data frame to store the name of the professors whose email is un available on 
# the website
exception_df = pd.DataFrame()
exception_df['EXCEPTIONS'] = EMAIL_NOT_FOUND

# now we need to export the data frames to excel files
with pd.ExcelWriter('./output/MichiganRoss_ProfessorData.xlsx') as writer:
    # first the main data frame
    df.to_excel(writer, "professors", index=False)

with pd.ExcelWriter('./output/Exceptions.xlsx') as writer:
    exception_df.to_excel(writer, "email not found", index=False)

print("-------------------------------------[END]--------------------------------------")