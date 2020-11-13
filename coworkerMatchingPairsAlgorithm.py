
# coding: utf-8

# In[ ]:


# Pip Installs
#!pip install pywin32


# In[15]:


# Task: Create an automated coworker matchmaking script
# Date started: October 23, 2020
# Created by: Dennis Huynh
# Log:
# 10/23/2020 - Completed initialization
# 10/26/2020 - Completed grouping
# 10/27/2020 - Completed automated email process. Ran into an issue pip installing googletrans (something about proxy networks)
# 10/28/2020 - Created separate program to translate (On the Cloud)
# 11/09/2020 - Matching algorithm only makes unique pairs
# 11/10/2020 - Completed French translation of email message

"""Initialization"""
# Import statements
import pandas as pd
import numpy as np
from IPython.display import display
import win32com.client as win32

# Create a map from original value to int dictionary
wordDict = {
    "English": -1,
    "French": 1,
    "Morning": -1,
    "Afternoon": 1,
    "No preference": 0,
    "Field 1 - Office of the Chief Statistician": int(1),
    "Field 3 - Corporate Strategy and Management": int(2),
    "Field 4 - Strategic Engagement": int(3),
    "Field 5 - Economics Statistics": int(4),
    "Field 6 - Strategic Data Management, Methods, and Analysis": int(5),
    "Field 7 - Census, Regional Services, and Operations": int(6),
    "Field 8 - Social Health and Labour Statistics": int(7),
    "Field 9 - Digital Solutions": int(8),
    "No": 0,
    "Yes": 1
}

# Create a English-French dictionary
engDict = {
    "English": "Anglais",
    "French": "Français",
    "Morning": "la Matinée",
    "Afternoon": "l'Après-midi",
    "Field 1 - Office of the Chief Statistician": "Secteur 1 - Bureau du Statisticien en Chef",
    "Field 3 - Corporate Strategy and Management": "Secteur 3 - Stratégies et Gestion Intégrées",
    "Field 4 - Strategic Engagement": "Secteur 4 - Engagement Stratégique",
    "Field 5 - Economics Statistics": "Secteur 5 - Statistiques Économique",
    "Field 6 - Strategic Data Management, Methods, and Analysis": "Secteur 6 - Gestion Stratégique des Données, Méthodes et Analyse",
    "Field 7 - Census, Regional Services, and Operations": "Secteur 7 - Recensement, Services Régionaux, et Opérations",
    "Field 8 - Social Health and Labour Statistics": "Secteur 8 - Statistiques Sociale, de la Santé et du Travail",
    "Field 9 - Digital Solutions": "Secteur 9 - Solutions Numériques"
}

# Read the translated, combined responses csv file skipping the column header row.
df = pd.read_csv("dataBeta.csv")

# Remove the Interest column (In Beta Mode)
df = df.iloc[:, :6]

"""
# For each row in the dataframe in the Interests column, make the list a frozenset
for i in range(len(df)):
    df.iat[i, 6] = frozenset(df.iat[i, 6].split(','))
"""

# Make a copy of the data frame
dfCopy = df.copy()

# Change the values in these columns to integers
dfCopy.iloc[:,2:6] = dfCopy.iloc[:,2:6].applymap(wordDict.get)

# Print df
display(df)

# Print dfCopy
display(dfCopy)


# In[19]:


"""Matchmaking"""
# From: https://stackoverflow.com/questions/53996421/matching-two-people-together-based-on-attributes

# An empty list for all possible matches
matches = []

# Convert data frame to numpy array
dfarr = np.array(dfCopy)

# Iterating the array row
for i in range(len(dfarr) - 1):
    
    # Iterating the array row + 1
    for j in range(i + 1, len(dfarr)):
        
        # Check for Language condition to include relevant records
        if dfarr[i][2] * dfarr[j][2] >= 0:
            
            # Check for Time condition to include relevant records
            if dfarr[i][3] * dfarr[j][3] >= 0:
                
                # Check for Only Within Field condition to include relevant records
                if dfarr[i][5] * dfarr[j][5] >= 0:
                
                    # Empty list to store pairs
                    row = []

                    # Appending the names
                    row.append(dfarr[i][1])
                    row.append(dfarr[j][1])

                    # Appending the final score
                    row.append((dfarr[i][2] * dfarr[j][2]) +
                               (dfarr[i][3] * dfarr[j][3]) +
                               (dfarr[i][5] * dfarr[j][5]) + 
                               (round((1 - (abs(dfarr[i][4] -
                                                dfarr[j][4]) / 10)), 2)))

                    # Appending the row to the Matches array
                    matches.append(row)

# Convert array to data frame
ndf = pd.DataFrame(matches)

# Sort the data frame on Final Score
ndf = ndf.sort_values(by=[2], ascending=False)

# Print the data frame
display(ndf)


# In[ ]:


# Create empty list (type: Dataframe) to store match grouped
matchGroups = []

# Create empty list (type: Dataframe) to store people with no matches
nomatchGroups = []

# Fill the matchGroups and nomatchGroups lists
for k in range(len(oGroups)):
    
    # If length of grouped dataframe is equal to 2 (it's a pair)
    if (len(oGroups[k]) == 2):
        
        # Add the groups into the match list
        matchGroups.append(oGroups[k])
        
    # Else if the length of the grouped dataframe is greater than 2
    elif (len(oGroups[k]) > 2):
        
        # Shuffle the rows (random match)
        oGroups[k] = oGroups[k].sample(frac = 1)
        
        # If the number of people in the group is even
        if (len(oGroups[k]) % 2 == 0):
            
            # Divide the group into pairs
            for p in range(0, len(oGroups[k]), 2):
                matchGroups.append(oGroups[k].iloc[p:p+2])
        
        # Else the number of people in the group is odd
        else:
            # Add the last person on the list to nomatchGroups
            nomatchGroups.append((oGroups[k].iloc[-1]).to_frame().transpose())
            
            # Divide the reaminder of the group into pairs
            for p in range(0, len(oGroups[k]) - 1, 2):
                matchGroups.append(oGroups[k].iloc[p:p+2])
            
    # Else length of dataframe is 1
    else:
        # Add the groups into the nomatch list
        nomatchGroups.append(oGroups[k])

# Print statements to check contents
print("Matches")
for i in range(len(matchGroups)):
    display(matchGroups[i])

print("No matches")
for j in range(len(nomatchGroups)):
    display(nomatchGroups[j])


# In[ ]:


"""Automated Emails for Matched Groups"""

"""
Desc:
    Function that sends an email. 
    Note: you must be logged onto your Outlook 2013 account first before this will run
"""
def email(recipients, subject, text, profilename="Outlook 2013"):
    oa = win32.Dispatch("Outlook.Application")

    Msg = oa.CreateItem(0)
    Msg.To = recipients

    Msg.Subject = subject
    Msg.Body = text

    Msg.Display()
    # Msg.Send()


# Subject message for matches
matchSubject = "Virtual Watercooler Matches"

# Make empty list of lists by input to format the message of the email (mg = matched groups)
# stores recipient lists
mgEmail = list()

# stores recipients' names
mgName = list()

# stores recipients' field
mgField = list()

# A list of messages for each match group
mgMsgList = list()

# For each match group, write the body of the email
for i in range(len(matchGroups)):
    # Fill lists with list of their column values
    mgEmail.append(matchGroups[i]['Please enter your @canada.ca email.'].tolist())
    mgName.append(matchGroups[i]['What is your preferred name?'].tolist())
    mgField.append(matchGroups[i]['Which field are you in?'].tolist())
    
    # Add the word "and" to the last element in mgName
    mgName[i][-1] = "and " + mgName[i][-1]

    # Add the word "and" to the last element in mgField
    mgField[i][-1] = "and " + mgField[i][-1]

    # Body text of email (matches)
    text = """Hello {}


You have been matched together for a Virtual Watercooler conversation. We recommend using MS 
Teams scheduled during regular business hours for a conversation of at about 10 minutes but it is up to 
you to decide how to proceed.

The group prefers to chat in {} in the {}. You work in {}, respectively.

As this is our beta version so please reach out to Innovation Coordinator / Coord de l'innovation 
(STATCAN) statcan.innovationcoordinator-coorddelinnovation.statcan@canada.ca with all of your 
feedback, questions and suggestions. Thank you for using the StatCan Virtual Watercooler.



Sincerely,

The StatCan Virtual Watercooler Team

Innovation Secretariat""".format(', '.join(mgName[i]) if len(mgName[i]) > 2 else ' '.join(mgName[i]), 
                                 matchGroups[i].iat[0, 2], 
                                 matchGroups[i].iat[0, 3], 
                                 ', '.join(mgField[i]) if len(mgField[i]) > 2 else ' '.join(mgField[i]))
    # *NOTE: this comment explains the above code. Format message input by converting the list of 
    # names per match group into string, take the first row's language (since they're all the same), 
    # take the first row's time preference (since they're all the same), convert list of fields per 
    # match group into string

    # French version
    # Remove the word "and" in the last element of mgName 
    mgName[i][-1] = mgName[i][-1].replace("and ", "")

    # Replace the first instance of the word "and" in the last element of mgField
    mgField[i][-1] = mgField[i][-1].replace("and ", "", 1)
    
    # Map the values for Language, Time Preference, and Field to their dictionary values (Translate English to French)
    matchGroups[i].iat[0, 2] = engDict.get(matchGroups[i].iat[0, 2])
    matchGroups[i].iat[0, 3] = engDict.get(matchGroups[i].iat[0, 3])
    
    # Translate the Field values into French
    for f in range(len(mgField[i])):
        mgField[i][f] = engDict.get(mgField[i][f])
    
   # Add the word "et" to the last element in mgName
    mgName[i][-1] = "et " + mgName[i][-1]

    # Add the word "et" to the last element in mgField
    mgField[i][-1] = "et " + mgField[i][-1]
    
    # French translation of the email
    textFr = """Bonjour {},


Vous avez été jumelés pour une causerie virtuelle. Nous vous recommandons d’utiliser MS Teams 
pendant les heures normales de travail pour discuter environ 10 minutes, mais c’est à vous de décider 
de la manière de procéder.

Le groupe préfère discuter en {} dans {}. Vous travaillez dans le {}, respectivement.

Comme il s’agit d’une version bêta, nous vous invitons à communiquer avec le coordonnateur de 
l’innovation de Statistique Canada (statcan.innovationcoordinator-coorddelinnovation.statcan@canada.ca) 
si vous avez des commentaires, des questions et des suggestions. Nous vous remercions de participer aux 
causeries virtuelles de Statistique Canada.


Bien cordialement,

L’Équipe des causeries virtuelles de Statistique Canada 

Secrétariat de l’innovation""".format(', '.join(mgName[i]) if len(mgName[i]) > 2 else ' '.join(mgName[i]), 
                                 matchGroups[i].iat[0, 2], 
                                 matchGroups[i].iat[0, 3], 
                                 ', '.join(mgField[i]) if len(mgField[i]) > 2 else ' '.join(mgField[i]))
    # Body of text is now bilingual
    message = text + "\n\n\n" + textFr
    
     # Add the group message to the matched group message list
    mgMsgList.append(message)
    
    # Check email contents
    print(mgMsgList[i])

# A list of recipients for each match group
mgRecList = [', '.join(rec) for rec in mgEmail]

"""EMAIL CODE BELOW"""
# Note that all of the lists have the same length as matchGroups (since these lists are built from matchGroups)
# For each match group, send an email
#for e in range(len(matchGroups)):
#    email(mgRecList[e], matchSubject, mgMsgList[e])


# In[ ]:


# NOT FOR BETA
"""Automated Emails for No Matches Groups"""

# Subject message for no matches
#nomatchSubject = "Unfortunately, there are no current Matches from the Meeting Colleagues Survey"

# A list of messages for each unmatched person
#nmgMsgList = list()

# For each person in the nomatch list, write the body of the email
#for i in range(len(nomatchGroups)):
    # Body of the email for those with no matches. The format will always output the name from each dataframe
#    body = """Hello {}

#You will be notified by email when there has been a match. Would you still like to meet new people?
#If yes, please contact the Innovation Coordinator at 
#statcan.innovationcoordinator-coorddelinnovation.statcan@canada.ca""".format(nomatchGroups[i].iat[0,1])
         
    # Add the group message to the matched group message list
#    nmgMsgList.append(body)

# For each person who was not matched, send an email
#for j in range(len(nomatchGroups)):
    # email column from dataframe will always be only row, 1st column (index at 0)
#    email(nomatchGroups[j].iat[0,0], nomatchSubject, nmgMsgList[j])

