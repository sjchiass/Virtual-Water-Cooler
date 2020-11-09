
# coding: utf-8

# In[48]:


# Pip Installs
get_ipython().system('pip install pywin32')


# In[1]:


# Task: Create an automated coworker matchmaking script
# Date started: October 23, 2020
# Created by: Dennis Huynh
# Log:
# 10/23/2020 - Completed initialization
# 10/26/2020 - Completed grouping
# 10/27/2020 - Completed automated email process. Ran into an issue pip installing googletrans (something about proxy networks)
# 10/28/2020 - Created separate program to translate (On the Cloud)

"""Initialization"""
# Import statements
import pandas as pd
from IPython.display import display
import win32com.client as win32
# look into accessing google sheet directly with python (https://www.twilio.com/blog/2017/02/an-easy-way-to-read-and-write-to-a-google-spreadsheet-in-python.html)

# Read the translated, combined responses csv file skipping the column header row.
df = pd.read_csv("testData.csv")

# Remove the Interest column (In Beta Mode)
df = df.iloc[:, :6]

"""
# For each row in the dataframe in the Interests column, make the list a frozenset
for i in range(len(df)):
    df.iat[i, 6] = frozenset(df.iat[i, 6].split(','))
"""

# Print df
display(df)


# In[3]:


"""Matchmaking"""

# Copies of df so df values don't change. Groups will update if df updates
# Only care about English, Morning Pairs
dfCopy1 = df.copy()

# Only care about English, Afternoon Pairs
dfCopy2 = df.copy()

# Only care about French, Morning Pairs
dfCopy3 = df.copy()

# Only care about French, Afternoon Pairs
dfCopy4 = df.copy()

# For each row in df (Note: since they're all copies, they all have the same length)
for r in range(len(df)):
    """Update dfCopy1"""
    # Process Language column such that English and No Preference match
    # If Language column value is English or No Preference
    if ((dfCopy1.iat[r, 2] == "English") or (dfCopy1.iat[r, 2] == "No preference")):
        # Assign English
        dfCopy1.iat[r, 2] = "English"
    
    # Process Time Preference column such that Morning and No Preference match
    # If Time Preference column value is Morning) or No Preference
    if ((dfCopy1.iat[r, 3] == "Morning") or (dfCopy1.iat[r, 3] == "No preference")):
        # Assign Morning
        dfCopy1.iat[r, 3] = "Morning"

    """Update of dfCopy2"""
    # Process Language column such that English and No Preference match
    # If Language column value is English or No Preference
    if ((dfCopy2.iat[r, 2] == "English") or (dfCopy2.iat[r, 2] == "No preference")):
        # Assign English
        dfCopy2.iat[r, 2] = "English"

    # Process Time Preference column such that Afternoon and No Preference match
    # If Time Preference column value is Afternoon or No Preference
    if ((dfCopy2.iat[r, 3] == "Afternoon") or (dfCopy2.iat[r, 3] == "No preference")):
        # Assign Afternoon
        dfCopy2.iat[r, 3] = "Afternoon"
    
    """Update of dfCopy3"""
    # Process Language column such that French and No Preference match
    # If Language column value is French or No Preference
    if ((dfCopy3.iat[r, 2] == "French") or (dfCopy3.iat[r, 2] == "No preference")):
        # Assign French
        dfCopy3.iat[r, 2] = "French"

    # Process Time Preference column such that Morning and No Preference match
    # If Time Preference column value is Morning) or No Preference
    if ((dfCopy3.iat[r, 3] == "Morning") or (dfCopy3.iat[r, 3] == "No preference")):
        # Assign Morning
        dfCopy3.iat[r, 3] = "Morning"
    
    """Update of dfCopy4"""
    # Process Language column such that French and No Preference match
    # If Language column value is 1 (French) or No Preference (2)
    if ((dfCopy4.iat[r, 2] == "French") or (dfCopy4.iat[r, 2] == "No Preference")):
        # Assign French
        dfCopy4.iat[r, 2] = "French"

    # Process Time Preference column such that Afternoon and No Preference match
    # If Time Preference column value is Afternoon or No Preference
    if ((dfCopy4.iat[r, 3] == "Afternoon") or (dfCopy4.iat[r, 3] == "No Preference")):
        # Assign Afternoon
        dfCopy4.iat[r, 3] = "Afternoon"
    
# See updated copies of df
display(dfCopy1)
display(dfCopy2)
display(dfCopy3)
display(dfCopy4)


# In[4]:


# Merge all the edited dataframes
merge1 = pd.merge(dfCopy1, dfCopy2, how="outer")
merge2 = pd.merge(merge1, dfCopy3, how="outer")
mergeData = pd.merge(merge2, dfCopy4, how="outer")

# Print complete merged dataframe
display(mergeData)

# Group by Only Within Field, Field, Language, and Time Preference
o = mergeData.groupby(['Only Within Field', 'Field', 'Language', 'Time Preference'])

# Get the groups keys
oKeys = list(o.groups.keys())

# Create empty list (type: Dataframe) to store match grouped
matchGroups = list()

# Create empty list (type: Dataframe) to store people with no matches
nomatchGroups = list()

# Print the groups in o
for k in range(len(oKeys)):
    # If length of grouped dataframe is more than 1 (at least a pair)
    if (len(o.get_group(oKeys[k])) > 1):
        # Add the groups into the match list
        matchGroups.append(o.get_group(oKeys[k]))
    # Else length of dataframe is 1
    else:
        # Add the groups into the nomatch list
        nomatchGroups.append(o.get_group(oKeys[k]))

# Print statements to check contents
print("Matches")
for i in range(len(matchGroups)):
    display(matchGroups[i])

print("No matches")
for j in range(len(nomatchGroups)):
    display(nomatchGroups[j])


# In[10]:


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
    mgEmail.append(matchGroups[i]['Email'].tolist())
    mgName.append(matchGroups[i]['Preferred Name'].tolist())
    mgField.append(matchGroups[i]['Field'].tolist())
    
    # Add the word "and" to the last element in mgName
    mgName[i][-1] = "and " + mgName[i][-1]
    
    # Add the word "and" to the last element in mgField
    mgField[i][-1] = "and " + mgField[i][-1]
    
    # Body text of email (matches)
    text = """Hello {}


You have been matched together for a Virtual Watercooler conversation. We recommend using MS 
Teams scheduled during regular business hours for a conversation of at about 10 minutes but it is up to 
you to decide how to proceed.

The group's preferences are {} and {}. Your field of work is {}, respectively.

As this is our beta version so please reach out to Innovation Coordinator / Coord de l'innovation 
(STATCAN) <username>@<email domain> with all of your 
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
    # Replace the word "and" with "et" in the last element of mgName
    mgName[i][-1] = mgName[i][-1].replace("and", "et")
    
    # Replace the word "and" with "et" in the last element of mgName
    mgField[i][-1] = mgField[i][-1].replace("and", "et")
    
    textFr = """Bonjour {}, 

Hardcode the French translation here then copy formatting"""
    message = text + "\n\n" + textFr
     # Add the group message to the matched group message list
    mgMsgList.append(message)
    print(mgMsgList[i])

# A list of recipients for each match group
mgRecList = [', '.join(rec) for rec in mgEmail]

"""EMAIL CODE BELOW"""
# Note that all of the lists have the same length as matchGroups (since these lists are built from matchGroups)
# For each match group, send an email
#for e in range(len(matchGroups)):
#    email(mgRecList[e], matchSubject, mgMsgList[e])


# In[103]:


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
#<username>@<email domain>""".format(nomatchGroups[i].iat[0,1])
         
    # Add the group message to the matched group message list
#    nmgMsgList.append(body)

# For each person who was not matched, send an email
#for j in range(len(nomatchGroups)):
    # email column from dataframe will always be only row, 1st column (index at 0)
#    email(nomatchGroups[j].iat[0,0], nomatchSubject, nmgMsgList[j])

