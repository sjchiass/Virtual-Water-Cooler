
# coding: utf-8

# In[17]:


# Pip Installs
#!pip install pywin32


# In[18]:


"""Initialization"""
# Import statements
import pandas as pd
from IPython.display import display
import win32com.client as win32


# In[19]:


"""Global Variables"""
# Create an empty list to store matches (pairs)
matches = []

# Create an empty list to store no matches
noMatches = []


# In[20]:


"""Helper Functions"""

"""
Desc:
    Function that creates groups based on a column's vlaues
    Params: colname (column name; string) and df (input data frame; pandas DataFrame)
    Output: a list of data frames (pandas DataFrame)
"""
def groupby(df, colname):
    # Group df by colname
    g = df.groupby([colname])

    # Get the groups keys
    keys = list(g.groups.keys())

    # Create an empty list to store the groups
    groups = []

    # Get the groups
    for k in range(len(keys)):
        groups.append(g.get_group(keys[k]))

    # Return the groups
    return groups


"""
Desc:
    Function that selects rows from a data frame based on the language preference column.
    Params: l (the language; string) and df (input data frame; pandas DataFrame)
    Output: a data frame (pandas DataFrame)
"""
def langGroup(l, df):
    # If language is English
    if (l == "English"):
        
        # Return all the rows that contain English or No preference
        return df.loc[((df['What language would you like to converse in?'] == "English") |
                      (df['What language would you like to converse in?'] == "No preference"))]
    
    # Else if language is French
    elif (l == "French"):
        
        # Return all the rows that contain French or No preference
        return df.loc[((df['What language would you like to converse in?'] == "French") |
                      (df['What language would you like to converse in?'] == "No preference"))]
    
    # Else the language is No preference
    else:
        # Return the original data frame
        return df


"""
Desc:
    Function that selects rows from a data frame based on the time preference column.
    Params: t (the time; string) and df (input data frame; pandas DataFrame)
    Output: a data frame (pandas DataFrame)
"""
def tGroup(t, df):
    # If time is Morning
    if (t == "Morning"):
        
        # Return all the rows that contain Morning or No preference
        return df.loc[((df['When would you like to chat?'] == "Morning") | 
                   (df['When would you like to chat?'] == "No preference"))]
    
    # Else if time is Afternoon
    elif (t == "Afternoon"):
        
        # Return all the rows that contain Afternoon or No preference
        return df.loc[((df['When would you like to chat?'] == "Afternoon") |
                  (df['When would you like to chat?'] == "No preference"))]
    
    # Else the time is No preference
    else:
        # Return the original data frame
        return df


"""
Desc:
    Function that picks the language.
    Params: pair (pandas DataFrame)
    Output: the language (string)
"""
def language(pair):
    # If the first person or the second person said English
    if ((pair.iat[0, 2] == "English") | (pair.iat[1, 2] == "English")):
        
        # Return English
        return "English"
    
    # Else if the first person or second person said French
    elif ((pair.iat[0, 2] == "French") | (pair.iat[1, 2] == "French")):
        
        # Return French
        return "French"
    
    # Else it was No preference
    else:
        return "either English or French"


"""
Desc:
    Function that picks the time.
    Params: pair (pandas DataFrame)
    Output: the time (string)
"""
def time(pair):
    # If the first person or the second person said Morning
    if ((pair.iat[0, 3] == "Morning") | (pair.iat[1, 3] == "Morning")):
        
        # Return Morning
        return "Morning"
    
    # Else if the first person or second person said Afternoon
    elif ((pair.iat[0, 3] == "Afternoon") | (pair.iat[1, 3] == "Afternoon")):
        
        # Return Afternoon
        return "Afternoon"
    
    # Else it was No preference
    else:
        return "Morning or Afternoon"


"""
Desc:
    Function that picks the language (French).
    Params: pair (pandas DataFrame)
    Output: the language (string)
"""
def langue(pair):
    # If the first person or the second person said Anglais
    if ((pair.iat[0, 2] == "Anglais") | (pair.iat[1, 2] == "Anglais")):
        
        # Return Anglais
        return "Anglais"
    
    # Else if the first person or second person said Français
    elif ((pair.iat[0, 2] == "Français") | (pair.iat[1, 2] == "Français")):
        
        # Return Français
        return "Français"
    
    # Else it was Pas de préférence
    else:
        return "soit Anglais ou Français"


"""
Desc:
    Function that picks the time (French).
    Params: pair (pandas DataFrame)
    Output: the language (string)
"""
def temps(pair):
    # If the first person or the second person said Matin
    if ((pair.iat[0, 3] == "Matin") | (pair.iat[1, 3] == "Matin")):
        
        # Return la Matinée
        return "la Matinée"
    
    # Else if the first person or second person said Après-midi
    elif ((pair.iat[0, 3] == "Après-midi") | (pair.iat[1, 3] == "Après-midi")):
        
        # Return l'Après-midi
        return "l'Après-midi"
    
    # Else it was Pas de préférence
    else:
        return "la Matinée ou l'Après-midi"
        
        
"""
Desc:
    Function that sends an email.
    Params: recipients (string), subject (string), text (string)
    Output: An email
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



# In[21]:


# Create a English-French dictionary
engDict = {
    "English": "Anglais",
    "French": "Français",
    "Morning": "Matin",
    "Afternoon": "Après-midi",
    "No preference": "Pas de préférence",
    "Field 1 - Office of the Chief Statistician": "Secteur 1 - Bureau du Statisticien en Chef",
    "Field 3 - Corporate Strategy and Management": "Secteur 3 - Stratégies et Gestion Intégrées",
    "Field 4 - Strategic Engagement": "Secteur 4 - Engagement Stratégique",
    "Field 5 - Economics Statistics": "Secteur 5 - Statistiques Économique",
    "Field 6 - Strategic Data Management, Methods, and Analysis": "Secteur 6 - Gestion Stratégique des Données, Méthodes et Analyse",
    "Field 7 - Census, Regional Services, and Operations": "Secteur 7 - Recensement, Services Régionaux, et Opérations",
    "Field 8 - Social Health and Labour Statistics": "Secteur 8 - Statistiques Sociale, de la Santé et du Travail",
    "Field 9 - Digital Solutions": "Secteur 9 - Solutions Numériques",
}

"""Load the Data"""
# Read the translated, combined responses csv file skipping the column header row.
df = pd.read_csv("dataBeta.csv")

# Remove the Interest column (In Beta Mode)
df = df.iloc[:, :6]

"""
# For each row in the dataframe in the Interests column, make the list a frozenset
for i in range(len(df)):
    df.iat[i, 6] = frozenset(df.iat[i, 6].split(','))
"""

# Print df
display(df)

# Make a copy of df
dfCopy = df.copy()


# In[22]:


"""Create Groups of People Who Said Yes to Only Within Field"""
# Create a data frame of the people who said Yes
yPeople = dfCopy.loc[dfCopy['Do you want to be matched ONLY WITHIN your field?'] == "Yes"]

# Display yPeople
display(yPeople)

# Create a list of groups of yPeople by Field
yOWF = groupby(yPeople, 'Which field are you in?')


# In[23]:


"""Match People Who Said Yes to Only Within Field to Other People Who said Yes"""
# For each group in yOWF
for g in range(len(yOWF)):
    
    # Display the groups
    display(yOWF[g])
    
    # Add the first person to the pair
    pair = yOWF[g].iloc[[0]]
    
    # Make all the possible pairs per group until there's only one left
    # While the length of the group is greater than 1
    while (len(yOWF[g]) > 1):
        
        # Group by language
        lang = langGroup(yOWF[g].iat[0, 2], yOWF[g])
        
        # Group by time
        t = tGroup(yOWF[g].iat[0, 3], lang)
        
        # If after filtering the group is greater than one
        if (len(t) > 1):
            
            # Remove the person from the t group
            t = t.drop(pair.isin(t).index)

            # Remove the person from the group
            yOWF[g] = yOWF[g].drop(pair.index)

            # Remove the person from the yPeople list
            yPeople = yPeople.drop(pair.isin(yPeople).index)

            # Remove the person from the data
            dfCopy = dfCopy.drop(pair.isin(dfCopy).index)
            
            # Shuffle the rows (Random match)
            t = t.sample(frac = 1)

            # Add the first person in the shuffled group to the pair
            pair = pair.append(t.iloc[[0]], ignore_index = True)

            # Remove the person from the group
            yOWF[g] = yOWF[g].drop(t.iloc[[0]].index)

            # Remove the person from the yPeople list
            yPeople = yPeople.drop(t.iloc[[0]].isin(yPeople).index)

            # Remove the person from the data
            dfCopy = dfCopy.drop(t.iloc[[0]].isin(dfCopy).index)

            # Add the pair to the matches list
            matches.append(pair)
            
        # Else there is only 1 person (left) in these groups within people who said yes
        else:
            # Do nothing. Exit the while loop
            break
        
    # Clear the pair
    pair = None


# In[24]:


"""Select the Groups People Who Said Yes & Still Haven't Been Matched Yet"""
# If the number of people who said yes is greater than 0
if (len(yPeople) > 0):
    
    # Create a list of groups from the data by Field
    yOWF = groupby(dfCopy, 'Which field are you in?')
    
    # Create an empty list to store the groups that have someone who said Yes
    yGroups = []
    
    # Select the groups that have someone who said yes
    for g in range(len(yOWF)):
        if ("Yes" in yOWF[g]['Do you want to be matched ONLY WITHIN your field?'].unique()):
            yGroups.append(yOWF[g])


# In[25]:


"""Match People Who Said Yes & Still Haven't Been Matched Yet"""
# For each group
for i in range(len(yGroups)):

    # Print the groups
    display(yGroups[i])
    
    # Create a data frame per group of all the people who said Yes
    yPersons = yGroups[i].loc[yGroups[i]['Do you want to be matched ONLY WITHIN your field?'] == "Yes"]
    
    # Remove everyone who said yes in each group 
    yGroups[i] = yGroups[i].drop(yPersons.index)
    
    # Make all the possible pairs per group with people who said yes until there's only one left
    # While there is more than 0 persons who said yes
    while (len(yPersons) > 0):
    
        # Add the first person to the pair
        pair = yPersons.iloc[[0]]
            
        # Group by language
        lang = langGroup(pair.iat[0, 2], yGroups[i])

        # Group by time
        t = tGroup(pair.iat[0, 3], lang)

        # If after filtering the group is greater than 0
        if (len(t) > 0):

            # Remove the person from yPerson list
            yPersons = yPersons.drop(pair.isin(yPersons).index)

            # Remove the person from the yPeople list
            yPeople = yPeople.drop(pair.isin(yPeople).index)

            # Remove the person from the data
            dfCopy = dfCopy.drop(pair.isin(dfCopy).index)

            # If after filtering the group is greater than 1
            if (len(t) > 1):

                # Shuffle the rows (Random match)
                t = t.sample(frac = 1)

            # Add the first person in the shuffled group to the pair
            pair = pair.append(t.iloc[[0]], ignore_index = True)

            # Remove the person from the group
            yGroups[i] = yGroups[i].drop(t.iloc[[0]].index)

            # Remove the person from the data
            dfCopy = dfCopy.drop(t.iloc[[0]].isin(dfCopy).index)

            # Add the pair to the matches list
            matches.append(pair)
            
            """
            Unsure if this needs to be here
            # Else there is only 1 person (left) in these groups within people who said yes
            else:
                # Add this person to the list of no matches
                noMatches.append(pair)

                # Remove the person from the yPersons list
                yPersons = yPersons.drop(pair.isin(yPersons).index)

                # Remove the person from the yPeople list
                yPeople = yPeople.drop(pair.isin(yPeople).index)

                # Remove the person from the data
                dfCopy = dfCopy.drop(pair.isin(dfCopy).index)
            """

        # Else there is only 1 person (left) in these groups within people who said yes
        else:
            # Add this person to the list of no matches
            noMatches.append(pair)
            
            # Remove the person from the yPersons list
            yPersons = yPersons.drop(pair.isin(yPersons).index)

            # Remove the person from the yPeople list
            yPeople = yPeople.drop(pair.isin(yPeople).index)

            # Remove the person from the data
            dfCopy = dfCopy.drop(pair.isin(dfCopy).index)
        
        # Clear the pair
        pair = None

# Print updated dfCopy
display(dfCopy)


# In[26]:


"""Create Group of People Who Speak French"""
# Create a data frame of the people who speak French
fr = dfCopy.loc[dfCopy['What language would you like to converse in?'] == "French"]

# Display fr
display(fr)


# In[27]:


"""Match People Who Speak French with Other People Who Speak French"""
"""WARNING: Potentially an infinite loop"""

# For each person in fr
for p in range(len(fr)):
    
    # Add the first person to the pair
    pair = fr.iloc[[0]]
    
    # Make all the possible pairs per group until there's only one left
    # While the length of the group is greater than 1
    while (len(fr) > 1):
        
        # Group by time
        t = tGroup(fr.iat[0, 3], lang)
        
        # If after filtering the group is greater than one
        if (len(t) > 1):
            
            # Remove the person from the t group
            t = t.drop(pair.isin(t).index)

            # Remove the person from the group
            fr = fr.drop(pair.index)

            # Remove the person from the data
            dfCopy = dfCopy.drop(pair.isin(dfCopy).index)
            
            # Shuffle the rows (Random match)
            t = t.sample(frac = 1)

            # Add the first person in the shuffled group to the pair
            pair = pair.append(t.iloc[[0]], ignore_index = True)

            # Remove the person from the group
            fr = fr.drop(t.iloc[[0]].index)

            # Remove the person from the data
            dfCopy = dfCopy.drop(t.iloc[[0]].isin(dfCopy).index)

            # Add the pair to the matches list
            matches.append(pair)
            
        # Else there is only 1 person (left)
        else:
            # Do nothing. Exit the while loop
            break
        
    # Clear the pair
    pair = None

# Print the updated French group
display(fr)


# In[28]:


"""Match People Who Speak French & Still Haven't Been Matched Yet"""
# Remove everyone who speaks French in the data frame
dfCopy = dfCopy.drop(fr.index)

# Make all the possible pair per group with people who speak French until there's only one left
# While there is more than 0 persons who speak French
while (len(fr) > 0):

    # Add the first person to the pair
    pair = fr.iloc[[0]]

    # Group by time
    t = tGroup(pair.iat[0, 3], dfCopy)

    # If after filtering the group is greater than 0
    if (len(t) > 0):

        # Remove the person from fr list
        fr = fr.drop(pair.isin(fr).index)

        # If after filtering the group is greater than 1
        if (len(t) > 1):

            # Shuffle the rows (Random match)
            t = t.sample(frac = 1)

        # Add the first person in the shuffled group to the pair
        pair = pair.append(t.iloc[[0]], ignore_index = True)

        # Remove the person from the data
        dfCopy = dfCopy.drop(t.iloc[[0]].isin(dfCopy).index)

        # Add the pair to the matches list
        matches.append(pair)

        """
        Unsure if this needs to be here
        # Else there is only 1 person (left) who speaks French
        else:
            # Add this person to the list of no matches
            noMatches.append(pair)

            # Remove the person from the fr list
            fr = fr.drop(pair.isin(fr).index)
        """

    # Else there is only 1 person (left) who speaks French
    else:
        # Add this person to the list of no matches
        noMatches.append(pair)

        # Remove the person from the fr list
        fr = fr.drop(pair.isin(fr).index)

    # Clear the pair
    pair = None

# Print updated dfCopy
display(dfCopy)


# In[29]:


"""Match People Who Still Haven't Been Matched Yet"""
# Make all the possible pairs until there's only one left
# While there are still people
while (len(dfCopy) > 0):

    # Add the first person to the pair
    pair = dfCopy.iloc[[0]]

    # Remove the person from list
    dfCopy = dfCopy.drop(pair.index)
    
    # Group by time
    t = tGroup(pair.iat[0, 3], dfCopy)

    # If after filtering the group is greater than 0
    if (len(t) > 0):

        # If after filtering the group is greater than 1
        if (len(t) > 1):

            # Shuffle the rows (Random match)
            t = t.sample(frac = 1)

        # Add the first person in the shuffled group to the pair
        pair = pair.append(t.iloc[[0]], ignore_index = True)

        # Remove the person from the data
        dfCopy = dfCopy.drop(t.iloc[[0]].isin(dfCopy).index)

        # Add the pair to the matches list
        matches.append(pair)

        """
        Unsure if this needs to be here
        # Else there is only 1 person (left) who speaks French
        else:
            # Add this person to the list of no matches
            noMatches.append(pair)
        """

    # Else there is only 1 person left
    else:
        # Add this person to the list of no matches
        noMatches.append(pair)

    # Clear the pair
    pair = None

# Print updated dfCopy (Should be empty)
display(dfCopy)


# In[30]:


# Check outputs
for pair in matches:
    display(pair)

print("\nThese are the persons who were not matched")
for person in noMatches:
    display(person)


# In[31]:


"""Automated Emails for Matched Groups"""
# For each pair in list of matches
for pair in matches:

    # Body text of email (English)
    text = """Hello {} and {},


You have been matched together for a Virtual Watercooler conversation. We recommend using MS 
Teams scheduled during regular business hours for a conversation of at about 10 minutes but it is up to 
you to decide how to proceed.

The group prefers to chat in {} in the {}. You work in {} and {}, respectively.

As this is our beta version so please reach out to Innovation Coordinator / Coord de l'innovation 
(STATCAN) <email>@<email domain> with all of your 
feedback, questions and suggestions. Thank you for using the StatCan Virtual Watercooler.



Sincerely,

The StatCan Virtual Watercooler Team

Innovation Secretariat""".format(pair.iat[0, 1],  pair.iat[1, 1], # Names
                                 language(pair), # Language preference
                                 time(pair), # Time preference
                                 pair.iat[0, 4], pair.iat[1, 4]) # Field

    # French version
    # Map the values for Language, Time Preference, and Field to their dictionary values (Translate English to French)
    pair.iloc[:, 2:5] = pair.applymap(engDict.get)
    
    # French translation of the email
    textFr = """Bonjour {} and {},


Vous avez été jumelés pour une causerie virtuelle. Nous vous recommandons d’utiliser MS Teams 
pendant les heures normales de travail pour discuter environ 10 minutes, mais c’est à vous de décider 
de la manière de procéder.

Le groupe préfère discuter en {} dans {}. Vous travaillez dans {} et {}, respectivement.

Comme il s’agit d’une version bêta, nous vous invitons à communiquer avec le coordonnateur de 
l’innovation de Statistique Canada (<email>@<email domain>) 
si vous avez des commentaires, des questions et des suggestions. Nous vous remercions de participer aux 
causeries virtuelles de Statistique Canada.


Bien cordialement,

L’Équipe des causeries virtuelles de Statistique Canada 

Secrétariat de l’innovation""".format(pair.iat[0, 1],  pair.iat[1, 1], # Names
                                 langue(pair), # Language preference
                                 temps(pair), # Time preference
                                 pair.iat[0, 4], pair.iat[1, 4]) # Field
    
    # Final email message
    message = text + "\n\n\n" + textFr
    print(message)

    # The emails from each person in the pair
    recipients = pair.iat[0, 0] + "; " + pair.iat[1, 0]
    
    # Send the emails
    #email(recipients, "Virtual Water Cooler", message)


# In[32]:


# NOT FOR BETA
"""Automated Emails for No Matches"""

# For each person in the nomatch list, write the body of the email
#for person in noMatches:
    # Body of the email for those with no matches. The format will always output the name from each dataframe
#    text = """Hello {},

#You will be notified by email when there has been a match. Would you still like to meet new people?
#If yes, please contact the Innovation Coordinator at 
#<email>@<email domain>""".format(nomatchGroups[i].iat[0,1])
         

#    email(person.iat[0,0], "Virtual Water Cooler", message)

