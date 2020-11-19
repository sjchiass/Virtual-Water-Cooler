
# coding: utf-8

# In[1]:


# Pip Installs
#!pip install pywin32


# In[2]:


"""Initialization"""
# Import statements
import pandas as pd
from IPython.display import display
import win32com.client as win32


# In[3]:


"""Global Variables"""
# Create an empty list to store matches (pairs)
matches = []

# Create an empty list to store no matches
noMatches = []


# In[4]:


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
        
        # Select all the rows that contain English or No preference
        Lang = df.loc[((df['What language would you like to converse in?'] == "English") |
                      (df['What language would you like to converse in?'] == "No preference"))]
        
        # Return the data frame of English and No preference speakers
        return Lang
    
    # Else if language is French
    elif (l == "French"):
        
        # Select all the rows that contain French or No preference
        Lang = df.loc[((df['What language would you like to converse in?'] == "French") |
                      (df['What language would you like to converse in?'] == "No preference"))]
        
        # Return the data frame of French and No preference speakers
        return Lang
    
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
        
        # Select all the rows that contain Morning or No preference
        T = df.loc[((df['When would you like to chat?'] == "Morning") | 
                   (df['When would you like to chat?'] == "No preference"))]
        
        # Return the data frame of people who prefer to chat in the Morning and No preference 
        return T
    
    # Else if time is Afternoon
    elif (t == "Afternoon"):
        
        # Select all the rows that contain Afternoon or No preference
        T = df.loc[((df['When would you like to chat?'] == "Afternoon") |
                  (df['When would you like to chat?'] == "No preference"))]
        
        # Return the data frame of people who prefer to chat in the Morning and No preference
        return T
    
    # Else the time is No preference
    else:
        # Return the original data frame
        return df



# In[5]:


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


# In[6]:


"""Create Groups of People Who Said Yes to Only Within Field"""
# Create a data frame of the people who said Yes
yPeople = dfCopy.loc[dfCopy['Do you want to be matched ONLY WITHIN your field?'] == "Yes"]

# Display yPeople
display(yPeople)

# Create a list of groups of yPeople by Field
yOWF = groupby(yPeople, 'Which field are you in?')


# In[7]:


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


# In[8]:


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


# In[9]:


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


# In[10]:


"""Create Group of People Who Speak French"""
# Create a data frame of the people who speak French
fr = dfCopy.loc[dfCopy['What language would you like to converse in?'] == "French"]

# Display fr
display(fr)


# In[11]:


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


# In[12]:


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


# In[13]:


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


# In[14]:


# Check outputs
for pair in matches:
    display(pair)

print("\nThese are the persons who were not matched")
for person in noMatches:
    display(person)

