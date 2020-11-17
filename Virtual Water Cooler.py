
# coding: utf-8

# In[ ]:


# Pip Installs
#!pip install pywin32


# In[113]:


"""Initialization"""
# Import statements
import pandas as pd
from IPython.display import display
import win32com.client as win32

"""Global Variables"""
# Create an empty list to store matches (pairs)
matches = []

# Create an empty list to store no matches
noMatches = []


# In[114]:


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


# In[125]:


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



# In[116]:


"""Create Groups of People Who Said Yes to Only Within Field"""
# Create a data frame of the people who said Yes
yPeople = dfCopy.loc[dfCopy['Do you want to be matched ONLY WITHIN your field?'] == "Yes"]

# Display yPeople
display(yPeople)

# Create a list of groups of yPeople by Field
yOWF = groupby(yPeople, 'Which field are you in?')


# In[129]:


"""Match People Who Said Yes to Only Within Field to Other People who said Yes"""
# For each group in yOWF
for g in range(len(yOWF)):
    
    # Display the groups
    display(yOWF[g])
    
    # Make all the possible pairs per group until there's only one left
    # while the length of the group is greater than 1
    while (len(yOWF[g]) > 1):
        
        # Create an empty list to store a pair of people (will be cleared each time)
        pair = []
        
        # Add the first person to the pair
        pair.append(yOWF[g].iloc[[0]])
        
        # Group by language
        lang = langGroup(yOWF[g].iat[0, 2], yOWF[g])
        
        # Group by time
        t = tGroup(yOWF[g].iat[0, 3], lang)
        
        # Remove the person from the t group
        t = t.drop(yOWF[g].iloc[[0]].isin(t).index)
        
        # Remove the person from the group
        yOWF[g] = yOWF[g].drop(yOWF[g].iloc[[0]].index)
        
        # Remove the person from the yPeople list
        yPeople = yPeople.drop([0])
        
        # Remove the person from the data
        dfCopy = dfCopy.drop(yOWF[g].iloc[[0]].isin(dfCopy).index)
        
        # Shuffle the rows (Random match)
        t = t.sample(frac = 1)
        
        # Add the first person in the shuffled group to the pair
        pair.append(t.iloc[[0]])

        # Remove the person from the group
        yOWF[g] = yOWF[g].drop(t.iloc[[0]].index)

        # Remove the person from the yPeople list
        yPeople = yPeople.drop(t.iloc[[0]].isin(yPeople).index)

        # Remove the person from the data
        dfCopy = dfCopy.drop(t.iloc[[0]].isin(dfCopy).index)
        
        # Add the pair to the matches list
        matches.append(pair)

    # There is only 1 person (left) in these groups within people who said yes
    # Do nothing
    pass


# In[130]:


"""Match People Who Said Yes & Still Haven't Been Matched Yet"""
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
    
    # For each remaining person in yPeople
    for i in range(len(yPeople)):

        # Print the groups
        display(yGroups[i])
        
        # if the length of the group is greater than 1
        if (len(yGroups[i]) > 1):

            # Create an empty list to store a pair of people (will be cleared each time)
            pair = []

            # Create a variable to store the person who said yes
            yPerson = yGroups[i].loc[yGroups[i]['Do you want to be matched ONLY WITHIN your field?'] == "Yes"]
            
            # Add the person who said yes into the pair (change this to be row concatenate with row)
            pair.append(yPerson)

            # Group by language
            lang = langGroup(yPerson.iat[0, 2], yGroups[i])

            # Group by time
            t = tGroup(yPerson.iat[0, 3], lang)

            # Remove the person from the t group
            t = t.drop(yPerson.isin(t).index)

            # Remove the person from the group
            yGroups[i] = yGroups[i].drop(yPerson.index)

            # Remove the person from the yPeople list
            yPeople = yPeople.drop(yPerson.isin(yPeople).index)

            # Remove the person from the data
            dfCopy= dfCopy = dfCopy.drop(yPerson.isin(dfCopy).index)

            # Shuffle the rows (Random match)
            t = t.sample(frac = 1)
            
            # Create a variable to store the person in the shuffled group
            first = t.iloc[[0]]
            
            # Add the first person in the shuffled group to the pair
            pair.append(first)

            # Remove the person from the group
            yGroups[i] = yGroups[i].drop(first.index)

            # Remove the person from the data
            dfCopy = dfCopy.drop(first.isin(dfCopy).index)

            # Add the pair to the matches list
            matches.append(pair)

        # There is only 1 person (left) in these groups within people who said yes
        # Do nothing
        pass

# Print updated dfCopy
display(dfCopy)


# In[131]:


for i in range(len(matches)):
    for j in range(len(matches[i])):
        display(matches[i][j])

