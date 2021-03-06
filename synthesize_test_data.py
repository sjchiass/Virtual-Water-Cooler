#
# A convenient script for generating synthetic data
#
# This will save a random_test_data.csv file in the same folder
#
# Requirements: pandas, Faker (available on pip)
# 
import pandas as pd
import random

from faker.factory import Factory
Faker = Factory.create
fake = Faker()
fake.seed(0)

# The number of rows to generate
number_of_obs = 100

# The minimum amount of interests to generate
# Must be more x >= 0
interests_min = 0

assert interests_min >= 0

# The maximum number of interests to generate
# Must be greater or equal to the minimum amount
interests_max = 0

assert interests_max >= interests_min

# There are the column names generated by the form and expected by the script.
column_names = [
        "Please enter your @canada.ca email.",
        "What is your preferred name?",
        "What language would you like to converse in?",
        "When would you like to chat?",
        "Which field are you in?",
        "Do you want to be matched ONLY WITHIN your field?",
        "What are your interests? Please list your interests separated by a comma with no spaces. I.e. cooking,drawing,etc."
        ]

# The use of fake.safe_email() here should minimize the damage if someone were
# to run the automated e-mail script on these fake addresses.
emails = [fake.safe_email() for x in range(number_of_obs)]

# Generate fake names
names = [fake.name() for x in range(number_of_obs)]

# Randomly choose a language preference
languages = [random.choice(["English", "French", "No preference"]) for x in range(number_of_obs)]

# Randomly choose a time of the day to meet
times = [random.choice(["Morning", "Afternoon", "No preference"]) for x in range(number_of_obs)]

# Randomly choose a field number
fields = [random.choice([1, 3, 5, 6, 7, 8, 9]) for x in range(number_of_obs)]

# Randomly choose a field preference
field_preferences = [random.choice(["Yes", "No"]) for x in range(number_of_obs)]

# Generate some fake interests
interests = [",".join(fake.words(nb=random.randint(interests_min, interests_max))) for x in range(number_of_obs)]

# Combine everything into a dataframe
df = pd.DataFrame([emails, names, languages, times, fields, field_preferences, interests]).T

# Set the column names
df.columns = column_names

# Save the dataframe to disk
df.to_csv("./random_test_data.csv", index=False)