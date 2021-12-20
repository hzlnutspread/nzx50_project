<h1>About the project</h1>

I was tasked with a project to create profiles for all 50 companies in the NZX50. This included creating my own weighted index based on the NZX50 Smartshares fund. In order to keep this up to date I need to update the share prices of each company every Friday. This is done in the <em>NZX50.xlsx</em> file. I was further given the task of creating a series of charts in drupal 8. The 5 charts for each company that would be uploaded to the website are the 5 .py files you see.

- Share price

- Rank

- Capitalisation 

- Capitalisation change (week's % change)

- Capitalisation change (week's $ change)

In addition to other manual tasks, this meant I had to update 250 csv files every friday, so with no prior coding experience I decided to try and automate this process using Python.

<h1>About the 5 csv files</h1>

**Share price:**
- Pulled values from csv link on Yahoo Finance.
- Sorted to pull only share prices that fall on Fridays
- Creates csv

**Rank:** 
- Based on my own master file called 'NZX 50.xlsl'
- Pulls the rank of each ticker
- Creates csv

**Capitalisation:**
- Based on my own master file called 'NZX 50.xlsl'
- Pulls capitalisation of each ticker
- Creates csv

**Capitalisation change (week's % change):**
- Based on the capitalisation file created
- Finds the % change in capitalisation between a week and the week before
- Creates csv

**Capitalisation change (week's $ change)**
- Based on the capitalisation file created
- Finds the $ change in capitalisation between a week and the week before
- Creates csv

<h1>Results</h1>

Overall this code has saved me almost 5 hours per week of work. My next steps will be as follows
- Combining the capitalisation files into one
- Calculating the % and $ changes using a loop (rather than the inbuild pct_change function in python).
- Automate the creation and updating of the master <em>NZX 50.xlsx</em> file.
