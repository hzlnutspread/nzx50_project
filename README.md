<h1>About the project</h1>

I was tasked with a project to create profiles for all 50 companies in the NZX50. This included creating my own weighted index based on the NZX50 Smartshares fund. This is the master file. In order to keep this up to date I create a new sheet and update the share prices of each company every Friday. This is done in the <em>NZX50.xlsx</em> file. I was further given the task of creating a series of 5 charts charts in drupal 8. This requires me to create and update 5 csvs for each company to upload to the company server. The 5 charts for each company that would be uploaded to the website are created by the 3 .py files you see. 

- <em>Create_share_price.py

- Create_rank.py

- Create_cap.py</em>

In addition to other manual tasks I do for this job, this meant I had to update an additional 250 csv files every friday. As such, with no prior coding experience I decided to try and automate this process using Python to save myself time.

<h1>About the 5 csv files</h1>

**Share price:**
- Pulled values from csv link on Yahoo Finance.
- Sorted to pull only share prices that fall on Fridays
- Creates csv

**Rank:** 
- Based on my own master file called <em>'NZX 50.xlsx'</em>
- Pulls the rank of each ticker
- Creates csv

**Capitalisation:**
- Based on my own master file called <em>'NZX 50.xlsx'</em>
- Pulls capitalisation of each ticker
- Creates csv

**Capitalisation change (week's % change):**
- Based on the capitalisation file created under **Capitalisation**
- Finds the % change in capitalisation between a week and the week before
- Creates csv

**Capitalisation change (week's $ change)**
- Based on the capitalisation file created under **Capitalisation**
- Finds the $ change in capitalisation between a week and the week before
- Creates csv

<h1>Results</h1>

Overall this code has saved me almost 5 hours per week of work. 

The way I plan to improve on the code is as follows:
- Combining the capitalisation files into one <em>(completed)</em>
- Calculating the % and $ changes using a loop (rather than the inbuild pct_change function in python).
- Automate the creation and updating of the master <em>NZX 50.xlsx</em> file.
