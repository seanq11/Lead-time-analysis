import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import numpy as np


# Read file and create variables for each column
df_excel = pd.read_excel('DATA.xlsx')
date_created = df_excel['Date Created']
date_started = df_excel['Date Started']
date_completed = df_excel['Date Completed']
work_item = df_excel['Work Item']
state = df_excel['State']
team = df_excel['Team']
application = df_excel['Application']


# Calculate lead and cycle
df_excel['Lead Time'] = (date_completed - date_started).dt.days
df_excel['Cycle Time'] = (date_completed - date_created).dt.days

# Group by Application and Team
grouped = df_excel.groupby(['Application', 'Team']).agg({'Lead Time': 'mean', 'Cycle Time': 'mean'}).reset_index()

# Pivot data for plotting
pivot_df = grouped.pivot(index='Application', columns='Team', values=['Lead Time', 'Cycle Time'])


# Colours for each team
colours = {'Red': 'red', 'Blue': 'blue', 'Green': 'green', 'Yellow': 'yellow'}
team_colours = [colours.get(x, '#333333') for x in pivot_df.columns.levels[1]]


def BarGraph():
    ax = pivot_df.plot(kind='bar', color=team_colours, figsize=(10, 6),width=1)
    plt.title('Average Lead Time and Cycle Time per Application by Team')
    plt.xlabel('Application')
    plt.ylabel('Days')
    plt.legend([plt.Rectangle((0, 1), 1, 1, color=colours[team]) for team in colours], colours.keys())
    plt.show()


# Define start_date and end_date based on the DataFrame's date range
start_date = df_excel['Date Completed'].min()
end_date = df_excel['Date Completed'].max()

filtered_df = df_excel[df_excel['Date Completed'].notna()]


# Plotting the lead time for each project
fig, ax = plt.subplots(figsize=(12, 6))

# Iterate over each team and plot their projects
for team in filtered_df['Team'].unique():
    team_df = filtered_df[filtered_df['Team'] == team]
    ax.plot(team_df['Date Completed'], team_df['Lead Time'], marker='o', linestyle='-', label=team)

# Formatting the plot
ax.set_xlim(start_date, end_date)  # Set x-axis limits to the range of completed dates
ax.xaxis.set_major_locator(mdates.MonthLocator(interval=1))  # Monthly interval
ax.xaxis.set_major_formatter(mdates.DateFormatter('%Y-%m'))
plt.xticks(rotation=45)  # Rotate x-axis labels for readability

def LineGraph():
    plt.title('Team Lead Time Progress Over Time')
    plt.xlabel('Date Completed')
    plt.ylabel('Lead Time (days)')
    plt.legend()
    plt.grid(True)
    plt.tight_layout()  # Adjust layout
    plt.show()

BarGraph()
LineGraph()
