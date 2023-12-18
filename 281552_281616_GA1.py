import os
import pandas as pd 
import matplotlib.pyplot as plt

# Clear the screen
os.system('cls')

# Get the current directory where the Python file is located
current_directory = os.path.dirname(os.path.abspath("c:\\Users\\System Manager\\Desktop\\U U M\\S E M 7\\SQIT3073 - BUSINESS ANALYTICS PROGRAMMING\\281552_281616_GA1"))

# Create folders and subfolders
main_folder = os.path.join(current_directory, '281552_281616_GA1')
os.makedirs(main_folder, exist_ok=True)

# Read the Excel file
excel_file_path = pd.read_excel("c:\\Users\\System Manager\\Desktop\\U U M\\S E M 7\\SQIT3073 - BUSINESS ANALYTICS PROGRAMMING\\GA\\reserve_money.xlsx", sheet_name = "original_data")

#Convert data into data frame
df = pd.DataFrame(excel_file_path)

# Create an Excel writer object
with pd.ExcelWriter(os.path.join(main_folder,'cleaned_reserve_money.xlsx')) as writer:
    
    # Save original DataFrame to Excel
    df.to_excel(writer, sheet_name='Original_Data', index=False)
    
    #Grouping and Sum to get total value
    grouped_df = df.groupby('End of period (month)')['Total Reserve Money'].sum().reset_index()
    grouped_df.to_excel(writer, sheet_name='Grouped_Data', index = False)
    
    #Filtering for selected years
    filtered_df = df[df['End of period (month)'] >= 2021][['End of period (month)', 'Month','Total Reserve Money']]
    filtered_df.to_excel(writer, sheet_name='Filtered_Data', index=False)

#Declaration for chart 1 
years = grouped_df['End of period (month)']
reserve_money = grouped_df['Total Reserve Money']    

#Chart1 - Total Reserve Money for 10 years
plt.figure(figsize=(10, 6))

plt.plot(years, reserve_money, color='green', marker='x', linestyle='-', linewidth=2, markersize=8, label='Total Reserve Money')

plt.xlabel('End of period (Year)')
plt.ylabel('Total Reserve Money (RM)')
plt.title('Total Reserve Money for 10 years (2013-2022)')
plt.legend()
plt.grid(True)

#Declaration for chart 2 
# Ensure 'Month' is categorical with the correct ordering
filtered_df['Month'] = pd.Categorical(filtered_df['Month'], categories=range(1, 13), ordered=True)

# Separate data for two years
Month1 = filtered_df.loc[filtered_df['End of period (month)'] == 2021, 'Month']
Month2 = filtered_df.loc[filtered_df['End of period (month)'] == 2022, 'Month']
selected_total = filtered_df.loc[filtered_df['End of period (month)'] == 2021, 'Total Reserve Money']
selected_total2 = filtered_df.loc[filtered_df['End of period (month)'] == 2022, 'Total Reserve Money']

#Chart2 - Comparative Analysis of Monthly Total Reserve Money for 2021 and 2022
plt.figure(figsize=(10, 6))

plt.plot(Month1, selected_total, color='blue', marker='x', linestyle='-', linewidth=2, markersize=8, label='Total Monthly Reserve Money in 2021')
plt.plot(Month2, selected_total2, color='orange', marker='x', linestyle='-', linewidth=2, markersize=8, label='Total Monthly Reserve Money in 2022')

plt.xlabel('End of period (Month)')
plt.ylabel('Total Monthly Reserve Money (RM)')
plt.title('Comparative Analysis of Monthly Total Reserve Money for 2021 and 2022')
plt.legend()
plt.grid(True)

plt.show()


