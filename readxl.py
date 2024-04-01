import pandas as pd

# Read the Excel file
df = pd.read_excel('Montreal_Cognitive_Assessment__MoCA__08Dec2023.xlsx')

# Define a function to classify the sum
def classify_sum(sum_value):
    if sum_value > 25 and sum_value <= 30:
        return 'normal'
    elif sum_value >= 18 and sum_value <= 25:
        return 'Mild cognitive impairment'
    elif sum_value >= 10 and sum_value < 18:
        return 'Moderate cognitive impairment'
    elif sum_value >= 0 and sum_value < 10:
        return 'Severe cognitive impairment'
    else:
        return 'Unknown'

# Iterate over each row and calculate the sum, then classify and append to a new column
for index, row in df.iterrows():
    # sum_value = row[['MCAALTTM', 'MCACUBE',	'MCACLCKC',	'MCASER7', 'MCACLCKH', 'MCALION', 'MCARHINO', 'MCACAMEL',	'MCAFDS', 'MCABDS',	'MCAVIGIL',	'MCASER7',	'MCASNTNC', 'MCAVF', 'MCAABSTR', 'MCAREC1',	'MCAREC2', 'MCAREC3','MCAREC4',	'MCAREC5',	'MCADATE',	'MCAMONTH',	'MCAYR', 'MCADAY',	'MCAPLACE',	'MCACITY']].sum()
    sum_value = row[['MCATOT']].sum()
    classification = classify_sum(sum_value)
    df.at[index, 'Calculated'] = sum_value
    df.at[index, 'Classification'] = classification
    # if index == 0:
    #     print(df)
    #     print(row)
    #     print('record #',index, 'value = ', sum_value)
    #     break;
# Save the updated DataFrame back to Excel
df.to_excel('output_file.xlsx', index=False)
