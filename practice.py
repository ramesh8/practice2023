import pandas as pd
fname = "files\\RRI127-Report 1.xlsx"

df = pd.read_excel(fname)

df.drop(df.head(2).index, inplace=True)
df.drop(df.tail(2).index, inplace=True)

# df.dropna()

df.dropna(how="all", axis=0, inplace=True) #row
df.dropna(how="all", axis=1, inplace=True) #columns


# print(df)
# print(df.columns)
merge_columns = [2, 4, 6]

# print(df.iloc[:,2])
for col in merge_columns:
    df.iloc[:,col+1] = df.iloc[:,col+1].combine_first(df.iloc[:,col])
    
df.drop(columns=df.columns[merge_columns], inplace=True)



df.to_excel("output.xlsx", index=None, header=None)
