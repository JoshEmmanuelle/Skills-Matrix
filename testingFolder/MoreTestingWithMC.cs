Create a simple pandas dataframe using python and print it out.

import pandas as pd
# Create a simple dataframe
data = {
    'Name': ['Alice', 'Bob', 'Charlie'],
    'Age': [25, 30, 35],
    'City': ['New York', 'Los Angeles', 'Chicago']
}
df = pd.DataFrame(data)
# Print the dataframe
print(df)

