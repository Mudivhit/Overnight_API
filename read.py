import pandas as pd
df = pd.read_json (r'zxc.json')
export_csv = df.to_csv (r'zxc.csv', index = None, header=True)