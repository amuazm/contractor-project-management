import numpy as np
import matplotlib.pyplot as plt
import pandas as pd

df_reports = pd.read_excel("./Excel Files/Output/Result_ARIMA.xlsx", sheet_name="Reports")

project_ids = list(df_reports["Project ID"].unique())

for project_id in project_ids:
    df = df_reports[df_reports["Project ID"] == project_id]
    df = df[["Date", "Completion"]]
    df = df.set_index("Date")
    print(df)
    
    plt.plot(df.index, df["Completion"])
    plt.show()