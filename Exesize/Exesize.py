import pandas as pd
import matplotlib.pyplot as plt
from datetime import datetime

file_path = "Exesize.xlsx"  
df = pd.read_excel(file_path)

current_year = datetime.now().year
df["Стаж"] = current_year - pd.to_datetime(df["Дата поступления"]).dt.year

def calculate_payment(row):
    if row["Стаж"] < 5:
        return row["Средний заработок"] * 0.5 * row["Количество больничных дней"]
    elif row["Стаж"] > 8:
        return row["Средний заработок"] * 1.0 * row["Количество больничных дней"]
    else:
        return row["Средний заработок"] * 0.8 * row["Количество больничных дней"]

df["К оплате по б/л"] = df.apply(calculate_payment, axis=1)

total_payment = df["К оплате по б/л"].sum()

average_experience = df["Стаж"].mean()

staff_less_than_8_years = df[df["Стаж"] < 8].shape[0]

min_sick_days = df["Количество больничных дней"].min()
max_sick_days = df["Количество больничных дней"].max()

average_sick_days_per_month = df.groupby("Месяц")["Количество больничных дней"].mean()

sick_days_by_position = df.groupby("Должность")["Количество больничных дней"].sum()

print("\nКоличество дней по больничному листу по должностям:")
print(sick_days_by_position)

plt.figure(figsize=(8, 6))
sick_days_by_position.plot(kind="bar", color="skyblue", edgecolor="black")
plt.title("Количество дней по больничному листу по должностям", fontsize=14)
plt.xlabel("Должность", fontsize=12)
plt.ylabel("Больничных дней", fontsize=12)
plt.grid(axis="y", linestyle="--", alpha=0.7)
plt.xticks(rotation=45, fontsize=10)
plt.tight_layout()
plt.savefig("sick_days_by_position_chart.png") 
plt.show()

output_file = "Updated_Exesize.xlsx"
df.to_excel(output_file, index=False)
print(f"\nРезультаты сохранены в файл: {output_file}")
