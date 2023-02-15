import numpy as np
import matplotlib.pyplot as plt
from openpyxl import load_workbook

data_file = 'data_13feb_perros.xlsx'

wb = load_workbook(data_file)

ws = wb['Sheet1']
all_rows = list(ws.rows)

data = []

for row in all_rows[1:120]:
    value = row[1].value
    if (value is not None):
        data.append(value)




mu = np.mean(data)

sigma = np.std(data, ddof=1)

n, bins, patches = plt.hist(data, bins=40, density=True, alpha=0.6, color='r')

y = ((1/(np.sqrt(2*np.pi)*sigma)) * np.exp(-0.5*(1/sigma*(bins-mu))**2))

plt.plot(bins, y, '--', color='black')
  
plt.xlabel('Edad', fontweight ="bold")
plt.ylabel('Numero', fontweight ="bold")
  
plt.title('Edad de los perrunos', fontweight ="bold")
  
plt.show()




