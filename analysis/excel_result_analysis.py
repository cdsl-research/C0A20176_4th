import numpy as np
import openpyxl
from collections import Counter

# setting 
EXCEL_NAME = "750_packet_loss.xlsx"
SHEET_NAME = 'uni_0.01_25'
cell_row = "70"

excel_url = 'C:\\Users\\admin\\Desktop\\CDSL\\Lab\\python\\excel\\' + EXCEL_NAME
book = openpyxl.load_workbook(excel_url)
# sheet = book['sleep_' + str(SLEEP_TIME)]
sheet = book[SHEET_NAME]

result = []
cell = 'C' + cell_row
for i in range(5):
  # recv_idのアルファベットを1進める (ex: "A" → "B")
  cell_value = sheet[cell].value
  cell_ary = list(map(int, cell_value.split(",")[:-1]))

  loss_ary_bit = [0 for i in range(750)]
  for idx in cell_ary:
    loss_ary_bit[idx] = 1

  result.append(loss_ary_bit)
  cell = chr(ord(cell[0])+2) + cell[1]

numpy_bits = np.array(result)
for i in range(4):
  numpy_bits[i+1] = numpy_bits[i+1] + numpy_bits[i]

counter = Counter(numpy_bits[-1])
most_common = counter.most_common()

print(most_common)  

# book.save(excel_url) 