import numpy as np
import openpyxl
from collections import Counter

# setting 
EXCEL_NAME = "750_packet_loss.xlsx"
SHEET_NAME1 = 'multi_0.01_20'
SHEET_NAME2 = 'multi_common_count_2'
PACKET_NUM = 750

excel_url = 'C:\\Users\\admin\\Desktop\\CDSL\\Lab\\python\\excel\\' + EXCEL_NAME
book = openpyxl.load_workbook(excel_url)
# sheet = book['sleep_' + str(SLEEP_TIME)]
sheet1 = book[SHEET_NAME1]
sheet2 = book[SHEET_NAME2]

# 100回計測
for cnt in range(2, 152):
  result = []
  cell = 'C' + str(cnt)

  for i in range(5):
    cell_value = sheet1[cell].value

    cell_ary = []
    if cell_value == "empty":
      cell_ary = []
    else:
      cell_ary = list(map(int, cell_value.split(",")[:-1]))

    # array[loss packet sequence number]
    loss_ary_bit = [0 for i in range(750)]
    for idx in cell_ary:
      if idx >= 750: continue
      loss_ary_bit[idx] = 1
    result.append(loss_ary_bit)

    # 次の受信機の受信状況を取得
    cell = chr(ord(cell[0])+2) + str(int(cell[1]) + 1)

  # numpyに変換し，各bitを加算
  numpy_bits = np.array(result)
  for i in range(4):
    numpy_bits[i+1] = numpy_bits[i+1] + numpy_bits[i]

  # それぞれの共通度の出現頻度を出力
  counter = Counter(numpy_bits[-1])
  most_common = counter.most_common()

  # cellに保存
  save_str1 = ""
  save_data = []
  most_common = sorted(most_common, key=lambda x: x[0])
  for j in range(len(most_common)):
    save_str1 +=  f'({most_common[j][0]}: {most_common[j][1]}) '
    
    percent = round(most_common[j][1] / PACKET_NUM * 100, 1)
    save_data.append(percent)

  cell2 = 'B' + str(cnt)
  sheet2[cell2] = save_str1

  for j in range(len(save_data)):
    cell2 = chr(ord('C')+j) + str(cnt)
    sheet2[cell2] = save_data[j]

book.save(excel_url) 