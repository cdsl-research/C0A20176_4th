import time
import socket
import threading
import math
import openpyxl

# basic setting
LOCALHOST = socket.gethostbyname(socket.gethostname())
multicast_grp = "239.1.2.3"
PORT = 1234

# send setting
UNDER = 2
ROUND = 150
SLEEP_TIME = 0.01
INTERVAL = 10
SEPARATE_SIZE = 1000
SET_TIMEOUT = 7
RECV_NUM = 5

excel_name = str(750) + "_" + "packet_loss.xlsx"
excel_url = f'C:\\Users\\admin\\Desktop\\CDSL\\Lab\\python\\excel\\{excel_name}'
book = openpyxl.load_workbook(excel_url)
sheet_prefix = 'multi'
sheet_name = sheet_prefix + "_" + str(SLEEP_TIME) + "_" + str(INTERVAL)
print("----------------------------")
print("excelName: ", excel_name)
print("sheetName: ", sheet_name)
print("----------------------------")
sheet = book[sheet_name]


start_t = 0
end_t = []
recv_status = []
read_data = ""
data_len = 0


chunks = 0
max_seq = 0
FILE_NAME = './sendFile/binary_750k.txt'
with open(FILE_NAME, 'r') as f:
    read_data = f.read()
    data_len = len(read_data)
    chunks = [read_data[j:j+SEPARATE_SIZE] for j in range(0, data_len, SEPARATE_SIZE)]
    max_seq = str(math.ceil(data_len / SEPARATE_SIZE))


# ソケットを作成
sock = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
sock.bind((LOCALHOST, PORT))
sock.settimeout(SET_TIMEOUT)

# クライアントからのメッセージを受信して返信する関数
def respond():
    # クライアントからのデータを受信
    while True:
        try:
            data, addr = sock.recvfrom(1024)
            # 受信タイムの計測
            end_t.append(time.time())

            recv_id, loss_seq = data.decode().split(":")
            print(f'Received message from {recv_id}:{addr}')
            print(loss_seq)

            recv_status.append({
                "recv_id": recv_id,
                "loss_seq": loss_seq
            })
        except Exception as e:
            print(f"error: {e}")
            break


# 計測開始
for i in range(UNDER, ROUND+UNDER):
    start_t = time.time()

    # 受信は別スレッドに委任
    respond_thread = threading.Thread(target=respond)
    respond_thread.start()

    print(f"start {i-(UNDER-1)}回目")

    seq = 1
    for chunk in chunks:
        chunk = max_seq + ":" + str(seq) + ":" + chunk
        sock.sendto(chunk.encode(), (multicast_grp, PORT))
        seq += 1

        if not seq % INTERVAL:
            time.sleep(SLEEP_TIME)

    # スレッドが終了するまで待機
    respond_thread.join()

    print("----------")
     # 受信状況の記録
    for st in recv_status:
        recv_id = st["recv_id"]
        loss_seq = st["loss_seq"]
        loss_num = 0

        cell = recv_id + str(i)

        if loss_seq == "done" :
            sheet[cell] = "empty"
        else:
            loss_packet = ""

            # パケットロスを判定し，記録する
            for p in range(len(loss_seq)):
                if loss_seq[p] == "0":
                    loss_packet += str(p+1) + ","
                    loss_num += 1
            sheet[cell] = loss_packet

        # recv_idのアルファベットを1進める (ex: "A" → "B")
        cell = chr(ord(recv_id)+1) + str(i)
        # ロスしたパケットの総数
        sheet[cell] = loss_num
        print(recv_id, loss_num)



    # 送信完了時間の記録
    cell = "B" + str(i)

    res_time = 0
    for ed in end_t:
        res_time = max(res_time, ed)
    res_time = round(res_time - start_t, 3)

    if res_time < 0:
        res_time = "error"
    sheet[cell] = res_time

    print(f'time: {res_time}')

    end_t = []
    recv_status = []

    time.sleep(2)

sock.close()
book.save(excel_url)