import mcrcon
import openpyxl
from datetime import datetime

def read_data_from_file():
    with open("data.txt", "r") as file:
        data = file.readlines()
        host = data[0].strip()
        port = int(data[1].strip())
        password = data[2].strip()
    return host, port, password

def send_rcon_command(host, port, password, command):
    client = mcrcon.MCRcon(host, password, port)
    try:
        client.connect()
        response = client.command(command)
        print(response)
        log_command(command, response)
    except Exception as e:
        print(f"Ошибка при выполнении команды: {e}")

def log_command(command, response):
    now = datetime.now()
    current_date = now.strftime("%Y-%m-%d")
    current_time = now.strftime("%H:%M:%S")
    log_file = f"log_{current_date}.xlsx"
    try:
        wb = openpyxl.load_workbook(log_file)
    except FileNotFoundError:
        wb = openpyxl.Workbook()
        sheet = wb.active
        sheet.append(["ДАТА", "ВРЕМЯ", "КОМАНДА"])
    sheet = wb.active
    sheet.append([current_date, current_time, command])
    wb.save(log_file)

def main():
    host, port, password = read_data_from_file()

    print("Введите 'exit', чтобы выйти из программы.")

    while True:
        command = input("Введите команду для отправки на сервер: ")
        if command.lower() == 'exit':
            break
        else:
            send_rcon_command(host, port, password, command)

if __name__ == "__main__":
    main()
