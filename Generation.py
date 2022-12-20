import traceback, subprocess, time, re, datetime, random, csv, os, shutil
import pandas as pd
from pythonping import ping
from ldap3 import Server, Connection

def save_log_errors(msg):
    try:
        file = ".\\log_errors.txt"
        temp = open(file, "a")
        temp.seek(0,0)
        temp.write("\n")
        temp.write(msg)
        temp.write("\n")
        temp.close()
    except:
        errors = "ОШИБКА: " + traceback.format_exc()
        return errors

# Выбор сервера AD


def function_controlerAD():
    try:
        data_out = []
        out = subprocess.run("dsquery server",stdout=subprocess.PIPE, text=True)
        hosts = str(out.stdout)
        # print(hosts)
        hosts = hosts.split("\n")
        for i in hosts:
            i = i.replace('"','')
            data_out.append(i)
        hosts = []
        for i in data_out:
            i = i.split(",")
            hosts.append(i[0])
        data_out = []
        for i in hosts:
            i = i.replace('CN=', '')
            data_out.append(i)
        ping_server = {}
        for server in data_out:
            if server == '':
                continue
            # print(server)
            response_server = ping(server, size=5, count=1)
            # print(response_server.rtt_avg_ms)
            ping_server[server] = (response_server.rtt_avg_ms)
        sorted_ping_server = sorted(ping_server.items(), key=lambda x: x[1])
        x = str(sorted_ping_server[0])
        x = x.split(",")
        # print(x[0])
        x = x[0].replace("(", "")
        data_out = x.replace("'", "")
        return data_out
    except:
        msg = str(traceback.format_exc())
        save_log_errors(msg)
        return msg


# таймер обновления сервера AD
class Timer:

    def __init__(self):
        self._start_time = None

    def start(self):


        if self._start_time is not None:
            elapsed_time = time.perf_counter() - self._start_time
            if elapsed_time > 300:
                self.temp_name_server_AD = function_controlerAD()
                self._start_time = time.perf_counter()
                # print("1:" + self.temp_name_server_AD)

            else:
                self.temp_name_server_AD
                # print("2:" + self.temp_name_server_AD)


        else:
            self._start_time = time.perf_counter()
            self.temp_name_server_AD = function_controlerAD()
            # print("3:" + self.temp_name_server_AD)


# запуск таймера
timer_AD = Timer()


# Подключение к AD
def connect_AD():
    try:
        timer_AD.start()
        temp_name_server_AD = timer_AD.temp_name_server_AD #скоректировать после перезда в облако
        # print(temp_name_server_AD)
        AD_SEARCH_TREE = 'DC=sigma,DC=sbrf,DC=ru'
        dict_login = {'login': 'sbt-filatov-vk', 'password': '<jlb,bklbyu5', 'domen': 'sigma', 'user': 'sigma' + '\\' + 'sbt-filatov-vk'}
        server_name = {'sigma': temp_name_server_AD}
        AD_SERVER = server_name['sigma']
        server = Server(AD_SERVER, use_ssl=True)
        conn = Connection(server, user=dict_login['user'], password=dict_login['password'])
        test_connect = conn.bind()
        # print(conn, AD_SEARCH_TREE)
        # print(test_connect)
        return conn, AD_SEARCH_TREE
    except:
        msg = str(traceback.format_exc())
        save_log_errors(msg)
        return msg, ""


# Отключение от AD
def disconnect_AD(conn):
    try:
        test_connect = conn.unbind()
        return 'Отключился от Active Directory'
    except:
        msg = str(traceback.format_exc())
        save_log_errors(msg)
        return msg


# запуск приложения
def start():
    try:
        generation_list_block()
        generation_list_activate()
        return "Готово"
    except:
        msg = str(traceback.format_exc())
        save_log_errors(msg)
        return msg

#\\sbt-oib-001\OIB\LockUZ
#C:\WORK\appLockerUser\

# создание списка блокировки
def generation_list_block():
    try:
        dir_hr = r"\\sbt-oib-001\OIB\LockUZ\HR"
        temp_file_hr_new = r"\\sbt-oib-001\OIB\LockUZ\files\temp_file_hr_new.csv"
        temp_file_hr_old = r"\\sbt-oib-001\OIB\LockUZ\files\temp_file_hr_old.csv"
        delta_file_hr = r"\\sbt-oib-001\OIB\LockUZ\files\delta_file_hr.csv"
        file_block = r"\\sbt-oib-001\OIB\LockUZ\files\block.csv"
        except_file = r"\\sbt-oib-001\OIB\LockUZ\files\except.csv"
        old_file_block = r"\\sbt-oib-001\OIB\LockUZ\files\old_block.csv"
        dir_files_hr = os.listdir(dir_hr)
        clear_dir_files_hr = []
        start_DB = []
        start_DB_old = []
        except_DB = []
        clear_DB = []
        clear_DB_2 = []
        clear_DB_3  = []
        idUAC = ["514","546", "66050"]
        delta_HR = []
# сохраняем block в old_block
        shutil.copyfile(file_block, old_file_block)

# ищем свежую выгрузку HR на сервере
#         print(dir_files_hr)
        for i in dir_files_hr:
            if "_Отсутствия_СБТ" in i:
                if ".xlsx" in i:
                    if "~$" not in i:
                        clear_dir_files_hr.append(i)
        clear_dir_files_hr.sort(reverse=True)
        # print(clear_dir_files_hr)

# читаем файлы выгрузки HR удаляем лишнее и сохраняем в temp_file_hr.csv
        file_name = dir_hr + "\\" + clear_dir_files_hr[0]
        # print(file_name)
        df_net_new = pd.read_excel(file_name, sheet_name=0)
        df_net_new.to_csv(temp_file_hr_new, index=False, sep = ';', encoding='windows-1251')
        # print(df_net_new)
        if clear_dir_files_hr[1]:
            file_name_2 = dir_hr + "\\" + clear_dir_files_hr[1]
            df_net_old = pd.read_excel(file_name_2, sheet_name=0)
            df_net_old.to_csv(temp_file_hr_old, index=False, sep=';', encoding='windows-1251')
            # print(df_net_old)
        else:
            pass

# читаем файл temp_file_hr_new.csv и сохраняем в переменную
        with open(temp_file_hr_new, mode="r") as r_file:
            file_read = csv.reader(r_file, delimiter=";")
            for row in file_read:
                if "ФИО сотрудника" in row:
                    pass
                else:
                    line = row[1], row[2], row[3], row[4], row[5]
                # print(line)
                    start_DB.append(line)
            # print(start_DB)

# читаем файл temp_file_hr_old.csv и сохраняем в переменную
        with open(temp_file_hr_old, mode="r") as r_file:
            file_read = csv.reader(r_file, delimiter=";")
            for row in file_read:
                if "ФИО сотрудника" in row:
                    pass
                else:
                    line = row[1], row[2], row[3], row[4], row[5]
                    # print(line)
                    start_DB_old.append(line)
            # print(start_DB_old)

# создаем дельту сотрудников из HR файлов
        for _i in start_DB_old:
            counter = 0
            # print(_i)
            for _i2 in start_DB:
                if _i[1] in _i2:
                    counter = counter + 1
            if counter == 0:
                delta_HR.append(_i)
        # print(delta_HR)
        with open(delta_file_hr, mode="w+") as w_file:
            file_write = csv.writer(w_file, delimiter=";", lineterminator="\r")
            for row in delta_HR:
                line = row[0], row[1], row[2], row[3], row[4]
                file_write.writerow(line)

# считываем исключений из файла except.csv
        with open(except_file, mode="r") as r_file:
            file_read = csv.reader(r_file, delimiter=";")
            for row in file_read:
                line = row[0]
                # print(line)
                except_DB.append(line)
            # print(except_DB)

# удаляем исключения из start_DB
        for i in start_DB:
            # print(i[1])
            if i[1] in except_DB:
                pass
            else:
                clear_DB.append(i)
        # print(clear_DB)

# проверяем clear_DB на критерии отключения УЗ и сохраняем в clear_DB_2
        for i in clear_DB:
            if i[2] == "Отпуск по уходу за ребенком" or \
                i[2] == "Неявки по невыясненным причинам, до выяснения обстоятельств" or \
                    i[2] == "Отпуск по беременности и родам":
                clear_DB_2.append([i[1],i[0]])
            elif i[2] == "Период временной нетрудоспособности (неподтвержденный листком нетрудоспособности)" or \
                i[2] == "Ежегодный основной и дополнительный оплачиваемый отпуск" or \
                i[2] == "Учебный отпуск" or \
                i[2] == "Отпуск без сохранения заработной платы":
                temp = i[3].split(".")
                temp2 = i[4].split(".")
                date_start = datetime.datetime(int(temp[2]), int(temp[1]), int(temp[0]))
                date_end = datetime.datetime(int(temp2[2]), int(temp2[1]), int(temp2[0]))
                delta = (date_end - date_start).days
                if delta > 14:
                    clear_DB_2.append([i[1], i[0]])
                else:
                    continue
            else:
                continue
        # print(clear_DB_2)

# сохраняем clear_DB_2 в файл block.csv
        with open(file_block, mode="w+") as w_file:
            file_write = csv.writer(w_file, delimiter=";", lineterminator="\r")
            for row in clear_DB_2:
                # print(len(row))
                file_write.writerow([row[0], row[1]])

# Подключаемся к AD
        connectAD, AD_SEARCH_TREE = connect_AD()

# ищем УЗ и фильтруем отключеные
#         print(len(clear_DB_2))
        for item in clear_DB_2:
            # print(item)
            i = item[1]
            # print(i)
            i = re.sub(r"\['", '*', str(i))
            i = re.sub(r"']", '', str(i))
            # print(i)
            searchtype = "employeeID"
            searchfilter = '(&(' + searchtype + '=' + str(i) + '))'
            att = ["userAccountControl", "SamAccountname"]
            connectAD.search(AD_SEARCH_TREE, searchfilter, 'SUBTREE', attributes=att)
            a = connectAD.entries
            # print(a)
            if a == []:
                pass
            else:
                for entry in a:
                    if entry.userAccountControl:
                        SamAccountname = str(entry.SamAccountname)
                        userAccountControl = str(entry.userAccountControl)
                        # print(userAccountControl)
                        # print(SamAccountname)
                        if userAccountControl in idUAC:
                           pass
                        else:
                            clear_DB_3.append([item[0],item[1],SamAccountname])
                    else:
                        pass
        # print(len(clear_DB_3))
        # print(clear_DB_3)

# отключаемся от AD
        disconnect_AD(connectAD)

# сохраняем clear_DB_3 в файл ZNO_block.csv
        with open(r"\\sbt-oib-001\OIB\LockUZ\forZNO\ZNO_block.csv", mode="w+") as w_file:
            file_write = csv.writer(w_file, delimiter=";", lineterminator="\r")
            for row in clear_DB_3:
                # print(len(row))
                file_write.writerow([row[0], row[1],row[2]])

        return
    except:
        msg = str(traceback.format_exc())
        save_log_errors(msg)
        return msg


# создание списка включения
def generation_list_activate():
    try:
        old_block = r"\\sbt-oib-001\OIB\LockUZ\files\old_block.csv"
        block = r"\\sbt-oib-001\OIB\LockUZ\files\block.csv"
        activate = r"\\sbt-oib-001\OIB\LockUZ\files\activate.csv"
        DB_old_block = []
        DB_block = []
        DB = []
        clear_DB = []
        idUAC = ["514","546", "66050"]

# Считываем и сохраняем предыдущий список блокировки old_block.csv
        with open(old_block, mode="r") as r_file:
            file_read = csv.reader(r_file, delimiter=";")
            for row in file_read:
                if row == []:
                    pass
                else:
                    # print(row)
                    line = [row[0], row[1]]
                    # print(line)
                    DB_old_block.append(line)
            # print(DB_old_block)
            # print(len(DB_old_block))

# Считываем и сохраняем новый список блокировки block.csv
        with open(block, mode="r") as r_file:
            file_read = csv.reader(r_file, delimiter=";")
            for row in file_read:
                line = [row[0], row[1]]
                # print(line)
                DB_block.append(line)
            # print(DB_block)
            # print(len(DB_block))

# Создаем дельту списков отключения
        for i in DB_old_block:
            # print(i)
            # print(type(i))
            if i in DB_block:
                pass
            else:
                DB.append(i)

        for i in DB_block:
            if i in DB_old_block:
                pass
            else:
                DB.append(i)
        # print(DB)
        # print(len(DB))

# сохраняем DB в файл activate.csv
        with open(activate, mode="w+") as w_file:
            file_write = csv.writer(w_file, delimiter=";", lineterminator="\r")
            for row in DB:
                # print(len(row))
                file_write.writerow([row[0], row[1]])

# Подключаемся к AD
        connectAD, AD_SEARCH_TREE = connect_AD()

# ищем УЗ и фильтруем активные
        # print(len(clear_DB_2))
        for item in DB:
            # print(item)
            i = item[1]
            # print(i)
            i = re.sub(r"\['", '*', str(i))
            i = re.sub(r"']", '', str(i))
            # print(i)
            searchtype = "employeeID"
            searchfilter = '(&(' + searchtype + '=' + str(i) + '))'
            att = ["userAccountControl", "SamAccountname"]
            connectAD.search(AD_SEARCH_TREE, searchfilter, 'SUBTREE', attributes=att)
            a = connectAD.entries
            # print(a)
            if a == []:
                pass
            else:
                for entry in a:
                    if entry.userAccountControl:
                        SamAccountname = str(entry.SamAccountname)
                        userAccountControl = str(entry.userAccountControl)
                        # print(userAccountControl)
                        # print(SamAccountname)
                        if userAccountControl in idUAC:
                            clear_DB.append([item[0], item[1], SamAccountname])
                        else:
                            pass
                    else:
                        pass
        # print(len(clear_DB))
        # print(clear_DB)

        # сохраняем clear_DB в файл ZNO_activate.csv
        with open(r"\\sbt-oib-001\OIB\LockUZ\forZNO\ZNO_activate.csv", mode="w+") as w_file:
            file_write = csv.writer(w_file, delimiter=";", lineterminator="\r")
            for row in clear_DB:
                # print(len(row))
                file_write.writerow([row[0], row[1], row[2]])

# отключаемся от AD
        disconnect_AD(connectAD)

    except:
        msg = str(traceback.format_exc())
        save_log_errors(msg)
        return msg


if __name__ == '__main__':
    start = start()
    print(start)