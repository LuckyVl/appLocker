import traceback, subprocess, time, re, csv
from pythonping import ping
from ldap3 import Server, Connection
from ldap3 import MODIFY_REPLACE


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
        print("Найден сервер AD")
        return data_out
    except:
        errors = [["Ошибка в функции function_controlerAD"],
                  ["ОШИБКА: " + traceback.format_exc()]]
        print("Сервер AD не найден")
        return errors


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
        AD_SEARCH_TREE = 'DC=sbertech,DC=local'
        dict_login = {'login': 'SА-CSRS-AD', 'password': '+xZJx0/zF1Q]RL=T{(mwQa_W)', 'domen': 'sbertech', 'user': 'sbertech' + '\\' + 'SА-CSRS-AD'}
        server_name = {'sbertech': temp_name_server_AD}
        AD_SERVER = server_name['sbertech']
        server = Server(AD_SERVER, use_ssl=True)
        conn = Connection(server, user=dict_login['user'], password=dict_login['password'])
        test_connect = conn.bind()
        # print(conn, AD_SEARCH_TREE)
        # print(test_connect)
        print("Подключился к AD")
        return conn, AD_SEARCH_TREE
    except:
        erorrs = ['Ошибка в функции connect_AD: ' + traceback.format_exc()]
        print("Подключиться к AD не удалось")
        return erorrs, ""


# Отключение от AD
def disconnect_AD(conn):
    try:
        test_connect = conn.unbind()
        print("Отключился от AD")
        return 'Отключился от Active Directory'
    except:
        erorrs = ['Ошибка в функции disconnect_AD: ' + traceback.format_exc()]
        print("Отключиться от AD не удалось")
        return erorrs


# Загрузка из файла в скрипт данных
def function_download_tabnom():
    list_block = ".\\block.csv"
    list_activate = ".\\activate.csv"
    temp_list_block = []
    temp_list_activate = []
    try:

        with open(list_block, mode="r") as r_file:
            file_read = csv.reader(r_file, delimiter=";")
            for row in file_read:
                temp_list_block.append(row[1])
            r_file.close()

        with open(list_activate, mode="r") as r_file:
            file_read = csv.reader(r_file, delimiter=";")
            for row in file_read:
                temp_list_activate.append(row[1])
            r_file.close()

        # print(temp_list_block, temp_list_activate)
        print("Считал block.csv и activate.csv")
        return temp_list_block, temp_list_activate
    except:
        errors = "ОШИБКА: " + traceback.format_exc()
        print("Считать block.csv и activate.csv не удалось")
        return errors, ''


# Поиск логина по табельному номеру
def function_ad_search_tabnom(connectAD,temp_list_block,temp_list_activate,AD_SEARCH_TREE):
    # print(temp_list_block)
    # print(temp_list_activate)
    temp_list = []
    temp_list2 = []
    DN_b = ""
    DN_a = ""
    try:
        for i in temp_list_block:
            # print(i)
            i = re.sub(r"\['", "", i)
            i = re.sub(r"']", "", i)
            # print(i)
            searchtype = "employeeID"
            searchfilter = '(&(' + searchtype + '=' + str(i) + '))'
            att = ["distinguishedName", "userAccountControl","SamAccountname"]
            connectAD.search(AD_SEARCH_TREE, searchfilter, 'SUBTREE', attributes=att)
            a = connectAD.entries
            # print(a)
            if a == []:
                pass
            else:
                for entry in a:
                    if entry.userAccountControl:
                        SamAccountname = str(entry.SamAccountname)
                        if "SP-" in SamAccountname or "SA-" in SamAccountname or\
                                "sp-" in SamAccountname or "sa-" in SamAccountname:
                            pass
                        else:
                            DN = str(entry.distinguishedName)
                            temp_list.append(DN)
                    else:
                        continue

            DN_b = temp_list
        # print(DN_b)

        for i in temp_list_activate:
            # print(i)
            i = re.sub(r"\['", '*', i)
            i = re.sub(r"']", '', i)
            # print(i)
            searchtype = "employeeID"
            searchfilter = '(&(' + searchtype + '=' + str(i) + '))'
            att = ["distinguishedName", "userAccountControl","SamAccountname"]
            connectAD.search(AD_SEARCH_TREE, searchfilter, 'SUBTREE', attributes=att)
            a = connectAD.entries
            # print(a)
            if a == []:
                pass
            else:
                for entry in a:
                    if entry.userAccountControl:
                        SamAccountname = str(entry.SamAccountname)
                        if "SP-" in SamAccountname or "SA-" in SamAccountname or\
                                "sp-" in SamAccountname or "sa-" in SamAccountname:
                            pass
                        else:
                            DN2 = str(entry.distinguishedName)
                            temp_list2.append(DN2)
                    else:
                        continue
            DN_a = temp_list2
        # print(DN_a)
        print("Найдены логины")
        return DN_b, DN_a

    except:
        errors = "ОШИБКА: " + traceback.format_exc()
        print("Поиск логинов не удался")
        return errors, ''


# Включение УЗ
def function_activate(connectAD,DN_a):
    # print(DN_a)
    for i in DN_a:
        try:
            connectAD.modify(i, {"userAccountControl": [(MODIFY_REPLACE, [512])],
                                  }
                             )
            result = connectAD.result
            # print(result)

        except:
            errors = "ОШИБКА: " + traceback.format_exc()
            print("Включение УЗ не удалось")
            return errors

        try:
            connectAD.modify(i, {"description": [(MODIFY_REPLACE),[]],
                                  }
                             )
            result = connectAD.result
            # print(result)
        except:
            errors = "ОШИБКА: " + traceback.format_exc()
            print("Не удалось изменить описание УЗ при включении")
            return errors
    print("УЗ включены")


# Отключение УЗ
def function_deactivate(connectAD,DN_b):
    file = ".\\description.txt"
    # print(file)
    try:
        temp = open(file, "r")
        description = temp.readline()
        temp.close()
        # print(description)
        print("Считал description.txt")
    except:
        errors = "ОШИБКА: " + traceback.format_exc()
        description = ""
        print("НЕ считал description.txt")
        return errors, description

    for i in DN_b:
        # print(connectAD)
        # print(i)
        try:
            connectAD.modify(i, {"userAccountControl": [(MODIFY_REPLACE, [514])],
                                  }
                             )
            result = connectAD.result
            # print(result)
        except:
            errors = "ОШИБКА: " + traceback.format_exc()
            print("НЕ смог отключить УЗ")
            return errors
        # print(description)
        try:
            connectAD.modify(i, {"description": [(MODIFY_REPLACE),[description]],
                                  }
                             )
            result = connectAD.result
            # print(result)
        except:
            errors = "ОШИБКА: " + traceback.format_exc()
            print("НЕ смог изменить описание УЗ")
            return errors
    print("УЗ отключены")

if __name__ == '__main__':
    step1 = function_controlerAD()
    print(step1)
    connectAD, AD_SEARCH_TREE = connect_AD()
    print(connectAD)
    print(AD_SEARCH_TREE)
    temp_list_block, temp_list_activate = function_download_tabnom()
    print(temp_list_block)
    print(temp_list_activate)
    DN_b,DN_a = function_ad_search_tabnom(connectAD,temp_list_block,temp_list_activate,AD_SEARCH_TREE)
    print(DN_b)
    print(DN_a)
    step5 = function_deactivate(connectAD,DN_b)
    print(step5)
    step6 = function_activate(connectAD,DN_a)
    print(step6)
    stepx = disconnect_AD(connectAD)
    print(stepx)

