import s_importer




if __name__ == "__main__":
    print('\n       Привет!  \n       \(^o^)/ \n')

print("Выгружать колонки:\n 1.Номер \n 2.Дата изменения \n",
      "3.Дата \n 4.Клиент\n 5.Инициатор\n 6.Описание\n",
      "7.Ответственный\n 8.Состояние \n\n")



s_importer.import_global_types()
s_importer.import_input()

input('\n Нажми что нибудь для выхода')






