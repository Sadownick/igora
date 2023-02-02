from tkinter import Tk, StringVar, END, ttk, messagebox as mb

import pyodbc
import csv
from datetime import datetime

conn = pyodbc.connect(
    'Driver={SQL Server};'
    'Server=OTP5\SQL_MY;'
    'Database=RK7_igora;'
    'uid=sa;'
    'pwd=ucspiter'
    )

window = Tk()
window.title('Выгрузка отчета в CSV')
window.geometry('300x250')
window.resizable(width=False, height=False)

cursor = conn.cursor()
rows = []
results = []
user_discount = StringVar()
s_date = StringVar()
e_date = StringVar()
now = datetime.now()
header = ('Дата', 'Название скидки', 'Табельный номер', 'ФИО', 'Итого')


# def test():
    # print(s_date.get(), e_date.get(), user_discount.get())
    # user_discount_tk.delete('0', END)
    # s_date_tk.delete('0', END)
    # e_date_tk.delete('0', END)


def tkinter_window():
    global user_discount_tk, s_date_tk, e_date_tk
    lable_1 = ttk.Label(text='Выберите название скидки')
    lable_1.pack()
    user_discount_tk = ttk.Combobox(window, width=30, values=rows, textvariable=user_discount)
    user_discount_tk.pack()
    lable_2 = ttk.Label(text='Начальная дата в формате ГодМесяцДень (20220101)')
    s_date_tk = ttk.Entry(textvariable=s_date)
    lable_2.pack()
    s_date_tk.pack()
    lable_3 = ttk.Label(text='Конечная дата в формате ГодМесяцДень (20220131)')
    e_date_tk = ttk.Entry(textvariable=e_date)
    lable_3.pack()
    e_date_tk.pack()
    button = ttk.Button(window, text='Создать отчет...', command=main_sql_query)
    button.pack()

    window.mainloop()


def get_discount():
    global rows, user_discount
    cursor.execute(
        '''
        select NAME from DISCOUNTS where SIFR > 1001000
        '''
        )
    res = cursor.fetchall()
    for row in res:
        rows.append(row[0])


def main_sql_query():
    global s_date, e_date, user_discount, results
    try:
        u = user_discount.get()
        s = s_date.get()
        e = e_date.get()
        cursor.execute(f'''
            declare @start_date date = '{s}'
            declare @end_date date = '{e}'        
            SELECT
              format(GlobalShifts00."SHIFTDATE", 'd', 'no') AS "SHIFTDATE",
              Discounts00."NAME" AS "DISCOUNT",
              PDSCards02."TEL1" as "Tabel_Number",
              DishDiscounts00."HOLDER" AS "HOLDER",                                                                                                                
              sum(PayBindings00.PAYSUM) AS "BINDEDSUM"
    
            FROM DISCPARTS                                                                                                                                         
            LEFT JOIN PayBindings PayBindings00                                                                                                                    
              ON (PayBindings00.Visit = DiscParts.Visit) AND (PayBindings00.MidServer = DiscParts.MidServer) AND (PayBindings00.UNI = DiscParts.BindingUNI)     
            LEFT JOIN CurrLines CurrLines00                                                                                                                        
              ON (CurrLines00.Visit = PayBindings00.Visit) AND (CurrLines00.MidServer = PayBindings00.MidServer) AND (CurrLines00.UNI = PayBindings00.CurrUNI)     
            LEFT JOIN PrintChecks PrintChecks00                                                                                                                    
              ON (PrintChecks00.Visit = CurrLines00.Visit) AND (PrintChecks00.MidServer = CurrLines00.MidServer) AND (PrintChecks00.UNI = CurrLines00.CheckUNI)    
            LEFT JOIN trk7EnumsValues trk7EnumsValues0800                                                                                                          
              ON (trk7EnumsValues0800.EnumData = PrintChecks00.State) AND (trk7EnumsValues0800.EnumName = 'TDrawItemState')                                        
            LEFT JOIN DishDiscounts DishDiscounts00                                                                                                                
              ON (DishDiscounts00.Visit = DiscParts.Visit) AND (DishDiscounts00.MidServer = DiscParts.MidServer) AND (DishDiscounts00.UNI = DiscParts.DiscLineUNI) 
            LEFT JOIN Discounts Discounts00
              ON (Discounts00.SIFR = DishDiscounts00.Sifr)
            LEFT JOIN Orders Orders00                                                                                                                              
              ON (Orders00.Visit = PrintChecks00.Visit) AND (Orders00.MidServer = PrintChecks00.MidServer) AND (Orders00.IdentInVisit = PrintChecks00.OrderIdent)  
            LEFT JOIN GlobalShifts GlobalShifts00                                                                                                                  
              ON (GlobalShifts00.MidServer = Orders00.MidServer) AND (GlobalShifts00.ShiftNum = Orders00.iCommonShift)                                             
            LEFT JOIN PDSCARDS PDSCards02
              ON PDSCards02.CARDCODE=DishDiscounts00.CARDCODE
            WHERE                                                                                                                                                  
              ((PrintChecks00.State = 6) OR (PrintChecks00.State = 7))                                                                                             
              and (DishDiscounts00."CARDCODE" <> '')
              and Discounts00."NAME" = '{u}'
              and GlobalShifts00."SHIFTDATE" between @start_date and @end_date
            group by
              GlobalShifts00."SHIFTDATE",
              Discounts00."NAME",
              PDSCards02."TEL1",
              DishDiscounts00."HOLDER"
    
        '''
                       )

        res = cursor.fetchall()
        for i, row in enumerate(res, 1):
            results.append([row[0], row[1], row[2], row[3], int(row[4])])
        with open(f'{u}_{s}-{e}_{datetime.now().strftime("%H-%M-%S")}.csv', 'w') as file:
            writer = csv.writer(file, delimiter=';', lineterminator='\r')
            writer.writerow(header)
            for j in range(len(results)):
                writer.writerow(results[j])

        results = []
        user_discount_tk.delete('0', END)
        s_date_tk.delete('0', END)
        e_date_tk.delete('0', END)
        mb.showinfo('INFO', f'Файл с именем {u}_{s}-{e}_{datetime.now().strftime("%H-%M-%S")}.csv создан')
    except pyodbc.DataError:
        mb.showerror('Ошибка', 'Введите корректную дату!')
        # print('[SQL Server]Ошибка преобразования даты или времени из символьной строки')


def start_date():
    global s_date
    s_date = int(input("Начальная дата в формате YYYYMMDD:\n"))


def end_date():
    global e_date
    e_date = int(input("Конечная дата в формате YYYYMMDD:\n"))


def main():
    get_discount()
    tkinter_window()


if __name__ == "__main__":
    main()
