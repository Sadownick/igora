import pyodbc
import csv
from datetime import datetime

conn = pyodbc.connect(
    'Driver={SQL Server};'
    'Server=PREDATOR\SQLEXPRESS;'
    'Database=RK7;'
    'uid=sa;'
    'pwd=ucspiter'
    )

cursor = conn.cursor()
rows = []
results = []
user_discount = ''
s_date = ''
e_date = ''
now = datetime.now()
header = ('Дата', 'Название скидки', 'Табельный номер', 'ФИО', 'Итого')


def get_discount():
    global rows, user_discount
    cursor.execute(
        '''
        select NAME from DISCOUNTS where SIFR > 1000000
        '''
        )
    res = cursor.fetchall()
    for row in res:
        rows.append(row[0])
    for i, n in enumerate(range(1, len(rows) + 1)):
        print(n, rows[i])
    user_input_discount = int(input('\n\nВыберите подразделение: \n'))
    user_discount = str(rows[user_input_discount - 1])


def main_sql_query():
    global s_date, e_date
    cursor.execute(f'''
        declare @start_date date = '{s_date}'
        declare @end_date date = '{e_date}'        
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
          and Discounts00."NAME" = '{user_discount}'
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
        # print(i, row[0], row[1], row[2], row[3], int(row[4]))
    with open(f'{user_discount}_{datetime.now().strftime("%d-%m-%Y_%H-%M-%S")}.csv', 'w') as file:
        writer = csv.writer(file, delimiter=';', lineterminator='\r')
        writer.writerow(header)
        for j in range(len(results)):
            writer.writerow(results[j])


def start_date():
    global s_date
    s_date = input("Начальная дата в формате YYYYMMDD:\n")


def end_date():
    global e_date
    e_date = input("Конечная дата в формате YYYYMMDD:\n")


def main():
    get_discount()
    start_date()
    end_date()
    print('Идет обработка запроса, ожидайте...')
    main_sql_query()


if __name__ == "__main__":
    main()
