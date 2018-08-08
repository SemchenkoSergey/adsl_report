# coding: utf-8

min_row = 5

def sessions_report(wb, cursor):
    sh = wb['Разрывы']
    command = '''
    SELECT
     ad.locality 'Нас. пункт',
     ad.street 'Улица',
     ad.house_number 'Дом',
     ad.apartment_number  '№ кв.',
     RIGHT(ad.phone_number, 5) 'Телефон',
     ad.tariff 'Тариф',
     ROUND(AVG(dd.max_dw_rate)) 'Скорость',
     ds.count '1 день',
     ds2.count '2 дня',
     ds3.count '3 дня'
    FROM abon_dsl ad INNER JOIN data_dsl dd
     ON ad.hostname=dd.hostname
     AND ad.board=dd.board
     AND ad.port=dd.port
     INNER JOIN data_sessions ds
      ON ad.account_name = ds.account_name
      AND ds.date = CAST(dd.datetime AS DATE)
      LEFT OUTER JOIN data_sessions ds2
       ON ad.account_name = ds2.account_name
       AND ds2.date = DATE_ADD(CAST(dd.datetime AS DATE), INTERVAL -1 DAY)
       LEFT OUTER JOIN data_sessions ds3
        ON ad.account_name = ds3.account_name
        AND ds3.date = DATE_ADD(CAST(dd.datetime AS DATE), INTERVAL -2 DAY)
    WHERE
     ds.date = DATE_ADD(CURRENT_DATE(), INTERVAL -1 DAY)
     AND ad.area LIKE '%Петровск%'
     AND ds.count > 30
    GROUP BY ad.phone_number
    ORDER BY ds.count DESC, ds2.count DESC, ds3.count DESC
    '''
    cursor.execute(command)
    result = cursor.fetchall()
    cur_row = min_row
    for data in result:
        sh['A{}'.format(cur_row)].value = data[0]
        sh['B{}'.format(cur_row)].value = data[1]
        sh['C{}'.format(cur_row)].value = data[2]
        sh['D{}'.format(cur_row)].value = data[3]
        sh['E{}'.format(cur_row)].value = data[4]
        sh['F{}'.format(cur_row)].value = data[5]
        sh['G{}'.format(cur_row)].value = data[6]
        sh['H{}'.format(cur_row)].value = data[7]
        sh['I{}'.format(cur_row)].value = data[8]
        sh['J{}'.format(cur_row)].value = data[9]
        cur_row += 1
    print('Построен отчет по разрывам')


def speed_report(wb, cursor):
    sh = wb['Скорость']
    command = '''
    SELECT
     ad.locality 'Нас. пункт',
     ad.street 'Улица',
     ad.house_number 'Дом',
     ad.apartment_number '№ кв.',
     RIGHT(ad.phone_number, 5) 'Номер тел.',
     ad.tariff 'Тариф',
     ROUND(AVG(dd.max_dw_rate)) 'Ср. прих. скор.',
     ROUND(AVG(dd.max_dw_rate)/ad.tariff, 2) 'Отнош. скор./тариф'
    FROM abon_dsl ad INNER JOIN data_dsl dd
     ON ad.hostname=dd.hostname
     AND ad.board=dd.board
     AND ad.port=dd.port
    WHERE
     CAST(dd.datetime AS DATE) = DATE_ADD(CURRENT_DATE(), INTERVAL -1 DAY)
     AND ad.tariff IS NOT NULL
     AND ad.area LIKE '%Петровск%'
    GROUP BY ad.phone_number
    HAVING ROUND(AVG(dd.max_dw_rate)/ad.tariff, 2) < 1
    ORDER BY 8, ad.street, CAST(ad.house_number AS INT), ad.apartment_number
    '''
    cursor.execute(command)
    result = cursor.fetchall()
    cur_row = min_row
    for data in result:
        sh['A{}'.format(cur_row)].value = data[0]
        sh['B{}'.format(cur_row)].value = data[1]
        sh['C{}'.format(cur_row)].value = data[2]
        sh['D{}'.format(cur_row)].value = data[3]
        sh['E{}'.format(cur_row)].value = data[4]
        sh['F{}'.format(cur_row)].value = data[5]
        sh['G{}'.format(cur_row)].value = data[6]
        sh['H{}'.format(cur_row)].value = data[7]
        cur_row += 1
    print('Построен отчет по скоростям')


def modems_report(wb, cursor):
    sh = wb['Модемы']
    command = '''
    SELECT
     ad.locality 'Нас. пункт',
     ad.street 'Улица',
     ad.house_number 'Дом',
     apartment_number '№ кв.',
     RIGHT(ad.phone_number, 5) 'Номер тел.'
    FROM abon_dsl ad INNER JOIN data_dsl dd
     ON ad.hostname=dd.hostname
     AND ad.board=dd.board
     AND ad.port=dd.port
    WHERE
     CAST(dd.datetime AS DATE) = DATE_ADD(CURRENT_DATE(), INTERVAL -1 DAY)
     AND ad.area LIKE '%Петровск%'
     AND ad.account_name IN (
      SELECT account_name
      FROM data_sessions
      WHERE date = DATE_ADD(CURRENT_DATE(), INTERVAL -1 DAY)
       AND count = 0
     )
    GROUP BY ad.phone_number
    HAVING AVG(dd.max_dw_rate) IS NOT NULL
    ORDER BY ad.locality, ad.street
    '''
    cursor.execute(command)
    result = cursor.fetchall()
    cur_row = min_row
    for data in result:
        sh['A{}'.format(cur_row)].value = data[0]
        sh['B{}'.format(cur_row)].value = data[1]
        sh['C{}'.format(cur_row)].value = data[2]
        sh['D{}'.format(cur_row)].value = data[3]
        sh['E{}'.format(cur_row)].value = data[4]
        cur_row += 1
    print('Построен отчет по модемам')

