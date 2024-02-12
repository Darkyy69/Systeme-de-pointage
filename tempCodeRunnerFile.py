def row_col_dic():

    row_sheet_map = {i: i + 13 for i in range (1, 13)}
    print(row_sheet_map)
    print('----------------')

    col_sheet_map = {i: chr(i + ord('B') - 1 ) for i in range(1, 26)}
    dic2 = {i: 'A' + chr(i+39) for i in range(26, 32)}
    col_sheet_map.update(dic2)
    print(col_sheet_map)

row_col_dic()  