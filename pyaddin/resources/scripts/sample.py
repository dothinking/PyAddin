from .utils import context



def hello_world_1(x:str, y:str):
    return int(x) + int(y)


def hello_world_2():
    # get the workbook calling this method, then do anything with win32com
    wb = context.get_caller()
    sheet = wb.ActiveSheet

    # get cells value
    x = sheet.Range('A1').Value
    y = sheet.Range('A2').Value

    # set cell value
    sheet.Range('A3').Value = x + y

    

