# pip install pywin32
import win32print

w = 80
h = 80

graphics = b""

def black(row, col):
    x = row - h // 2
    y = col - w // 2
    r2 = x*x + y*y
    return not(r2 > 35*35 and r2 < 40*40)

for row in range(h):
    for col in range(w // 8):
        data = 0
        for bit in range(8):
            if black(row, 8*col+bit):
                data |= 1 << (7-bit)
        graphics += bytes([data])

printer_name = win32print.GetDefaultPrinter()
hprinter = win32print.OpenPrinter(printer_name)
try:
    hjob = win32print.StartDocPrinter(hprinter, 1, ("Label", None, "RAW"))
    win32print.StartPagePrinter(hprinter)
    win32print.WritePrinter(hprinter, b'N\nQ0,0\nGW' + bytes(f'0,0,{w // 8},{h},', encoding="ascii") + graphics + b'\nP1\n')
    win32print.EndPagePrinter(hprinter)
    win32print.EndDocPrinter(hprinter)
finally:
    win32print.ClosePrinter(hprinter)
