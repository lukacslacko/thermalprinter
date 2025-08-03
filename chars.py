# pip install pywin32
import win32print

# 8 dots per millimeter
# 54 millimeter width

CHARS = {
    "A": [
        " ## ",
        "#  #",
        "####",
        "#  #",
        "#  #",
    ],
    "B": [
        "### ",
        "#  #",
        "### ",
        "#  #",
        "### ",
    ],
    "C": [
        " ###",
        "#   ",
        "#   ",
        "#   ",
        " ###",
    ],
    "D": [
        "### ",
        "#  #",
        "#  #",
        "#  #",
        "### ",
    ],
    "E": [
        "####",
        "#   ",
        "### ",
        "#   ",
        "####",
    ],
    "F": [
        "####",
        "#   ",
        "### ",
        "#   ",
        "#   ",
    ],
    "G": [
        " ###",
        "#   ",
        "# ##",
        "#  #",
        " ## ",
    ],
    "H": [
        "#  #",
        "#  #",
        "####",
        "#  #",
        "#  #",
    ],
    "I": [
        "#",
        "#",
        "#",
        "#",
        "#",
    ],
    "J": [
        "   #",
        "   #",
        "   #",
        "#  #",
        " ## ",
    ],
    "K": [
        "#  #",
        "# # ",
        "##  ",
        "# # ",
        "#  #",
    ],
    "L": [
        "#   ",
        "#   ",
        "#   ",
        "#   ",
        "####",
    ],
    "M": [
        "#   #",
        "## ##",
        "# # #",
        "#   #",
        "#   #",
    ],
    "N": [
        "#   #",
        "##  #",
        "# # #",
        "#  ##",
        "#   #",
    ],
    "O": [
        " ## ",
        "#  #",
        "#  #",
        "#  #",
        " ## ",
    ],
    "P": [
        "### ",
        "#  #",
        "### ",
        "#   ",
        "#   ",
    ],
    "Q": [
        " ### ",
        "#   #",
        "# # #",
        "#  # ",
        " ## #",
    ],
    "R": [
        "### ",
        "#  #",
        "### ",
        "# # ",
        "#  #",
    ],
    "S": [
        " ###",
        "#   ",
        " ## ",
        "   #",
        "### ",
    ],
    "T": [
        "#####",
        "  #  ",
        "  #  ",
        "  #  ",
        "  #  ",
    ],
    "U": [
        "#  #",
        "#  #",
        "#  #",
        "#  #",
        " ## ",
    ],
    "V": [
        "#   #",
        "#   #",
        "#   #",
        " # # ",
        "  #  ",
    ],
    "W": [
        "#     #",
        "#     #",
        "#  #  #",
        "# # # #",
        " #   # ",
    ],
    "X": [
        "#   #",
        " # # ",
        "  #  ",
        " # # ",
        "#   #",
    ],
    "Y": [
        "#   #",
        "#   #",
        " ### ",
        "  #  ",
        "  #  ",
    ],
    "Z": [
        "#####",
        "   # ",
        "  #  ",
        " #   ",
        "#####",
    ],
    "0": [
        " ### ",
        "#  ##",
        "# # #",
        "##  #",
        " ### ",
    ],
    "1": [
        " # ",
        "## ",
        " # ",
        " # ",
        "###",
    ],
    "2": [
        " ## ",
        "#  #",
        "  # ",
        " #  ",
        "####",
    ],
    "3": [
        " ## ",
        "#  #",
        "  # ",
        "#  #",
        " ## ",
    ],
    "4": [
        "#  #",
        "#  #",
        "####",
        "   #",
        "   #",
    ],
    "5": [
        "####",
        "#   ",
        "### ",
        "   #",
        "### ",
    ],
    "6": [
        " ###",
        "#   ",
        "### ",
        "#  #",
        " ## ",
    ],
    "7": [
        "####",
        "   #",
        "  # ",
        " #  ",
        " #  ",
    ],
    "8": [
        " ## ",
        "#  #",
        " ## ",
        "#  #",
        " ## ",
    ],
    "9": [
        " ## ",
        "#  #",
        " ###",
        "   #",
        " ## ",
    ],
    ".": [
        "  ",
        "  ",
        "  ",
        "  ",
        "# ",
    ],
    ",": [
        "   ",
        "   ",
        "   ",
        " # ",
        "#  ",
    ],
    " ": [
        "   ",
        "   ",
        "   ",
        "   ",
        "   ",
    ],
    "!": [
        "# ",
        "# ",
        "# ",
        "  ",
        "# ",
    ],
    "?": [
        " ##  ",
        "#  # ",
        "  #  ",
        "     ",
        "  #  ",
    ],
    "-": [
        "    ",
        "    ",
        "####",
        "    ",
        "    ",
    ],
    "'": [
        "#",
        "#",
        " ",
        " ",
        " ",
    ],
    '"': [
        "# #",
        "# #",
        "   ",
        "   ",
        "   ",
    ],
    ":": [
        "  ",
        "# ",
        "  ",
        "# ",
        "  ",
    ],
    ";": [
        "   ",
        " # ",
        "   ",
        " # ",
        "#  ",
    ],
}

def wordlen(s: str) -> int:
    """Calculate the length of a word in characters."""
    return sum(len(CHARS[c][0]) for c in s if c in CHARS) + len(s) - 1

def one_line(s: str, width: int) -> tuple[list[str], int]:
    result = ["" for _ in range(5)]
    first = True
    x = 0
    c = 0
    for word in s.split():
        if not first and x + wordlen(" " + word) > width:
            while x < width:
                for j in range(5):
                    result[j] += "_"
                x += 1
            return result, c + 1
        
        for ch in (" " if not first else "") + word:
            first = False
            if ch not in CHARS:
                raise ValueError(f"Character '{ch}' not found in CHARS.")
            c += 1
            x += len(CHARS[ch][0]) + 1
            for j in range(5):
                result[j] += CHARS[ch][j] + " "
    while x < width:
        for j in range(5):
            result[j] += "_"
        x += 1
    return result, len(s)

def multiline(s: str, width: int) -> list[str]:
    lines = []
    while s:
        line, i = one_line(s, width)
        lines.extend(line)
        lines.append("-" * width)
        s = s[i:]
    return lines

def gw_bytes(lines: list[str]) -> bytes:
    """Convert a list of strings to bytes for printing."""
    result = b""
    for line in lines:
        for i in range(0, len(line), 8):
            byte = 0
            for j in range(8):
                if i + j < len(line) and line[i + j] != "#":
                    byte |= (1 << (7 - j))
            result += bytes([byte])
    return result

def print_multiline(s: str, width: int) -> None:
    lines = multiline(s.upper(), width)

    printer_name = win32print.GetDefaultPrinter()
    hprinter = win32print.OpenPrinter(printer_name)
    try:
        hjob = win32print.StartDocPrinter(hprinter, 1, ("Label", None, "RAW"))
        win32print.StartPagePrinter(hprinter)
        win32print.WritePrinter(hprinter, b'N\nQ0,0\nGW' + bytes(f'0,0,{width // 8},{len(lines)},', encoding="ascii") + gw_bytes(lines) + b'\nP1\n')
        win32print.EndPagePrinter(hprinter)
        win32print.EndDocPrinter(hprinter)
    finally:
        win32print.ClosePrinter(hprinter)

print_multiline("The quick brown fox jumps over the lazy dog." * 30, 400)
