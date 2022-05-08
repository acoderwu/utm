import sys

from readExcel import ReadExcel
if __name__ == '__main__':
    if not sys.argv:
        sys.stdout.write("argv error\n")
    else:
        sys.stdout.write("\nbegin read excel and calculate\n")
        for src in sys.argv:
            if "exe" in src:
                continue
            try:
                sys.stdout.write("excel src: " + src + "\n")
                ReadExcel(src).read_calculate()
                sys.stdout.write("complete calculate\n")
            except Exception as e:
                print(e)
