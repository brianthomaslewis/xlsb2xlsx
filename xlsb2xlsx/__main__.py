import sys
import jpype
from xlsb2xlsx.xlsb2xlsx import run_xlsb2xlsx, parse_args_fun

if __name__ == "__main__":
    run_xlsb2xlsx(parse_args_fun(sys.argv[1:]))
    print(
        f"\nDone! Created converted copies of all .xlsb files in {sys.argv[1]} to .xlsx files.\n"
    )
    jpype.shutdownJVM()
