import sys
import jpype
import asposecells
jpype.startJVM()
from xlsb2xlsx.xlsb2xlsx import run_xlsb2xlsx, parse_args_fun

if __name__ == "__main__":

    ###### NEED TO CREATE SOME OPTIONS FOR THE RECURSIVE SECOND ARGUMENT #####

    # Run program
    run_xlsb2xlsx(parse_args_fun(sys.argv[1:]))
    print(f"\nDone! Created converted copies of all .xlsb files in {sys.argv[1]} to .xlsx files.\n")
    jpype.shutdownJVM()
