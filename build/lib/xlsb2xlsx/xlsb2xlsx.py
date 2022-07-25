import os
import re
import glob
import argparse
import jpype
import asposecells
jpype.startJVM()
from asposecells.api import Workbook

def parse_args_fun(args):
    """Argument parser pulled out for modularity"""
    # Parser for command line argument parsing
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "xlsb_directory",
        default=None,
        help="path to directory containing .xlsb files to convert to .xlsx",
    )
    parser.add_argument(
        "--recursive",
        "-r",
        default=False,
        type=str,
        help="Boolean flag to determine if the function should run recursively through child directories",
    )
    input_args = parser.parse_args(args)
    return input_args


def glob_re(pattern, strings):
    return list(filter(re.compile(pattern).match, strings))

def convert_xlsb_to_xlsx(fp):
    df = Workbook(fp)
    df.save(fp[:-5] + '.xlsx')    

def run_xlsb2xlsx(dir_path, recursive=False):
    dir_path = str(dir_path)
    print(dir_path)
    if recursive==True:
        fps = glob_re(r'.*\.(XLSB|xlsb)', glob.glob(dir_path + '/**', recursive=True))
        print(fps)
    else:
        fps = glob_re(r'.*\.(XLSB|xlsb)', glob.glob(dir_path + '/**'))
        print(fps)
    for filepath in fps:
        print(filepath)
        convert_xlsb_to_xlsx(filepath)

jpype.shutdownJVM()