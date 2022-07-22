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
        "msg_directory",
        default=None,
        help="path to directory containing .msg files with attachments to extract",
    )
    parser.add_argument(
        "--processes",
        "-p",
        default=4,
        type=int,
        help="number of processes on the server to employ in running the msg_extractor job",
    )
    input_args = parser.parse_args(args)
    return input_args


def glob_re(pattern, strings):
    return list(filter(re.compile(pattern).match, strings))

def convert_xlsb_to_xlsx(fp):
    df = Workbook(fp)
    df.save(fp[:-5] + '.xlsx')    

def run_xlsb2xlsx(dir_path, recursive=False):
    if recursive==True:
        fps = glob_re(r'.*\.(XLSB|xlsb)', glob.glob(dir_path + '/**', recursive=True))
    else:
        fps = glob_re(r'.*\.(XLSB|xlsb)', os.listdir(dir_path))
    for filepath in fps:
        convert_xlsb_to_xlsx(filepath)

jpype.shutdownJVM()