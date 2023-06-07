import re
import os
import pandas as pd
import argparse

# Initialize parser
parser = argparse.ArgumentParser(prog='gen_excel.py', description='Simple tool for convert raw nmap report.', add_help=True)


# Adding optional argument
parser.add_argument("-p", "--path", default=".", help="Path to nmap report folder")
parser.add_argument("-o", "--output", default="./result.xlsx", help="Path to export excel file")

lstfile = []
lstip = []

# this function by https://regex101.com/codegen?language=python
# last group get from https://stackoverflow.com/q/45002313
regex = r"(\d+)\/(tcp|udp)\s+(open)\s+([^\s]+)(.*)?"

# declare for get info server

lst = []


def getinfo(regex, test_str, ip):
    matches = re.finditer(regex, test_str, re.MULTILINE)
    for match in matches:
        lst.append({
            "IP": ip,
            "Port": match.group(1),
            "Protocol": match.group(2),
            "State": match.group(3),
            "Service": match.group(4),
            "Version": match.group(5)
        })
    return lst


def getfile(path):
    for i in path:
        # this regex by https://stackoverflow.com/a/36760050
        filename_re = r"((25[0-5]|(2[0-4]|1\d|[1-9]|)\d)\.?\b){4}"
        checks = re.finditer(filename_re, i, re.MULTILINE)
        for check in checks:
            if check.group(0) in i:
                lstfile.append(i)
    return lstfile


def getdatafile(lstfile, folder):
    for filename in lstfile:
        ip = filename.replace(".txt", "")
        with open(str(folder) + str(filename)) as f:
            data = f.read()
            getinfo(regex, data, ip)
    return lst


def write_to_excel(lst, sheet_name, output):
    for i in lst:
        # https://www.programmingfunda.com/how-to-convert-dictionary-to-excel-in-python/
        # convert into dataframe
        df = pd.DataFrame(data=lst)
        # convert into Excel
        # df.to_excel("test3.xlsx", index=False)  # simple
        # merge row by https://stackoverflow.com/a/68208815
        df.set_index(df.columns.tolist()).to_excel(str(output), sheet_name=sheet_name)


def main():

    parser.print_help()

    # Read arguments from command line
    args = parser.parse_args()

    # get path argument for get files and ips from folder
    path_to_report_folder = args.path
    path = os.listdir(path=path_to_report_folder)

    # get output argument to export file
    output = args.output

    getdatafile(getfile(path), path_to_report_folder)
    # write sheet Server
    write_to_excel(lst, "Server", output)


if __name__ == "__main__":
    main()
