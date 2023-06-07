import re
import os
import pandas as pd

# requirement pandas, openpyxl


# declare for get files and ips from folder
path = os.listdir(path='.')
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


def getdatafile(lstfile):
    for filename in lstfile:
        ip = filename.replace(".txt", "")
        with open(filename) as f:
            data = f.read()
            getinfo(regex, data, ip)
    return lst


def main():
    getdatafile(getfile(path))
    for i in lst:
        # https://www.programmingfunda.com/how-to-convert-dictionary-to-excel-in-python/
        # convert into dataframe
        df = pd.DataFrame(data=lst)
        # convert into Excel
        # df.to_excel("test3.xlsx", index=False)  # simple
        # merge row by https://stackoverflow.com/a/68208815
        df.set_index(df.columns.tolist()).to_excel('filename.xlsx')


if __name__ == "__main__":
    main()
