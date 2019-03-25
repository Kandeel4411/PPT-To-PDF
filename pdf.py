"""
Short script that uses Libreoffice to convert ppt files in copied path
to pdf

"""
import os
import re
import subprocess as sp
import sys
import pyperclip


def main():
    if len(sys.argv) < 2:
        print("Converts ppt files in copied path to pdf.")
        print("Usage: pdf convertppt")
        sys.exit()

    if sys.argv[1] != "convertppt":
        print("Incorrect usage. Try entering 'pdf'")
        sys.exit()

    if os.path.exists(pyperclip.paste())\
            and not os.path.isfile(pyperclip.paste()):
        os.chdir(pyperclip.paste())
        r = re.compile(r".*\.ppt")

        files = os.listdir(os.getcwd())
        for file in files:
            if re.search(r, file):
                sp.call(['soffice', '--headless', '--convert-to', 'pdf', file])
    else:
        print("Invalid path, Try again with valid path")


if __name__ == "__main__":
    main()
