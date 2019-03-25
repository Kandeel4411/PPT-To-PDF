"""
Short script that uses Libreoffice to convert ppt files in copied path
to pdf

"""
import os
import subprocess as sp
import sys
import pyperclip


def main():

    # Checks if copied path is a valid path which exists.
    if os.path.exists(pyperclip.paste())\
            and not os.path.isfile(pyperclip.paste()):

        # Changes to copied directory and takes all files into a list.
        os.chdir(pyperclip.paste())
        files = os.listdir(os.getcwd())
        for file in files:

            # Checks if the file extension has ".ppt" in it.
            if os.path.splitext(file)[1].find(".ppt") != -1:
                print(f"{file} Converting..")

                # Calls Libreoffice to convert said file into pdf.
                sp.call(['soffice', '--headless', '--convert-to', 'pdf', file])
                print("Done.")
    else:
        print("Invalid path, Try again with valid path")


if __name__ == "__main__":
    main()
