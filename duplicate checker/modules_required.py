import os

try:
        os.system("python get-pip.py")
        os.system("set PATH=$PATH:.")
        os.system("pip install string")

        os.system("pip install regex")
        os.system("pip install openpyxl")

        os.system("pip install pandas")
        os.system("pip install numpy")
        print("Successfully installed all packages")
except:
        print("Error occurred, contact me")
        time.sleep(100)
