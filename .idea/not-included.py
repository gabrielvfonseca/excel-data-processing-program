def uploadLocalFile(path, local):
        while True:
            if (str(path)[0:2] == "ls"):
                fileArr = os.listdir(".")
                if (str(path) == "ls -x"):
                    # Only display .xlsx files
                    for item in fileArr:
                        if (str(item)[-4:] == "xlsx"):
                            print(f'  {item}')
                else:
                    # Display all files and directorys
                    for element in fileArr:
                        print(f'  {element}')
            else:
                if (os.path.isfile(path)):
                    # os.rename(path, local)
                    break

            print(f"{colors.WARNING}File does not exist!{colors.DEFAULT}")
            time.sleep(2)
            path = askInput()