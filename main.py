from sup import result

re = '1'
while re == '1':

    while True:
        chc = input(
            "Choose your branch,\n Enter 1 for CS \t 2 for EE \t 3 for ME \t 4 for CE \t 5 for BT \t 6 for EC")
        if chc == '1':
            branch = 'C S'
            break
        elif chc == '2':
            branch = 'EE'
            break
        elif chc == '3':
            branch = 'ME'
            break
        elif chc == '4':
            branch = 'CE'
            break
        elif chc == '5':
            branch = 'BT'
            break
        elif chc == '6':
            branch = 'EC'
            break
        else:
            print("mze baad me lena sahi enter kro")

    while True:
        try:
            if chc == '4':
                name = int(input("enter your college ID \n"))
            else :
                name = str.lower(input("Enter your full name \n"))
                break
        except ValueError:
            print("mzak baad me,sahi likho")

    s = result(name, branch)
    s.display()
    re = input("sab sahi he na ya dubara chalana chahte ho ise? \t To run again enter 1 else press any key")
    s.clear_screen()

else:
    exit()
