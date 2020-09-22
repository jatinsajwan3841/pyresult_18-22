from sup import result

re = '1'
while re == '1':
    s = result()
    s.select(1)
    s.select(2)
    s.select(3)
    s.clear_screen()
    s.display()
    re = input("sab sahi he na ya dubara chalana chahte ho ise? \t To run again enter 1 else press any key")
    s.clear_screen()

else:
    exit()
