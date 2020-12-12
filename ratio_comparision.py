import slot_adder

# prints suitable ratios 
def ratio_comparison(list_of_ratios,i,car,ratio_exe):
    if i == 0:
        print("***************************************")
        slot_num = 1
        print("one slot already fulled with content-area")
    elif i == 1:
        slot_num = 2
        ratio_exe(list_of_ratios[i], list_of_ratios[i+1], car,
                  slot_num, list_of_ratios[i+2], list_of_ratios[i+3])
    elif i == 2:
        slot_num = 3
        ratio_exe(list_of_ratios[i], list_of_ratios[i+1], car,
                  slot_num, list_of_ratios[i+2], list_of_ratios[i+3])
    elif i == 3:
        slot_num = 4
        ratio_exe(list_of_ratios[i], list_of_ratios[i+1], car,
                  slot_num, list_of_ratios[i+2], list_of_ratios[i+3])
    elif i == 4:
        slot_num = 5
        ratio_exe(list_of_ratios[i], list_of_ratios[i+1], car,
                  slot_num, list_of_ratios[i+2], list_of_ratios[i+3])
    elif i == 5:
        slot_num = 6
        ratio_exe(list_of_ratios[i], list_of_ratios[i+1], car,
                  slot_num, list_of_ratios[i+2], list_of_ratios[i+3])
