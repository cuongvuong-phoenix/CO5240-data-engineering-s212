import pandas as pd
from itertools import combinations
from itertools import permutations
import time
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox

import win32gui, win32con
import os
import time
import numbers
# hide = win32gui.GetForegroundWindow()
# win32gui.ShowWindow(hide , win32con.SW_HIDE)

def export_data_source (writer, row, sheet_name):
    export_excel = pd.DataFrame(
        row,
        columns=[
        'LEFT',
        'RIGHT',    
        'CONFIDENCE'
                ]
            )

    export_excel.to_excel(writer, sheet_name, index = True)


def process(source_file, min_s, min_c):
    # start = time.time()

   
    file_name_convert = os.path.splitext(source_file.split("/")[-1])[0]
    print(file_name_convert)
    
    try:
        source = pd.read_excel(source_file, header = 0,  dtype=str)
    except Exception as e:
        return 0
    

    sup = min_s*len(source.values)
    conf = min_c
    print(len(source.values))
    print(sup)
    print(conf)
    lst_init_item = []
    list_init_index = []
    for x in source.values:
        item = x[1].split(',')
        index = x[0]
        list_init_index.append(index)
        try:
            item.remove("")
            item.remove('nan')
        except Exception as e:
            pass
        

        lst_init_item_temp = [lst_init_item_temp_ele[0] for lst_init_item_temp_ele in lst_init_item]

        for i in item:
            if (i not in ('', 'nan')):
                try:
                    location = lst_init_item_temp.index(i)
                    lst_init_item[location][2]+=1
                    try:
                        location_2 = lst_init_item[location][1].index(index)
                    except Exception as e:
                        lst_init_item[location][1].append(index)
                except Exception as e:
                    lst_init_item.append([i, [index], 1])

    k = 2
    lst_proc = list(filter(lambda i: i[1] >=sup,     list([[x[0]], x[2]] for x in lst_init_item.copy())))

    while (True):
        print(k)
        pre_candidate = list(filter(lambda x: len(x[0]) == k-1, lst_proc))
        current_candidate = []
        for x in combinations([pre_candidate_inside[0] for pre_candidate_inside in pre_candidate], k):
            temp = []
            for x_1 in x:
                for x_2 in x_1:
                    try:
                        location = temp.index(x_2)
                    except Exception as e:
                        temp.append(x_2)

            if (len(temp) == k):
            
                
                # count = 0
                # for source_inside in source.values:
                #   item = source_inside[1].split(',')
                #   check = True
                #   for temp_inside in temp:
                #       if (temp_inside not in item):
                #           check = False
                #           break

                #   if (check):
                #       count+=1

                # if (count>=sup):
                #   current_candidate.append([temp, count])


                count = 0
                list_search = []
                lst_init_item_temp = [lst_init_item_temp_ele[0] for lst_init_item_temp_ele in lst_init_item]
                
                for temp_inside in temp:
                    location = lst_init_item_temp.index(temp_inside)
                    for index_check in lst_init_item[location][1]:
                        try:
                            location_index = list_search.index(index_check)
                        except Exception as e:
                            list_search.append(index_check)


                for search_index in list_search:
                    
                    location_search_index = list_init_index.index(search_index)

                    source_target = source.values[location_search_index]


                    item = source_target[1].split(',')
                    check = True
                    for temp_inside in temp:
                        if (temp_inside not in item):
                            check = False
                            break

                    if (check):
                        count+=1

                if (count>=sup):
                    current_candidate.append([temp, count])
                    print([temp, count])
                    
            
        if (current_candidate == []):
            break
        lst_proc+=current_candidate
        k+=1

        
    lst_proc_item = [lst_proc_item_ele[0] for lst_proc_item_ele in lst_proc]

    result = []
 
    for x in lst_proc:
        
        if (len(x[0])>=2):
            temp = []
            for k in range(1, len(x[0])):
                temp+=combinations(x[0], k)
            
            
            for hey_1 in temp:
                temp2 = x[0].copy()
                for hey_2 in hey_1:
                    
                    temp2.remove(hey_2)

                # print("if " + str(hey_1) + " then " + str(temp2))
                
                
                nX = None
                for l in permutations(hey_1, len(hey_1)):
                    l_conv = [l_conv_ele for l_conv_ele in l]
                    

                    try:
                        location = lst_proc_item.index(l_conv)
                        nX = lst_proc[location][1]
                        break
                    except Exception as e:
                        pass

                input_result = [[hey_1_conv_ele for hey_1_conv_ele in hey_1], temp2, x[1]/nX]
                
                if (input_result[2] >= conf):
                    result.append(input_result)
                    print("if " + str(input_result[0]) + " then " + str(input_result[1]) + " with rate " + str(input_result[2]))


    # end = time.time()

    # print("{:.2f}".format(end-start)+"s")


    # for x in result:
    #   print(x)
    try:
        writer = pd.ExcelWriter(os.getcwd()+f'\\result\\{file_name_convert}_result.xlsx', engine='xlsxwriter')
    except Exception as e:
        return 1
    export_data_source(writer, result, 'SOL')
    writer.save()

    return 2


def UploadSource(sourse_file_input, sourse_display):
    
    sourse_file = filedialog.askopenfilename()
    if (sourse_file != ''):
        temp = sourse_file.split('/')
        sourse_display.set(temp[len(temp)-1])
        print('Selected:', sourse_file)
        sourse_file_input[0] = sourse_file

def Calulate(source_file, entry1, entry2, source_display, min_s_display, min_c_display, interval_time):
    print(type(entry1.get()))
    print(type(entry2.get()))

    check = True
    
    if (source_file[0] == None):
        source_display.set("Missing Course config!")
        check = False


    try:
        min_s = float(entry1.get())
    except Exception as e:
        min_s_display.set("check min_s!")
        check = False
    
    try:
        min_c = float(entry2.get())
    except Exception as e:
        min_c_display.set("check min_c!")
        check = False



    print(source_file[0])
    if (check):
        try:
            if (os.path.exists(os.getcwd()+"\\"+'result') != True):
                os.mkdir(os.getcwd()+"\\"+'result')
            start = time.time()
            solution = process(source_file[0], min_s, min_c)
            end = time.time()
            print(solution)
            if (solution == 1):
                tk.messagebox.showinfo("Message", "Close export file!") 
            elif (solution == 0):
                tk.messagebox.showinfo("Message", "Can't open input file") 
            elif (solution == 2):
                file_name_convert = os.path.splitext(source_file[0].split("/")[-1])[0]
    
                print(end-start)
                interval_time.set("{:.2f}".format(end-start)+"s")
                cwd = os.getcwd()
                file = cwd+f"\\result\\{file_name_convert}_result.xlsx"
                os.startfile(file)
                tk.messagebox.showinfo("Message", "DONE") 

        except Exception as e:
            tk.messagebox.showinfo("Message", "Error format input file") 


def main():
    root = tk.Tk()

    source_file = [None]
 

    source_display = tk.StringVar()
    min_s_display = tk.StringVar()
    min_c_display = tk.StringVar()
    interval_time = tk.StringVar()

    frame1 = tk.Canvas(master=root, width=300, height=300)
    frame1.pack()

    label = tk.Label(text="Format file: xls, xlsx")
    label.place(x=20, y=10)

    label = tk.Label(root, textvariable=source_display)
    label.place(x=100, y=32)
    source_display.set("Source file")

    button = tk.Button(root, text='Import', command=lambda:UploadSource(source_file, source_display))
    button.place(x=20, y=30)

    label = tk.Label(root, textvariable=min_s_display)
    label.place(x=20, y=60)
    min_s_display.set('min_s')
 
    entry1 = tk.Entry (root) 
    frame1.create_window(130, 70, window=entry1, width=50)

    label = tk.Label(root, textvariable=min_c_display)
    label.place(x=20, y=90)
    min_c_display.set('min_c')
 
    entry2 = tk.Entry (root) 
    frame1.create_window(130, 100, window=entry2, width=50)

    # label = tk.Label(root, textvariable=reg_display)
    # label.place(x=100, y=62)
    # reg_display.set("Registration file")

    # button = tk.Button(root, text='Import', command=lambda:UploadReg(reg_file, reg_display))
    # button.place(x=20, y=60)

    # label = tk.Label(root, textvariable=teacher_display)
    # label.place(x=100, y=92)
    # teacher_display.set("Teacher file")

    # button = tk.Button(root, text='Import', command=lambda:UploadTeacher(teacher_file, teacher_display))
    # button.place(x=20, y=90)

    button = tk.Button(root, text='Calculate', command=lambda:Calulate(source_file, entry1, entry2, source_display, min_s_display, min_c_display, interval_time))
    button.place(x=120, y=150)


    label = tk.Label(root, textvariable=interval_time)
    label.place(x=130, y=200)
    interval_time.set("")

    root.mainloop()

main()



