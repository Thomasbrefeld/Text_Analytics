'''
Made by: Thomas Brefeld, Jr.
Email: thomas.brefeld@gmail.com
Date: 08/07/2020
Title: Text Analytics
Discription: Will generate an excel file with grouped topics and subtopics to improve productivity.
'''

import csv
import os
import sys
import nltk
import time
import pandas as pd
import xlsxwriter
from itertools import islice
from itertools import cycle
from nltk.corpus import wordnet
from tkinter import filedialog

#gui imports
from tkinter import Tk
from tkinter import IntVar
from tkinter import Frame
from tkinter import Label 
from tkinter import Entry
from tkinter import Button
from tkinter import ACTIVE
from tkinter import HORIZONTAL
from tkinter import DISABLED
from tkinter import Checkbutton
from tkinter import ttk
from tkinter import END
from tkinter import LabelFrame
from tkinter.font import Font

path = os.getcwd()

wordnet.synsets('load')

root = Tk()
root.resizable(False, False)
root.title('Text Analytics')
notebook = None

TOPIC_FILE = path + '\\src\\topic_categorize.csv'
SUBTOPIC_FILE = path + '\\src\\subtopic_categorize.csv'

log_file = path + 'topic_error_log.txt'

file_path = ''
save_path = ''

global tmp_label
tmp_label = None

autorun_output = IntVar()
exclude_first_line = IntVar()

def output_calc():
    try:
        global run_frame
        global notebook

        global file_entry
        global save_entry

        add_progress(.1)

        Button(run_frame, text="Find Input", state=DISABLED, command=set_file, padx=10).grid(row=0, column=2, padx=5)
        Button(run_frame, text="Find Ouptut", state=DISABLED, command=set_save, padx=5).grid(row=1, column=2, padx=5)
        Button(run_frame, text="Start", state=DISABLED, padx=25).grid(row=2, column=2, padx=25, pady=(25,5))
        
        notebook.tab(1, state='disabled')

        file_entry.configure(state='disabled')
        save_entry.configure(state='disabled')

        try:
            if (file_path[-4:] == '.csv'):
                df = pd.read_csv(file_path, index_col = 0)
            elif (file_path[-5:] == '.xlsx'):
                df = pd.read_excel(file_path, index_col = 0)
            else:
                post_run(2) #bad input file

            if (save_path[-5:] != '.xlsx'):
                post_run(3) #bad output file
            
            if (TOPIC_FILE[-4:] != '.csv' and TOPIC_FILE[-5:] != '.xlsx'):
                post_run(4) #bad topic file

            if (SUBTOPIC_FILE[-4:] != '.csv' and SUBTOPIC_FILE[-5:] != '.xlsx'):
                post_run(5) #bad subtopic file
        except:
            post_run(8)

        workbook = xlsxwriter.Workbook(save_path)
        worksheet = workbook.add_worksheet()

        categories = []
        graph_data_topics = []
        graph_data_values = []

        tmp_file_1 = []
        tmp_file_2 = []

        topic_subtopic_mismatch = False

        if (os.path.exists(log_file)):
            os.remove(log_file)
        
        f = open(log_file, 'w')

        try:
            with open (TOPIC_FILE, 'r', encoding='utf-8-sig') as file1:
                for line1 in csv.reader(file1, delimiter=','):
                    tmp_file_1.append(str(line1[0]).strip().lower())
            with open (SUBTOPIC_FILE, 'r', encoding='utf-8-sig') as file2:
                for line2 in csv.reader(file2, delimiter=','):
                    tmp_file_2.append(str(line2[0]).strip().lower())

            if (len(tmp_file_1) < len(tmp_file_2)):
                f.write('Topic file is ' + str(len(tmp_file_2) - len(tmp_file_1)) + ' element smaller then subtopic file.\n\n')
                topic_subtopic_mismatch = True
                
                for x in range(len(tmp_file_1)):
                    if (tmp_file_1[x] != tmp_file_2[x]):
                        topic_subtopic_mismatch = True
                        f.write('mismatch words line: ' + str(x) + ' => Topic: \'' + tmp_file_1[x] + '\'  Subtopic: \'' + tmp_file_2[x] + '\'\n')
            
            else:
                if (len(tmp_file_1) > len(tmp_file_2)):
                    topic_subtopic_mismatch = True
                    f.write('Subtopic file is ' + str(len(tmp_file_1) - len(tmp_file_2)) + ' element smaller then topic file.\n\n')

                for x in range(len(tmp_file_2)):
                    if (tmp_file_1[x] != tmp_file_2[x]):
                        topic_subtopic_mismatch = True
                        f.write('mismatch words line: ' + str(x) + ' => Topic: \'' + tmp_file_1[x] + '\'  Subtopic: \'' + tmp_file_2[x] + '\'\n')

            f.close()

            if (topic_subtopic_mismatch):
                post_run(9)
            os.remove(log_file)
        except Exception as e:
            print(e)
            post_run(6)

        try:
            with open (TOPIC_FILE, 'r', encoding='utf-8-sig') as file1, open (SUBTOPIC_FILE, 'r', encoding='utf-8-sig') as file2:
                for line1, line2 in zip(csv.reader(file1, delimiter=','), csv.reader(file2, delimiter=',')):
                    categories.append([str(line1[0]).strip(), [line1 for line1 in [str(line1).strip() for line1 in line1] if line1 != '' if line1[0] != '-' if line1[0] != '/'], [line2 for line2 in [str(line2).strip() for line2 in line2[1:]] if line2 != ''], [line1[1:]for line1 in [str(line1).strip() for line1 in line1] if line1 != '' if line1[0] in ['-','/']]])
            categories.append(['Other', [], ['Other'], []])
        except:
            post_run(6) #failed reading src files

        file_dataframe = pd.DataFrame(columns=['topic','subtopic','feedback'])

        try:
            tmp = pd.DataFrame(columns=['phrase']) #skips the first line
            for x, _ in df.iterrows():
                x = str(x)
                if(len(x) >= 3):
                    tmp = tmp.append({'phrase': x}, ignore_index=True)
                add_progress(16.9/df.shape[0])
                root.update()
        except:
            post_run(7) #failed sorting input file 

        for feedback_string_in in tmp.iterrows():
            add_progress(80/(tmp.size - 1))
            root.update()

            feedback_string = feedback_string_in[1].values[0]
            feedback_string = str(feedback_string).strip().lower()
            feedback_string = ' ' + feedback_string

            new_topic = False
            subtopic_found = False

            for topic in categories[:-1]:
                topic_found = False
                if (feedback_string.find(topic[0].lower()) > 0):
                        topic_found = True
                else:
                    for phrase in topic[1]: #include topic words/phrases
                        if (feedback_string.find(phrase.lower()) > 0):
                            topic_found = True
                            break

                for phrase in topic[3]: #exclude topic words/phrases
                    if (feedback_string.find(phrase.lower()) > 0):
                        topic_found = False
                        break

                if topic_found:
                    for subtopic in topic[2]: #loop through topics
                        syn_found = False
                        for syn in wordnet.synsets(subtopic): #checks if synanomys of the subtopic are in the sentence
                            for l in syn.lemmas():
                                if (feedback_string.find(l.name().lower().replace('_', ' ')) > 0):
                                    syn_found = True
                                    break

                        if syn_found or (feedback_string.find(subtopic.lower()) > -1):
                            file_dataframe = file_dataframe.append({'topic' : topic[0], 'subtopic' : subtopic, 'feedback' : feedback_string}, ignore_index=True)
                            subtopic_found = True
                            break
                        
                    if not subtopic_found:
                        file_dataframe = file_dataframe.append({'topic' : topic[0], 'subtopic' : 'Other', 'feedback' : feedback_string}, ignore_index=True)
                    topic_found = False
                    new_topic = True
        
            if not new_topic:
                file_dataframe = file_dataframe.append({'topic' : 'Other', 'subtopic' : 'Other','feedback' : feedback_string}, ignore_index=True)
        add_progress(80/tmp.size)

        file_dataframe = file_dataframe.set_index(['topic', 'subtopic']).sort_index()

        topic_format = workbook.add_format({
            'bold' : True,
            'font_size' : 12,
            'border' : 1,
            'align' : 'center',
            'valign' : 'vcenter',
            'bg_color' : '#EBEBEB',
            'left' : False,
            'right' : False})

        topic_merge_format = workbook.add_format({
            'bold' : True,
            'border' : 1,
            'align' : 'center',
            'valign' : 'top',
            'bg_color' : '#FFFFFF',
            'left' : False,
            'right' : False,
            'bottom': False})

        topic_count_merge_format = workbook.add_format({
            'bold' : False,
            'border' : 1,
            'align' : 'center',
            'valign' : 'top',
            'num_format' : '0',
            'bg_color' : '#FFFFFF',
            'left' : False,
            'right' : False,
            'bottom': False})

        topic_perc_merge_format = workbook.add_format({
            'bold' : False,
            'border' : 1,
            'align' : 'center',
            'valign' : 'top',
            'num_format' : '0.00%',
            'bg_color' : '#FFFFFF',
            'left' : False,
            'right' : False,
            'bottom': False})

        subtopic_merge_format1 = workbook.add_format({
            'bold' : False,
            'border' : 1,
            'align' : 'center',
            'valign' : 'top',
            'bg_color' : '#EEEEEE',
            'left' : False,
            'right' : False,
            'bottom': False})

        subtopic_merge_format2 = workbook.add_format({
            'bold' : False,
            'border' : 1,
            'align' : 'center',
            'valign' : 'top',
            'bg_color' : '#FFFFFF',
            'left' : False,
            'right' : False,
            'bottom': False})

        subtopic_count_merge_format1 = workbook.add_format({
            'bold' : False,
            'border' : 1,
            'align' : 'center',
            'valign' : 'top',
            'num_format' : '0',
            'bg_color' : '#EEEEEE',
            'left' : False,
            'right' : False,
            'bottom': False})

        subtopic_count_merge_format2 = workbook.add_format({
            'bold' : False,
            'border' : 1,
            'align' : 'center',
            'valign' : 'top',
            'num_format' : '0',
            'bg_color' : '#FFFFFF',
            'left' : False,
            'right' : False,
            'bottom': False})

        subtopic_perc_merge_format1 = workbook.add_format({
            'bold' : False,
            'border' : 1,
            'align' : 'center',
            'valign' : 'top',
            'num_format' : '0.00%',
            'bg_color' : '#EEEEEE',
            'left' : False,
            'right' : False,
            'bottom': False})

        subtopic_perc_merge_format2 = workbook.add_format({
            'bold' : False,
            'border' : 1,
            'align' : 'center',
            'valign' : 'top',
            'num_format' : '0.00%',
            'bg_color' : '#FFFFFF',
            'left' : False,
            'right' : False,
            'bottom': False})

        feedback_sentence_format1 = workbook.add_format({
            'bold' : False,
            #'border' : 0,
            #'top' : 4,
            'align' : 'left',
            'valign' : 'top',
            'bg_color' : '#EEEEEE'})
        feedback_sentence_format1.set_text_wrap()

        feedback_sentence_format2 = workbook.add_format({
            'bold' : False,
            #'border' : 0,
            #'top' : 4,
            'align' : 'left',
            'valign' : 'top',
            'bg_color' : '#FFFFFF'})
        feedback_sentence_format2.set_text_wrap()

        subtopic_perc_merge_format = cycle([subtopic_perc_merge_format1, subtopic_perc_merge_format2])
        subtopic_count_merge_format = cycle([subtopic_count_merge_format1, subtopic_count_merge_format2])
        subtopic_merge_format = cycle([subtopic_merge_format1, subtopic_merge_format2])
        feedback_sentence_format = cycle([feedback_sentence_format1, feedback_sentence_format2])

        topic_loc = 2
        idx = 0

        worksheet.write('A1', 'Topic', topic_format)
        worksheet.write('B1', 'Count', topic_format)
        worksheet.write('C1', 'Percent', topic_format)
        worksheet.write('D1', 'Sub Topic', topic_format)
        worksheet.write('E1', 'Count', topic_format)
        worksheet.write('F1', 'Percent', topic_format)
        worksheet.write('G1', 'Sentences', topic_format)

        other_count = file_dataframe.loc['Other'].count()
        total_line = 0
        for topic in categories:
            try:
                topic_count = file_dataframe.loc[topic[0]].count()
                subtopic_count = file_dataframe.loc[topic[0]].groupby('subtopic').count()

                graph_data_topics.append(topic[0])
                graph_data_values.append(topic_count.values[0])

                unique_list1 = set(topic[2])
                unique_list2 = set(subtopic_count.index.values)
                subtopic_set = topic[2] + list(unique_list2 - unique_list1)

                if (topic_count.values[0] == 0):
                    worksheet.write('A{}'.format(topic_loc), topic[0], topic_merge_format)
                    worksheet.write('B{}'.format(topic_loc), topic_count.values[0], topic_count_merge_format)
                    worksheet.write('C{}'.format(topic_loc), topic_count.values[0] / file_dataframe.shape[0], topic_perc_merge_format)
                else:
                    worksheet.merge_range('A{}:A{}'.format(topic_loc, topic_loc + topic_count.values[0] + len(subtopic_count.values)), topic[0], topic_merge_format)
                    worksheet.merge_range('B{}:B{}'.format(topic_loc, topic_loc + topic_count.values[0] + len(subtopic_count.values)), topic_count.values[0], topic_count_merge_format)
                    worksheet.merge_range('C{}:C{}'.format(topic_loc, topic_loc + topic_count.values[0] + len(subtopic_count.values)), topic_count.values[0] / file_dataframe.shape[0], topic_perc_merge_format)

                subtopic_loc = 0
                
                try:
                    for x in range(topic_loc, topic_loc + topic_count.values[0] + len(subtopic_count.values)):
                        worksheet.set_row(x, None, None, {'level' : 1, 'hidden' : False})
                except:
                    pass

                for subtopic in subtopic_set:
                    try:
                        sel_subtopic_perc_merge_format = next(subtopic_perc_merge_format)
                        sel_subtopic_count_merge_format = next(subtopic_count_merge_format)
                        sel_subtopic_merge_format = next(subtopic_merge_format)
                        sel_feedback_sentence_format = next(feedback_sentence_format)

                        worksheet.write('G{}'.format(topic_loc + subtopic_loc), None, sel_feedback_sentence_format)
                        worksheet.write('G{}'.format(topic_loc + subtopic_loc + subtopic_count.loc[subtopic].values[0] + 1), None, sel_feedback_sentence_format)

                        if (subtopic != 'Other'):
                            if (subtopic_count.loc[subtopic].values[0] == 0):
                                worksheet.write('D{}'.format(topic_loc + subtopic_loc), subtopic, sel_subtopic_merge_format)
                                worksheet.write('E{}'.format(topic_loc + subtopic_loc), subtopic_count.loc[subtopic].values[0], sel_subtopic_count_merge_format)
                                worksheet.write('F{}'.format(topic_loc + subtopic_loc), subtopic_count.loc[subtopic].values[0] / file_dataframe.shape[0], sel_subtopic_perc_merge_format)
                            else:
                                worksheet.merge_range('D{}:D{}'.format(topic_loc + subtopic_loc, topic_loc + subtopic_loc + subtopic_count.loc[subtopic].values[0]), subtopic, sel_subtopic_merge_format)
                                worksheet.merge_range('E{}:E{}'.format(topic_loc + subtopic_loc, topic_loc + subtopic_loc + subtopic_count.loc[subtopic].values[0]), subtopic_count.loc[subtopic].values[0], sel_subtopic_count_merge_format)
                                worksheet.merge_range('F{}:F{}'.format(topic_loc + subtopic_loc, topic_loc + subtopic_loc + subtopic_count.loc[subtopic].values[0]), subtopic_count.loc[subtopic].values[0] / file_dataframe.shape[0], sel_subtopic_perc_merge_format)
                        else:
                            if (subtopic_count.loc[subtopic].values[0] == 0):
                                worksheet.write('D{}'.format(topic_loc + subtopic_loc), subtopic, sel_subtopic_merge_format)
                                worksheet.write('E{}'.format(topic_loc + subtopic_loc), subtopic_count.loc[subtopic].values[0], sel_subtopic_count_merge_format)
                                worksheet.write('F{}'.format(topic_loc + subtopic_loc), subtopic_count.loc[subtopic].values[0] / file_dataframe.shape[0], sel_subtopic_perc_merge_format)
                            else:
                                worksheet.merge_range('D{}:D{}'.format(topic_loc + subtopic_loc, topic_loc + subtopic_loc + subtopic_count.loc[subtopic].values[0] + 1), subtopic, sel_subtopic_merge_format)
                                worksheet.merge_range('E{}:E{}'.format(topic_loc + subtopic_loc, topic_loc + subtopic_loc + subtopic_count.loc[subtopic].values[0] + 1), subtopic_count.loc[subtopic].values[0], sel_subtopic_count_merge_format)
                                worksheet.merge_range('F{}:F{}'.format(topic_loc + subtopic_loc, topic_loc + subtopic_loc + subtopic_count.loc[subtopic].values[0] + 1), subtopic_count.loc[subtopic].values[0] / file_dataframe.shape[0], sel_subtopic_perc_merge_format)

                        feedback_loc = 0

                        try:
                            for x in range(topic_loc + subtopic_loc, topic_loc + subtopic_loc + subtopic_count.loc[subtopic].values[0]):
                                worksheet.set_row(x, None, None, {'level' : 2, 'hidden' : True})
                        except:
                            pass

                        for sentence in file_dataframe.loc[topic[0]].loc[subtopic].values:
                            if(len(sentence[0]) != 1):
                                sentence = sentence[0]
                            worksheet.write('G{}'.format(topic_loc + subtopic_loc + feedback_loc + 1), '•' +  sentence, sel_feedback_sentence_format)
                            feedback_loc += 1

                        idx += 1
                        subtopic_loc += subtopic_count.loc[subtopic].values[0] + 1
                    except:
                        sel_subtopic_perc_merge_format = next(subtopic_perc_merge_format)
                        sel_subtopic_count_merge_format = next(subtopic_count_merge_format)
                        sel_subtopic_merge_format = next(subtopic_merge_format)
                        sel_feedback_sentence_format = next(feedback_sentence_format)
                topic_loc += topic_count.values[0] + len(subtopic_count.values) + 1
                total_line = topic_loc
            except:
                pass
            add_progress(2/len(categories))
        
        worksheet.set_column('A:A', 15, None)
        worksheet.set_column('B:B', 6, None)
        worksheet.set_column('C:C', 8, None)
        worksheet.set_column('D:D', 15, None)
        worksheet.set_column('E:E', 6, None)
        worksheet.set_column('F:F', 8, None)
        worksheet.set_column('G:G', 140, None, {'warp' : True})

        worksheet.conditional_format('F2:F{}'.format(total_line - other_count.values[0] - 3), {
            'type' : '2_color_scale',
            'min_color' : "#CCFFCC",
            'max_color' : "#FF8080"})
        worksheet.conditional_format('C2:C{}'.format(total_line - other_count.values[0] - 3), {
            'type' : '2_color_scale',
            'min_color' : "#CCFFCC",
            'max_color' : "#FF8080"})

        worksheet.set_zoom(125)

        workbook.close()
        set_progress(100)
        post_run(0)
    except Exception as e:
        post_run(e) #any other error

def set_file():
    global file_path
    global file_entry
    try:
        file_path = filedialog.askopenfilename(initialdir="input", title="Select Input File")
    except:
        file_path = ''
    file_entry.configure(state='normal')
    file_entry.delete(0, END)
    file_entry.insert(0, file_path)
    file_entry.configure(state='disabled')
    is_ready()

def set_save():
    global save_path
    global save_entry

    global tmp_label
    global run_frame

    try:
        save_path = filedialog.asksaveasfile(initialdir="output", title="Select Output File", filetypes = [('Excel File', '*.xlsx')], defaultextension = [('Excel File', '*.xlsx')])
        if (tmp_label != None):
            tmp_label.pack_forget()
            tmp_label.destroy()
            tmp_label = None
            run_frame.update_idletasks()
    except:
        large_font = Font(family="Times", size=10)
        tmp_label = Label(run_frame, text='Insufficient Permissions to access that file: File might be open by another program', padx=113, bg='#f44336', fg='white', font=large_font)
        tmp_label.grid(row=1, column=1)
        save_path = None
        
    if (save_path == None):
        save_path = ''
    else:
        save_path = save_path.name
    save_entry.configure(state='normal')
    save_entry.delete(0, END)
    save_entry.insert(0, save_path)
    save_entry.configure(state='disabled')
    is_ready()

def set_topic_file():
    global TOPIC_FILE
    global topic_entry
    try:
        TOPIC_FILE = filedialog.askopenfilename(initialdir="src/", title="Select Topic File", filetypes = [('Comma-separated values file', '*.csv')], defaultextension = [('Comma-separated values file', '*.csv')])
    except:
        TOPIC_FILE = 'src/topic_categorize.csv'
    if (TOPIC_FILE == ''):
        TOPIC_FILE = 'src/topic_categorize.csv'

    topic_entry.configure(state='normal')
    topic_entry.delete(0, END)
    topic_entry.insert(0, TOPIC_FILE)
    topic_entry.configure(state='disabled')

def set_subtopic_file():
    global SUBTOPIC_FILE
    global subtopic_entry
    try:
        SUBTOPIC_FILE = filedialog.askopenfilename(initialdir="src/", title="Select Subtopic File", filetypes = [('Comma-separated values file', '*.csv')], defaultextension = [('Comma-separated values file', '*.csv')])
    except:
        SUBTOPIC_FILE = 'src/subtopic_categorize.csv'
    if (SUBTOPIC_FILE == ''):
        SUBTOPIC_FILE = 'src/subtopic_categorize.csv'

    subtopic_entry.configure(state='normal')
    subtopic_entry.delete(0, END)
    subtopic_entry.insert(0, SUBTOPIC_FILE)
    subtopic_entry.configure(state='disabled')
    is_ready()

def is_ready():
    global run_frame
    if(file_path == '' or save_path == ''):
        Button(run_frame, text="Start", state=DISABLED, padx=25).grid(row=2, column=2, padx=25, pady=(25,5))
    else:
        Button(run_frame, text="Start", state=ACTIVE, padx=25, command=output_calc).grid(row=2, column=2, padx=25, pady=(25,5))

def add_progress(inc):
    global progress_bar
    global run_frame
    progress_bar['value'] += inc
    run_frame.update_idletasks()

def set_progress(inc):
    global progress_bar
    global run_frame
    progress_bar['value'] = inc
    run_frame.update_idletasks()

def exit_prog():
    sys.exit()

def reset():
    global file_path
    global save_path
    
    file_path = ''
    save_path = ''

def run_frame_fun(run_frame):
    global file_entry
    global save_entry

    global file_path
    global save_path

    global progress_bar
    global autorun_output

    Label(run_frame, text='Input File: ', padx=10, pady=10).grid(row=0, column=0)
    Label(run_frame, text='Output File: ', padx=10, pady=10).grid(row=1, column=0)

    file_entry = Entry(run_frame, width=110)
    file_entry.grid(row=0, column=1)
    file_entry.insert(0, file_path)
    file_entry.configure(state='disabled')

    save_entry = Entry(run_frame, width=110)
    save_entry.grid(row=1, column=1)
    save_entry.insert(0, save_path)
    save_entry.configure(state='disabled')

    Button(run_frame, text="Find Input", state=ACTIVE, command=set_file, padx=10).grid(row=0, column=2, padx=5)
    Button(run_frame, text="Find Ouptut", state=ACTIVE, command=set_save, padx=5).grid(row=1, column=2, padx=5)

    progress_bar = ttk.Progressbar(run_frame, orient=HORIZONTAL, length=666, mode='determinate')
    progress_bar.grid(row=2, column=1, pady=(25,5))

    is_ready()

    c = Checkbutton(run_frame, text='Auto-Open', variable=autorun_output)
    c.grid(row=2, column=0, pady=(25,5))

    return run_frame

def settings_frame_fun(settings_frame):
    global topic_entry
    global subtopic_entry

    def open_topic_file():
        os.system("start " + '\"\" \"{}\"'.format(TOPIC_FILE))
    
    def open_subtopic_file():
        os.system("start " + '\"\" \"{}\"'.format(SUBTOPIC_FILE))

    def reset_settings():
        global TOPIC_FILE
        global SUBTOPIC_FILE

        global file_path
        global save_path

        global save_entry
        global file_entry

        global root
        global notebook

        TOPIC_FILE = 'src/topic_categorize.csv'
        SUBTOPIC_FILE = 'src/subtopic_categorize.csv'

        file_path = ''
        save_path = ''

        file_entry.configure(state='normal')
        file_entry.delete(0, END)
        file_entry.insert(0, file_path)
        file_entry.configure(state='readonly')

        save_entry.configure(state='normal')
        save_entry.delete(0, END)
        save_entry.insert(0, save_path)
        save_entry.configure(state='readonly')

        is_ready()

        topic_entry.configure(state='normal')
        topic_entry.delete(0, END)
        topic_entry.insert(0, TOPIC_FILE)
        topic_entry.configure(state='readonly')

        subtopic_entry.configure(state='normal')
        subtopic_entry.delete(0, END)
        subtopic_entry.insert(0, SUBTOPIC_FILE)
        subtopic_entry.configure(state='readonly')

        root.destroy()

        root = Tk()
        root.resizable(False, False)
        root.title('Text Analytics')
        notebook = None

        main_window()

    Label(settings_frame, text='Topic Source File: ', padx=10, pady=10).grid(row=0, column=0)
    Label(settings_frame, text='Subtopic Source File: ', padx=10, pady=10).grid(row=1, column=0)
    
    topic_entry = Entry(settings_frame, width=82)
    topic_entry.grid(row=0, column=1)
    topic_entry.insert(0, TOPIC_FILE)
    topic_entry.configure(state='disabled')

    subtopic_entry = Entry(settings_frame, width=82)
    subtopic_entry.grid(row=1, column=1)
    subtopic_entry.insert(0, SUBTOPIC_FILE)
    subtopic_entry.configure(state='disabled')

    Button(settings_frame, text="Find Topic File", state=ACTIVE, command=set_topic_file, padx=14).grid(row=0, column=2, padx=5)
    Button(settings_frame, text="Find Subtopic File", state=ACTIVE, command=set_subtopic_file, padx=5).grid(row=1, column=2, padx=5)

    Button(settings_frame, text="Open Topic File", state=ACTIVE, command=open_topic_file, padx=13).grid(row=0, column=3, padx=5)
    Button(settings_frame, text="Open Subtopic File", state=ACTIVE, command=open_subtopic_file, padx=5).grid(row=1, column=3, padx=5)

    Button(settings_frame, text="Help", command=help, padx= 40, pady=2).grid(row=2, column=2, padx=2, pady=(15,5))
    Button(settings_frame, text="Restart Program", command=reset_settings, padx= 14, pady=2).grid(row=2, column=3, padx=2, pady=(15,5))

    return settings_frame

def help():
    os.system("start \"\" \"" + path + "\\README.txt\"")

def post_run(status_num):
    global root
    notebook.destroy()
    large_font = Font(family="Times", size=14)
    med_font = Font(family="Times", size=12)

    post_run_frame = Frame(root)
    post_run_frame.pack()

    frame = LabelFrame(post_run_frame, padx=5, pady=5, borderwidth=0, highlightthickness=0)
    frame.grid(row=1, column=0, padx=10, pady=(0,10))

    frame2 = LabelFrame(post_run_frame, padx=5, pady=5)
    frame2.grid(row=2, column=0, padx=20, pady=(0,10))

    def run_again():
        reset()
        post_run_frame.destroy()
        main_window()
    
    def show_xlsx(event=None):
        os.system("start \"\" \"" + save_path + "\"")

    def show_error():
        os.system("start \"\" \"" + log_file + "\"")

    if (status_num == 0):
        if autorun_output.get():
            show_xlsx()
        Label(post_run_frame, text='Program Finished Successfully!', font=large_font).grid(row=0, column=0, padx=10, pady=10)

        Label(frame2, text='Output Location:', font=med_font).grid(row=1, column=0, padx=10, pady=(3,0))
        save_entry = Entry(frame2, width=90)
        save_entry.grid(row=2, column=0)
        save_entry.insert(0, save_path)
        save_entry.configure(state='readonly')

        Button(frame, text="Exit", command = exit_prog, padx=38, pady=2).grid(row=1, column=0, padx=25)
        Button(frame, text="Run Again", command = run_again, padx=26, pady=2).grid(row=1, column=1, padx=25)
        Button(frame, text="Open Output", command = show_xlsx, padx=20, pady=2).grid(row=1, column=2, padx=25)
    else:
        Button(frame, text="Exit", command = exit_prog, padx= 38, pady=2).grid(row=1, column=0, padx=2)
        Button(frame, text="Run Again", command = run_again, padx= 26, pady=2).grid(row=1, column=1, padx=2)
        Button(frame, text="Help", command=help, padx= 20, pady=2).grid(row=1, column=2, padx=2)
        if (status_num == 2):
            Label(post_run_frame, text='ERROR 2: Bad input file', font=large_font).grid(row=0, column=0, padx=10, pady=10)
        elif (status_num == 3):
            Label(post_run_frame, text='ERROR 3: Bad output file', font=large_font).grid(row=0, column=0, padx=10, pady=10)
        elif (status_num == 4):
            Label(post_run_frame, text='ERROR 4: Bad topic file', font=large_font).grid(row=0, column=0, padx=10, pady=10)
        elif (status_num == 5):
            Label(post_run_frame, text='ERROR 5: Bad subtopic file', font=large_font).grid(row=0, column=0, padx=10, pady=10)
        elif (status_num == 6):
            Label(post_run_frame, text='ERROR 6: Failed reading topic or subtopic', font=large_font).grid(row=0, column=0, padx=10, pady=10)
        elif (status_num == 7):
            Label(post_run_frame, text='ERROR 7: Failed to sort input file', font=large_font).grid(row=0, column=0, padx=10, pady=10)
        elif (status_num == 8):
            Label(post_run_frame, text='ERROR 8: Unsupported file type', font=large_font).grid(row=0, column=0, padx=10, pady=10)
        elif (status_num == 9):
            Label(post_run_frame, text='ERROR 9: Mismatch Topic and Subtopic files check log file', font=large_font).grid(row=0, column=0, padx=10, pady=10)
            show_error()
        else:
            Label(post_run_frame, text='ERROR 50: Uncaught Exception', font=large_font).grid(row=0, column=0, padx=10, pady=10)

    post_run_frame.bind('<Return>', show_xlsx)
    root.protocol("WM_DELETE_WINDOW", exit_prog)
    post_run_frame.mainloop()

def main_window():
    global run_frame
    global notebook

    notebook = ttk.Notebook(root)
    notebook.pack(pady=(5,0))

    run_frame = Frame(notebook)
    settings_frame = Frame(notebook)

    root.focus_set()

    notebook.add(run_frame_fun(run_frame), text='Run')
    notebook.add(settings_frame_fun(settings_frame), text='Settings')

    root.protocol("WM_DELETE_WINDOW", exit_prog)
    root.mainloop()


if __name__ == "__main__":
    main_window()