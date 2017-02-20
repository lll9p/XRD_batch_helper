#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import configparser
import logging
import os
import pathlib
import subprocess
import tkinter as tk
import tkinter.filedialog as filedialog
from tkinter import font, messagebox, ttk


class Config(configparser.ConfigParser):

    def __init__(self, file, default):
        converters = {'path': pathlib.Path}
        super().__init__(converters=converters)
        self.file = str(file)
        self.default = default
        self.get_config()

    def get_config(self):
        self.read(self.file)
        if not self.sections():
            self.reset_config()
            self.update_config()

    def reset_config(self):
        self.read_dict(self.default)

    def update_config(self):
        with open(self.file, 'w') as file:
            self.write(file)


class LogHandler(logging.Logger):

    def __init__(self, name, level, file):
        super().__init__(name)
        self.level = level
        self.setLevel(self.level)
        self.file = file
        self.set_handler()

    def set_handler(self):
        file_handler = logging.FileHandler(self.file)
        file_handler.setLevel(self.level)
        formatter = logging.Formatter(
            '%(asctime)s - %(name)s - %(levelname)s - %(message)s')
        file_handler.setFormatter(formatter)
        self.file_handler = file_handler
        self.addHandler(file_handler)


class ProcesserFrame(ttk.Frame):

    def __init__(self, master, control_frame, data, logger, **kwarg):
        super().__init__(master)
        self.data = data
        self.tasks = self.data.get('tasks')
        self.app = self.data.get('app')
        self.logger = logger
        self.control_frame = control_frame
        self.INP_names = list(self.data.get('INPs').keys())
        self.INP_filename = tk.StringVar()
        self.INP_filename.set(self.INP_names[0] if self.INP_names else '')
        self.output_filename = tk.StringVar()
        self.task = {'program': self.data.get('TC_location', ''),
                     'inp': self.INP_filename,  # INP_filename could be void string
                     'patterns': list(),
                     'processer': self,
                     }  # TODO: add a default TC to it
        self.tasks[hash(self)] = self.task
        self.create_widgets()

    def destroy_processer(self):
        self.control_frame.destroy_processer(self)

    def create_widgets(self):
        self.separator = ttk.Separator(master=self, orient=tk.HORIZONTAL)

        self.destry_button = ttk.Button(
            self, text='X', command=self.destroy_processer, width=2)
        self.destry_button.grid(row=1, column=0)

        self.INP_combobox = ttk.Combobox(
            self, state='readonly',
            textvariable=self.INP_filename,
            values=self.INP_names)
        self.INP_combobox.grid(row=1, column=1)
        self.INP_combobox.bind('<<ComboboxSelected>>', self.select_INP)

        self.INP_choose_button = ttk.Button(
            self, text='I', command=self.choose_INP, width=2)
        self.INP_choose_button.grid(row=1, column=2)

        self.pattens_choose_button = ttk.Button(
            self, text='F', command=self.choose_pattens, width=2)
        self.pattens_choose_button.grid(row=1, column=3)

        self.pattens_dir_choose_button = ttk.Button(
            self, text='D', command=self.choose_pattens_dir, width=2)
        self.pattens_dir_choose_button.grid(row=1, column=4)

        self.output_choose_button = ttk.Button(
            self, text='O', command=self.choose_output, width=2)
        self.output_choose_button.grid(row=1, column=5)

        self.process_button = ttk.Button(
            self, text='√', command=self.process, width=2)
        self.process_button.grid(row=1, column=6)

    def choose_INP(self, title='Choose a TOPAS INP file'):
        INP_file = filedialog.askopenfilename(
            defaultextension='.inp', title=title)
        if not os.path.isfile(INP_file):
            return
        INP_file_short_name = os.path.split(os.path.abspath(INP_file))[1]
        if INP_file_short_name in self.INP_names:
            self.logger.info(
                f'User choose a INP ({INP_file}) which duplicates with INP dir\'s.')
        conflict_filenames = list(
            filter(lambda x: x.startswith(INP_file_short_name), self.INP_names))
        if len(conflict_filenames) != 0:
            for i in range(len(conflict_filenames) + 1):
                tmp = INP_file_short_name + f'({i})'
                if tmp not in self.INP_names:
                    INP_file_short_name = tmp
                    break
        self.data['INPs'][INP_file_short_name] = INP_file
        self.INP_filename.set(INP_file_short_name)
        self.task['inp'] = self.INP_filename.get()
        self.update_INP_combobox_values()

    def update_INP_combobox_values(self):
        self.control_frame.update_processers_INP_combobox()

    def select_INP(self, vent):
        '''
        Change task filename after select an INP
        '''
        self.task['inp'] = self.INP_filename.get()

    def choose_pattens(self):
        patterns = filedialog.askopenfilenames(defaultextension='.raw', filetypes=[(
            'Bruker raw file patterns', '.raw')], title='Choose .raw file(s) of patterns')
        self.task['patterns'].extend(set(patterns))

    def choose_pattens_dir(self):
        patterns_dir = filedialog.askdirectory(
            title='Choose a dir to read all patterns it contains')
        if os.path.isdir(patterns_dir):
            patterns = list(map(lambda f: os.path.join(patterns_dir, f), filter(
                lambda f: f.endswith('.raw'), os.listdir(patterns_dir))))
            self.task['patterns'].extend(patterns)
            self.task['patterns'] = list(set(self.task['patterns']))
        else:
            messagebox.showinfo(
                title='Alert', message='You choose an invalid directory!')

    def choose_output(self):
        output = filedialog.asksaveasfilename(defaultextension='.txt', filetypes=[(
            'text report', '.txt'), ('microsoft excel', '.xlsx')], title='Choose file to output results')
        self.output = output

    def process(self):
        self.app.process(hash(self))
        self.logger.info(
            f'Using {self.data["TC_location"]} process {", ".join(self.task["patterns"])} with {self.INP_filename.get()}')


class ControlFrame(ttk.Frame):

    def __init__(self, master, data, logger, **kwarg):
        super().__init__(master)
        self.data = data
        self.tasks = self.data.get('tasks')
        self.app = self.data.get('app')
        self.logger = logger
        self.create_widgets()

    def create_processer(self):
        processer = ProcesserFrame(
            self.master, control_frame=self, data=self.data, logger=self.logger)
        processer.grid()

    def destroy_processer(self, processer):
        try:
            self.tasks.pop(hash(processer))
            processer.destroy()
        except IndexError:
            self.logger.info('processer to be remove but not exists.')

    def create_widgets(self):
        self.create_processer_button = ttk.Button(
            self, text='+', command=self.create_processer, width=2)
        self.create_processer_button.grid(row=0, column=0)
        self.TC_button = ttk.Button(
            self, text='TC', command=self.select_TC, width=2)
        self.TC_button.grid(row=0, column=1)

        self.process_all_button = ttk.Button(
            self, text='√', command=self.process_all, width=2)
        self.process_all_button.grid(row=0, column=2)

        self.about_button = ttk.Button(
            self, text='About', command=self.about)
        self.about_button.grid(row=0, column=3)

    def select_TC(self):
        self.data['TC_location'] = pathlib.Path(filedialog.askopenfilename(
            defaultextension='.exe', filetypes=[('Executable', '.exe')], title='Find the TOPAS tc.exe'))

    def process_all(self):
        for task in self.tasks.keys():
            self.app.process(task)

    def update_processers_INP_combobox(self):
        # Update the INP_name list
        INP_names = list(self.data.get('INPs').keys())
        for processer in self.processers:
            if isinstance(processer, ProcesserFrame):
                processer.INP_names = INP_names
                processer.INP_combobox.configure(values=INP_names)

    def about(self):
        string = '''Auther : Lao Lilin
Email  : LAOLILIN1@crcement.com
Version: 0.01
Date   : 2017-02-07 10:17
'''
        messagebox.showinfo(title='About', message=string)


class AppGUI(ttk.tkinter.Tk):

    def __init__(self, data, logger, title, theme):
        self.title = title
        super().__init__(className=self.title)
        self.wm_title(self.title)
        self.protocol('WM_DELETE_WINDOW', self.on_exit)
        self.resizable(False, False)
        self.minsize(width=font.Font().measure(title + '-' * 6), height=0)
        self.data = data
        self.tasks = self.data.get('tasks')
        self.app = self.data.get('app')
        self.logger = logger
        self.theme = theme
        self.set_style()
        self.control_frame = ControlFrame(
            master=self, data=self.data, logger=self.logger)
        self.control_frame.grid(sticky=tk.N)
        # Create one processerframe
        self.control_frame.create_processer()

    def on_exit(self):
        self.destroy()

    def set_style(self):
        self.style = ttk.Style()
        self.style.theme_use(self.theme)
        self.style.configure('red.TButton',
                             foreground='white',
                             background='red',
                             bordercolor='red',
                             )
        self.style.map('red.TButton',
                       foreground=[('active', 'white'),
                                   ('pressed', 'white')
                                   ],
                       background=[('active', 'red'),
                                   ('pressed', 'red')
                                   ],
                       highlightcolor=[('focus', 'red'),
                                       ('!focus', 'red')
                                       ],
                       )
        self.style.configure('green.TButton',
                             foreground='black',
                             background='green',
                             bordercolor='green',
                             )
        self.style.map('green.TButton',
                       foreground=[('active', 'black'),
                                   ('pressed', 'black')
                                   ],
                       background=[('active', 'green'),
                                   ('pressed', 'green')
                                   ],
                       highlightcolor=[('focus', 'green'),
                                       ('!focus', 'green')
                                       ],
                       )


class App:

    def __init__(self, name, path, logger, config):
        self.path = path
        self.logger = logger
        self.name = name
        self.config = config
        INP_path = self.config['PATH'].getpath('INP_path')
        self.TC_location = self.config['PATH'].getpath('TC_location')
        self.TOPAS_location = self.config['PATH'].getpath('TOPAS_location')
        if not INP_path.exists():
            INP_path.mkdir()
        self.tasks = dict()
        self.data = {'TC_location': self.config['PATH'].getpath('TC_location'),
                     'TOPAS_location': self.config['PATH'].getpath('TOPAS_location'),
                     'INPs': {file.name: file for file in INP_path.glob('*.inp')},
                     'tasks': self.tasks,
                     'app': self,
                     }

    def run(self):
        self.gui = AppGUI(data=self.data,
                          logger=self.logger,
                          title=self.name,
                          theme=self.config['APPEARANCE']['theme'])
        self.gui.mainloop()
        # TODO when self.data changed, config needs to be changed too.
        self.config.update_config()

    def process(self, task_id):
        task = self.tasks.get(task_id, {})
        program = task.get('program')
        inp = task.get('inp')
        patterns = task.get('patterns', [])
        processer = task.get('processer')
        if program.absolute() == self.data.get('TC_location').absolute():
            process_func = self.process_TC
        elif program.absolute() == self.data.get('TOPAS_location').absolute():
            process_func = self.process_TP
        else:
            raise
        for pattern in patterns:
            process_func(inp, pattern)
        processer.process_button.configure(style='green.TButton')

    def process_TC(self, inp, raw):
        # raw_path > QT_TMP
        # 'C:\TOPAS5\tc "C:\Users\Administrator\Desktop\batcher\INP\QT_TMP"'
        '''
        function get_params($result_path)
        {
            $lines = (Get-Content $result_path)
            $params = [System.Collections.ArrayList]@()
            $values = [System.Collections.ArrayList]@()
            foreach($line in $lines){
                if(($line -match '^r_exp.*(r_wp)  (\d+\.\d+)')){
                    $_T = $params.Add($Matches[1])
                        $_T = $values.Add($Matches[2])
                }
                if(($line -match 'phase_name "(.*)"'))
                {
                   $_T = $params.Add($Matches[1])
                }
                if($line -match 'MVW\(.* (\d+\.\d+)\`\)$')
                {
                    $_T = $values.Add($Matches[1])
                }
            }
            return $params,$values
        }
        '''
        print(f'processing {raw}')
        return
        # d.replace(d[d.find('xdd'):stop], 'xss sdsdfsdf')
        proc = subprocess.Popen(args=self.TC_location,
                                stdout=subprocess.PIPE,
                                stderr=subprocess.PIPE)
        try:
            outs, errs = proc.communicate(timeout=150)
        except subprocess.TimeoutExpired:
            proc.kill()
            outs, errs = proc.communicate()

    def process_TP(self, inp, raw):
        # TODO: Use AutoHotkey control TOPAS GUI-Mode
        print(f'processing {raw}')
        return


def main(file_path=os.path.dirname(os.path.abspath(__file__))):
    PATH = pathlib.Path(file_path)
    LOGLEVEL = 'DEBUG'
    NAME = 'XRD batch helper'
    CONFIGFILE = PATH / 'config.ini'
    DEFAULT_SETTINGS = {
        'PATH': {
            'INP_path': PATH / 'INP',
            'tc_location': r'C:\TOPAS5\tc.exe',
            'TOPAS_location': r'C:\TOPAS5\Topas.exe',
        },
        'APPEARANCE': {
            'theme': 'clam',
        },
    }
    LOGFILE = PATH / 'app.log'
    logger = LogHandler(name=NAME, level=LOGLEVEL, file=LOGFILE)
    config = Config(file=CONFIGFILE, default=DEFAULT_SETTINGS)
    app = App(name=NAME, path=PATH, logger=logger, config=config)
    app.run()


if __name__ == "__main__":
    main(file_path=os.path.dirname(os.path.abspath(__file__)))
