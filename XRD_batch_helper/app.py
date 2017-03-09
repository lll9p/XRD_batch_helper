#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import configparser
import csv
import logging
import os
import pathlib
import re
import subprocess
import sys
import time
import tkinter as tk
import tkinter.filedialog as filedialog
from tkinter import font, messagebox, ttk

import psutil


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
        self.inp_path = self.data.get('inp_path')
        self.control_frame = control_frame
        self.INP_names = list(self.data.get('inp_filenames').keys())
        self.INP_filename = tk.StringVar()
        self.INP_filename.set(self.INP_names[0] if self.INP_names else '')
        self.task = {'program': pathlib.Path(self.data.get('TC_location', '')),
                     # INP_filename could be void string
                     'inp': self.inp_path.absolute() / self.INP_filename.get(),
                     'patterns': list(),
                     'processer': self,
                     'output': None,
                     }  # TODO: add a default TC to it
        self.tasks[hash(self)] = self.task
        self.create_widgets()

    def destroy_processer(self):
        self.control_frame.destroy_processer(self)

    def create_widgets(self):
        self.separator = ttk.Separator(master=self, orient=tk.HORIZONTAL)

        self.destry_button = ttk.Button(
            self, text='X',
            command=self.destroy_processer,
            width=2)
        self.destry_button.grid(row=1, column=0)

        self.INP_combobox = ttk.Combobox(
            self, state='readonly',
            textvariable=self.INP_filename,
            values=self.INP_names)
        self.INP_combobox.grid(row=1, column=1)
        self.INP_combobox.bind('<<ComboboxSelected>>', self.select_INP)

        self.INP_choose_button = ttk.Button(
            self, text='I',
            command=self.choose_INP,
            width=2)
        self.INP_choose_button.grid(row=1, column=2)

        self.pattens_choose_button = ttk.Button(
            self, text='F',
            command=self.choose_pattens,
            width=2)
        self.pattens_choose_button.grid(row=1, column=3)

        self.pattens_dir_choose_button = ttk.Button(
            self, text='D',
            command=self.choose_pattens_dir,
            width=2)
        self.pattens_dir_choose_button.grid(row=1, column=4)

        self.output_choose_button = ttk.Button(
            self, text='O',
            command=self.choose_output,
            width=2)
        self.output_choose_button.grid(row=1, column=5)

        self.process_button = ttk.Button(
            self, text='√',
            command=self.process,
            width=2)
        self.process_button.grid(row=1, column=6)

    def choose_INP(self, title='Choose a TOPAS INP file'):
        INP_file = filedialog.askopenfilename(
            defaultextension='.inp',
            filetypes=[('INP files', '.inp')],
            title=title)
        INP_file = pathlib.Path(INP_file)
        if not INP_file.is_file():
            return
        INP_file_short_name = INP_file.name
        if INP_file_short_name in self.INP_names:
            self.logger.info(
                f'User choose a INP ({INP_file}) which duplicates with INP dir\'s.')
        conflict_filenames = list(
            filter(lambda x: x.startswith(INP_file_short_name),
                   self.INP_names))
        if len(conflict_filenames) != 0:
            for i in range(len(conflict_filenames) + 1):
                tmp = INP_file_short_name + f'({i})'
                if tmp not in self.INP_names:
                    INP_file_short_name = tmp
                    break
        self.data['inp_filenames'][INP_file_short_name] = INP_file
        self.INP_filename.set(INP_file_short_name)
        self.task['inp'] = INP_file
        self.update_INP_combobox_values()

    def update_INP_combobox_values(self):
        self.control_frame.update_processers_INP_combobox()

    def select_INP(self, vent):
        '''
        Change task filename after select an INP
        '''
        self.task['inp'] = self.data['inp_filenames'][self.INP_filename.get()]

    def choose_pattens(self):
        patterns = filedialog.askopenfilenames(
            defaultextension='.raw',
            filetypes=[
                ('Bruker raw file patterns', '.raw')],
            title='Choose .raw file(s) of patterns')
        self.task['patterns'].extend(set(patterns))
        self.task['patterns'] = list(set(self.task['patterns']))

    def choose_pattens_dir(self):
        patterns_dir = filedialog.askdirectory(
            title='Choose a dir to read all patterns it contains')
        if os.path.isdir(patterns_dir):
            patterns = list(
                map(lambda f: os.path.join(patterns_dir, f),
                    filter(
                    lambda f: f.endswith('.raw'),
                        os.listdir(patterns_dir))))
            self.task['patterns'].extend(patterns)
            self.task['patterns'] = list(set(self.task['patterns']))
        else:
            messagebox.showinfo(
                title='Alert', message='You choose an invalid directory!')

    def choose_output(self):
        output = filedialog.asksaveasfilename(
            defaultextension='.csv',
            filetypes=[('Comma-separated values file', '.csv'),
                       ('text report', '.txt'),
                       ('microsoft excel', '.xlsx')],
            title='Choose file to output results')
        if not output:
            output = self.app.path / 'result.csv'
        self.task['output'] = pathlib.Path(output)
        return self.task['output']

    def choose_output_alert(self):
        patterns = '\n' + '\n'.join(self.task['patterns'])
        string = f'You must choose an output file for patterns:{patterns}'
        messagebox.showinfo(title='About', message=string)
        self.choose_output()

    def process(self):
        try:
            self.app.process(hash(self))
            self.master.change_button_color(self.process_button, 'green')
        except Exception as e:
            self.master.change_button_color(self.process_button, 'red')
            self.logger.info(f'processerframe.process: {e}')
            raise e


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
            self.master, control_frame=self,
            data=self.data,
            logger=self.logger)
        processer.grid()

    def destroy_processer(self, processer):
        try:
            self.tasks.pop(hash(processer))
            processer.destroy()
        except IndexError as e:
            self.logger.info('processer to be remove but not exists.')
            raise e

    def create_widgets(self):
        self.create_processer_button = ttk.Button(
            self, text='+',
            command=self.create_processer,
            width=2)
        self.create_processer_button.grid(row=0, column=0)
        self.TC_button = ttk.Button(
            self, text='TC',
            command=self.select_TC,
            width=2)
        self.TC_button.grid(row=0, column=1)

        self.process_all_button = ttk.Button(
            self, text='√',
            command=self.process_all,
            width=2)
        self.process_all_button.grid(row=0, column=2)

        self.about_button = ttk.Button(
            self, text='About',
            command=self.about)
        self.about_button.grid(row=0, column=3)

    def select_TC(self):
        self.data['TC_location'] = pathlib.Path(
            filedialog.askopenfilename(
                defaultextension='.exe',
                filetypes=[('Executable', '.exe')],
                title='Find the TOPAS tc.exe'))

    def process_all(self):
        success = 0
        for task in self.tasks.keys():
            try:
                processer = self.tasks[task]['processer']
                processer.process()
            except Exception as e:
                success += 1
                self.logger.info(f'controlframe.process_all: {e}')
        if success == 0:
            self.master.change_button_color(self.process_all_button, 'green')
        elif success != 0:
            self.master.change_button_color(self.process_all_button, 'red')

    def update_processers_INP_combobox(self):
        # Update the INP_name list
        INP_names = list(self.data.get('inp_filenames').keys())
        for task, v in self.tasks.items():
            processer = v.get('processer')
            if isinstance(processer, ProcesserFrame):
                processer.INP_names = INP_names
                processer.INP_combobox.configure(values=INP_names)

    def about(self):
        string = '''Auther : Lao Lilin
Email  : LAOLILIN1@crcement.com
Version: 0.02
Date   : 2017-02-25 15:55
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

    @staticmethod
    def change_button_color(button, color):
        if color == 'green':
            button.configure(style='green.TButton')
        elif color == 'red':
            button.configure(style='red.TButton')

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
        self.inp_path = self.config['PATH'].getpath('inp_path')
        self.TC_location = self.config['PATH'].getpath('TC_location')
        self.TOPAS_location = self.config['PATH'].getpath('TOPAS_location')
        if not self.inp_path.exists():
            self.inp_path.mkdir()
        self.tasks = dict()
        self.data = {
            'TC_location': self.config['PATH'].getpath('TC_location'),
            'TOPAS_location': self.config['PATH'].getpath('TOPAS_location'),
            'inp_filenames': {file.name: file.absolute() for file in self.inp_path.glob('*.inp')},
            'inp_path': self.inp_path,
            'tasks': self.tasks,
            'app': self,
        }

    def process(self, task_id):
        task = self.tasks.get(task_id, {})
        output = task.get('output')
        if output is None:
            task.get('processer').choose_output_alert()
            self.process(task_id)
            return
        patterns = list(set(task.get('patterns', [])))
        if len(patterns) == 0:
            self.logger.info('app.process: processing Nothing')
            return
        results = list()
        program = task.get('program').absolute()
        inp = task.get('inp').absolute()
        tc = self.data.get('TC_location').absolute()
        topas = self.data.get('TOPAS_location').absolute()
        self.logger.info(
            f'app.process: Using {program} process {", ".join(patterns)} with {inp}')
        if program == tc:
            process_func = self.process_TC
        elif program == topas:
            process_func = self.process_TP
        else:
            errs = 'Cannot find process program (TC/TP)'
            self.logger.info(f'app.process: {errs}')
            raise Exception(errs)
        try:
            for pattern in patterns:
                results.append(process_func(tc, inp, pattern))
        except Exception as e:
            self.logger.info(f'app.process: {program} {inp} ERROR:{e}')
            raise e
        # Assume that the same INP generate the same columns of data
        headers = []
        for result in results:
            for key in result:
                if key not in headers:
                    headers.append(key)
        with open(output, 'a', newline='') as f:
            f_csv = csv.DictWriter(f, headers)
            f_csv.writeheader()
            f_csv.writerows(results)

    def process_TC(self, tc, inp, pattern):
        pattern = pathlib.Path(pattern).absolute()
        self.logger.info(f'app.process_TC: processing {inp}-{pattern}')
        if not inp.exists() or not pattern.exists():
            raise
        result = dict()
        r_wp = []
        weights = []
        inp_content = inp.read_text()
        xdd_start = inp_content.find('xdd')
        xdd_stop = inp_content.find('\n\t', xdd_start)
        xdd_string = inp_content[xdd_start:xdd_stop]
        inp_content = inp_content.replace(xdd_string, f'xdd "{pattern}"')
        inp_tmp = pathlib.Path('.') / (str(round(time.time())) + '.inp')
        with open(inp_tmp, mode='w') as f:
            f.write(inp_content)
        inp_out = inp_tmp.with_suffix('.out')
        proc = subprocess.Popen(args=[str(tc), str(inp_tmp)],
                                stdout=subprocess.PIPE,
                                stderr=subprocess.PIPE)
        try:
            outs, errs = proc.communicate(timeout=150)
        except subprocess.TimeoutExpired:
            proc.kill()
            outs, errs = proc.communicate()
            self.logger.info(f'app.process_TC: {outs}-{errs}')
        try:
            with open(inp_out, mode='r') as f:
                inp_output = f.read()
            id = [('id', pattern.with_suffix('').name)]
            r_wp = re.findall(r'.*?(r_wp)  (\d+\.\d+).*',
                              inp_output, re.DOTALL)
            weights = re.findall(
                r'phase_name "(.*?)".*?MVW\( \d+\.\d+, \d+\.\d+`, (\d+\.\d+)`\)',
                inp_output, re.DOTALL)
        except Exception as e:
            self.logger.info(f'app.process_TP: {e}')
        result = {name: value for name, value in id + r_wp + weights}
        inp_tmp.unlink()
        inp_out.unlink()
        return result

    def process_TP(self, tp, inp, pattern):
        # TODO: Use AutoHotkey control TOPAS GUI-Mode
        print(f'processing {pattern}')
        return

    def run(self):
        pid = 42
        try:
            with open('app.pid', mode='r') as f:
                pid = int(f.read())
        except Exception:
            with open('app.pid', mode='w') as f:
                pid = os.getpid()
                f.write(str(pid))
        if psutil.pid_exists(pid):
            proc = psutil.Process(pid)
            if proc.name() == 'python.exe' or proc.cwd() == os.getcwd():
                sys.exit()
        else:
            with open('app.pid', mode='w') as f:
                pid = os.getpid()
                f.write(str(pid))
        self.gui = AppGUI(data=self.data,
                          logger=self.logger,
                          title=self.name,
                          theme=self.config['APPEARANCE']['theme'])
        self.gui.mainloop()
        # TODO when self.data changed, config needs to be changed too.
        self.config.update_config()


def main(file_path):
    PATH = pathlib.Path(file_path)
    LOGLEVEL = 'DEBUG'
    NAME = 'XRD batch helper'
    CONFIGFILE = PATH / 'config.ini'
    DEFAULT_SETTINGS = {
        'PATH': {
            'inp_path': 'INP',
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
    if hasattr(sys, 'frozen'):
        basis = sys.executable
    else:
        basis = sys.argv[0]
    required_folder = os.path.split(basis)[0]
    main(file_path=required_folder)
