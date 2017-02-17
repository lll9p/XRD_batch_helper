#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import configparser
import logging
import os
import pathlib
import tkinter as tk
import tkinter.filedialog as filedialog
from tkinter import font, messagebox, ttk


class App:

    def __init__(self, name, path, logger, config):
        self.path = path
        self.logger = logger
        self.name = name
        self.config = config
        INP_path = self.config['PATH'].getpath('INP_path')
        if not INP_path.exists():
            INP_path.mkdir()
        self.data = {'TC_location': self.config['PATH'].getpath('TC_location'),
                     'TOPAS_location': self.config['PATH'].getpath('TOPAS_location'),
                     'INPs': {file.name: file for file in INP_path.glob('*.inp')}
                     }

    def run(self):
        self.gui = AppGUI(data=self.data,
                          logger=self.logger,
                          title=self.name,
                          theme=self.config['APPEARANCE']['theme'])
        self.gui.mainloop()
        # TODO when self.data changed, config needs to be changed too.
        self.config.update_config()


class AppGUI(ttk.tkinter.Tk):

    def __init__(self, data, logger, title, theme):
        self.title = title
        super().__init__(className=self.title)
        self.wm_title(self.title)
        self.protocol('WM_DELETE_WINDOW', self.on_exit)
        self.resizable(False, False)
        self.minsize(width=font.Font().measure(title + '-' * 6), height=0)
        self.processers = list()
        self.data = data
        self.logger = logger
        self.theme = theme
        self.set_style()
        self.control_frame = ControlFrame(
            master=self, processers=self.processers, data=self.data, logger=self.logger)
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


class ControlFrame(ttk.Frame):

    def __init__(self, master, processers, data, logger, **kwarg):
        super().__init__(master)
        self.processers = processers
        self.data = data
        self.logger = logger
        self.create_widgets()

    def create_processer(self):
        if hasattr(self, 'processers'):
            processer = ProcesserFrame(
                self.master, control_frame=self, processer_id=len(self.processers), data=self.data, logger=self.logger)
            self.processers.append(processer)
            processer.grid()
        else:
            self.processers = list()
            self.create_processer()

    def destroy_processer(self, processer_id):
        try:
            self.processers[processer_id].destroy()
            self.processers[processer_id] = None
        except IndexError:
            self.logger.info('processer_id not exsists.')

    def create_widgets(self):
        self.create_processer_button = ttk.Button(
            self, text='+', command=self.create_processer, width=2)
        self.create_processer_button.grid(row=0, column=0)
        self.TC_button = ttk.Button(
            self, text='TC', command=self.select_TC, width=2)
        self.TC_button.grid(row=0, column=1)

        self.output_choose_button = ttk.Button(
            self, text='√', command=self.process_all, width=2)
        self.output_choose_button.grid(row=0, column=2)

        self.about_button = ttk.Button(
            self, text='About', command=self.about)
        self.about_button.grid(row=0, column=3)

    def select_TC(self):
        self.data['TC_location'] = pathlib.Path(filedialog.askopenfilename(
            defaultextension='.exe', filetypes=[('Executable', '.exe')], title='Find the TOPAS tc.exe'))

    def process_all(self):
        TC_location = self.data.get('TC_location')
        if TC_location.suffix != '.exe':
            self.select_TC()
            self.process_all()
        else:
            for processer in self.processers:
                if isinstance(processer, ProcesserFrame):
                    processer.process()

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

    def __init__(self, master, control_frame, processer_id, data, logger, **kwarg):
        super().__init__(master)
        self.processer_id = processer_id
        self.data = data
        self.logger = logger
        self.patterns = list()
        self.control_frame = control_frame
        self.INP_names = list(self.data.get('INPs').keys())
        self.INP_filename = tk.StringVar()
        self.INP_filename.set(self.INP_names[0] if self.INP_names else '')
        self.output_filename = tk.StringVar()
        self.create_widgets()

    def destroy_processer(self):
        self.control_frame.destroy_processer(self.processer_id)

    def create_widgets(self):
        self.separator = ttk.Separator(master=self, orient=tk.HORIZONTAL)

        self.destry_button = ttk.Button(
            self, text='X', command=self.destroy_processer, width=2)
        self.destry_button.grid(row=1, column=0)

        self.INP_combobox = ttk.Combobox(
            self, state='readonly', textvariable=self.INP_filename, values=self.INP_names)
        self.INP_combobox.grid(row=1, column=1)

        self.INP_choose_button = ttk.Button(
            self, text='C', command=self.select_INP, width=2)
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

        self.output_choose_button = ttk.Button(
            self, text='√', command=self.process, width=2)
        self.output_choose_button.grid(row=1, column=6)

    def select_INP(self, title='Choose a TOPAS INP file'):
        INP_file = filedialog.askopenfilename(
            defaultextension='.inp', title=title)
        if not os.path.isfile(INP_file):
            return
        INP_file_short_name = os.path.split(os.path.abspath(INP_file))[1]
        if INP_file_short_name in self.INP_names:
            self.logger.info(
                f'User choose a INP ({INP_file}) which duplicate with INP dir\'s.')
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
        self.update_INP_combobox_values()

    def update_INP_combobox_values(self):
        self.control_frame.update_processers_INP_combobox()

    def choose_pattens(self):
        patterns = filedialog.askopenfilenames(defaultextension='.raw', filetypes=[(
            'Bruker raw file patterns', '.raw')], title='Choose .raw file(s) of patterns')
        if not hasattr(self, 'patterns'):
            self.patterns = []
        self.patterns.extend(patterns)
        self.patterns = list(set(self.patterns))

    def choose_pattens_dir(self):
        patterns_dir = filedialog.askdirectory(
            title='Choose a dir to read all patterns it contains')
        if not os.path.isdir(patterns_dir):
            self.choose_pattens_dir()
        else:
            patterns = list(map(lambda f: os.path.join(patterns_dir, f), filter(
                lambda f: f.endswith('.raw'), os.listdir(patterns_dir))))
            if not hasattr(self, 'patterns'):
                self.patterns = []
            self.patterns.extend(patterns)

    def choose_output(self):
        output = filedialog.asksaveasfilename(defaultextension='.txt', filetypes=[(
            'text report', '.txt'), ('microsoft excel', '.xlsx')], title='Choose file to output results')
        self.output = output

    def process(self):
        self.logger.info(
            f'Using {self.data["TC_location"]} process {", ".join(self.patterns)} with {self.INP_filename.get()}')


def main():
    PATH = pathlib.Path(os.path.dirname(os.path.abspath(__file__)))
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
    main()
