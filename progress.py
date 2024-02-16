import os

class Progress:
    class Color:
        purple = '\033[38;5;5m'
        pink = '\033[38;2;249;38;114m'
        gray = '\033[38;2;58;58;58m'
        light_green = '\033[38;2;114;156;31m'
        green = '\033[38;5;2m'
        clear = '\033[0m'
        def colored(text, color):
            return color + text + Progress.Color.clear
    background_color = Color.gray
    progress_color = Color.pink
    complete_color = Color.light_green
    percent_color = Color.purple
    number_color = Color.green

    __slots__ = ['container', 'length', 'index', 'number_width']

    def __init__(self, container):
        self.container = container
        self.length = len(container)
        self.index = 0
        self.number_width = len(str(self.length))
    def __iter__(self):
        print('\033[?25l', end='')
        for item in self.container:
            self.progress()
            yield item
        self.complete()
        print('\033[?25h', end='')
    def message(self, complete=False):
        default_terminal_columns = os.get_terminal_size().columns
        number = self.Color.colored(f'{self.index:>{self.number_width}}/{self.length}', self.number_color)
        percentage = 100 * self.index // self.length
        percent = self.Color.colored(f'{percentage:>3}%', self.percent_color)
        bar_length = default_terminal_columns - (self.number_width * 2 + 2) - 5
        count = bar_length * self.index // self.length
        bar = self.Color.colored('━' * count, self.complete_color if complete else self.progress_color)
        background_length = bar_length - count
        if background_length != 0:
            background = self.Color.colored('╺' + '━' * (background_length - 1), self.background_color)
        else:
            background = ''
        return f'\r{percent} {bar}{background} {number}'
    def progress(self):
        print(self.message(), end='')
        self.index += 1
    def complete(self):
        print(self.message(True))