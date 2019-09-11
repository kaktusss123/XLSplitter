import openpyxl
import easygui
import pandas as pd
import logging as log
from json import dumps

log.basicConfig(format='%(filename)s[LINE:%(lineno)d]# %(levelname)-8s [%(asctime)s]  %(message)s', level=log.DEBUG)


def ask_file():
    log.debug('Asking file')
    ask_msg = 'Выберите Excel файл'
    ask_title = 'Открыть'
    settings_msg = 'Настройки'

    while 1:
        file = easygui.fileopenbox(msg=ask_msg, title=ask_title, filetypes=[
                                   '*.xlsx'], default='C:/', multiple=True)
        if file:
            log.info('Selected file: {}'.format(file))
            break
    while 1:
        log.debug('Asking settings')
        settings = easygui.multenterbox(settings_msg, fields=[
                                        'Разделить файл по n строк:', 'Удалить дубликаты по столбцам(через запятую):'], values=['100000'])
        if settings is None:
            log.critical('Settings cancelled')
            return
        if not filter(lambda x: x.isdigit(), settings[0]):
            log.warning('Количество строк должно быть заполнено!')
            settings_msg = 'Количество строк должно быть заполнено!'
        else:
            split_by = ''.join(filter(lambda x: x.isdigit(), settings[0]))
            log.debug('All settings set')
            return {
                'file': file, 
                'split_by': int(split_by if split_by else 999999), 
                'drop_by': list(map(lambda x: x.strip(), settings[1].split(','))) if settings[1].strip() else []
            }


def process(inp):
    if type(inp['file']) == str:
        inp['file'] = [inp['file']]
    for file in inp['file']:
        log.info('Reading file {}'.format(file))
        df = pd.read_excel(file)
        log.info('File read successfuly')
        if inp['drop_by']:
            if type(inp['drop_by']) == str:
                log.debug('Filtering duplicates by {}'.format(inp['drop_by']))
                df.drop_duplicates(subset=inp['drop_by'], inplace=True)
            elif type(inp['drop_by']) == list:
                for i in inp['drop_by']:
                    log.debug('Filtering duplicates by {}'.format(i))
                    df.drop_duplicates(subset=i, inplace=True)
        counter = 1
        file_path = file[:file.rfind('.')]
        log.info('Starting splitting')
        while len(df) > inp['split_by']:
            micro_df = df[:inp['split_by']]
            df = df[inp['split_by']:]
            log.debug('Splitted into {} and {}'.format(len(micro_df), len(df)))
            
            log.info('Exporting to {}'.format(f'{file_path}_{counter}.xlsx'))
            micro_df.to_excel(f'{file_path}_{counter}.xlsx', index=False)
            counter += 1
        log.debug('Length of df less than {}'.format(inp['split_by']))
        if len(df) > 0:
            log.debug('Final export to {}'.format(f'{file_path}_{counter}.xlsx'))
            df.to_excel(f'{file_path}_{counter}.xlsx', index=False)
        log.info('Export finished')


if __name__ == '__main__':
    log.info('Starting')
    inp = ask_file()
    log.debug(dumps(inp, ensure_ascii=False))
    log.debug('Starting process')
    process(inp)
    log.info('Process finished')
