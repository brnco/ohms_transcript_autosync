'''
tr
handler for google translate
'''
from googletrans import Translator

def translate(kwargs):
    '''
    translates a string
    returns googletranslate object OR just the translated text
    '''
    trs = Translator()
    translation = trs.translate(kwargs['textO'], src=kwargs['src'], dest=kwargs['dest'])
    if not kwargs['returnTextOnly']:
        return translation
    else:
        return translation.text
