#!/usr/bin/env python

'''
SpeechRecognition Utilities
handles all SR requests for OHP Microservices
'''
import util as ut
import speech_recognition as sr


def p(string):
    '''
    print handler
    '''
    print(string)

def load_audio(path):
    '''
    loads audio from path into SpeechRecognition object
    '''
    p('loading audio...')
    auFile = sr.AudioFile(path)
    p('audio loaded')
    return auFile

def r_audio(kwargs):
    '''
    returns transcribable audio object from auFile object
    '''
    r = sr.Recognizer()
    with kwargs.auFile as source:
        p('scrubbing to audio offset...')
        audio = r.record(source, kwargs.duration, kwargs.auOffset)
    return audio

def transcribe(kwargs):
    '''
    transcribes audio via Google sr
    '''
    p('transcribing audio...')
    r = sr.Recognizer()
    transcript = r.recognize_google(kwargs.audio, None, kwargs.language)
    p('text received')
    #p(transcript)
    return transcript

def main():
    '''
    do the thing
    '''
    p('starting...')
    kwargs = ut.dotdict({})
    kwargs.auPath = '/home/brnco/devs/media_sources/AlfonsoArau_spa_full.wav'
    kwargs.auFile = load_audio(kwargs.auPath)
    p('parsing audio...')
    kwargs.duration = 8
    kwargs.offset = 88
    kwargs.language = 'pt-BR'
    kwargs.audio = r_audio(kwargs)
    p('audio parsed')
    transcript = transcribe(kwargs)
    p(transcript)


if __name__ == "__main__":
    main()
