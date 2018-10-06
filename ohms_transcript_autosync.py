'''
OHMS Transcript AutoSync

returns OHMS transcript sync data for audio file + transcript
'''

import sr
import wave
import util as ut
import msoffice
import subprocess
import argparse
from fuzzywuzzy import fuzz
from fuzzywuzzy import process

def format_sync_data(kwargs):
    '''
    takes the sync data list of tuples and returns string on sync data formatted for OHMS
    '''
    ohmsSyncString = '1:'
    for syncPoint in kwargs.syncData:
        ohmsSyncString = ohmsSyncString + "|" + str(syncPoint[0]) + "(" + str(syncPoint[1]) + ")"
    return ohmsSyncString

def get_auDuration(path):
    '''
    returns duration in s for audio file at path
    '''
    f = wave.open(path)
    frames = f.getnframes()
    rate = f.getframerate()
    duration = frames / float(rate)
    f.close()
    return duration

def get_middle_word(phrase):
    '''
    returns the middle word of a string
    '''
    phraseList = phrase.split()
    #print(phraseList)
    middleIndx = float(len(phraseList) -1) / 2
    middleIndx = int(middleIndx)
    return phraseList[middleIndx]

def get_middle_word_line_index(middleWord, line):
    '''
    takes areguments for a word and a line
    finds the word in the line
    if not, returns middle word of line
    '''
    lineList = line.split()
    if middleWord in lineList:
        return lineList.index(middleWord)
    else:
        middleOfLine = get_middle_word(line)
        return lineList.index(middleOfLine)

def make_auTranscript(kwargs):
    '''
    returns a string of the transcribed audio
    '''
    #parse the audio into a segment
    kwargs.audio = sr.r_audio(kwargs)
    try:
        #transcribe the audio segment
        auTranscript = sr.transcribe(kwargs)
    except:
        auTranscript = False
    return auTranscript

def verify_sync_point(kwargs, newline):
    '''
    if sync point deviates markedly from expected out put we need to fix it
    e.g. 958(2)|1858(2)  <- that's not good sync data!
    '''
    if len(kwargs.syncData) <1:
        return True
    totalSyncPoints = len(kwargs.syncData)
    totalLines = kwargs.txOffset
    try:
        avg = totalLines / totalSyncPoints
    except ZeroDivisionError:
        avg = 11
    #print(avg)
    #print(str(int(newline) - int(kwargs.syncData[-1][0])))
    if (int(newline) - int(kwargs.syncData[-1][0])) > (avg *2.5):
        return False
    else:
        return True

def make_sync_point_backup(kwargs, txTranscript, auTranscript):
    '''
    creates a sync point based on the best match in next avg * 2 lines
    called when sync point verification fails or a line is unmatched in the transcript
    '''
    totalSyncPoints = len(kwargs.syncData)
    totalLines = kwargs.txOffset
    matches = []
    bestMatch = (0,0)
    try:
        avg = totalLines / totalSyncPoints
        avg = int(avg)
    except ZeroDivisionError:
        avg = 11
    endOffsetIndx = kwargs.txOffset + (avg*2)
    if endOffsetIndx >= len(txTranscript):
        endOffsetIndx = len(txTranscript) - 1
    #print(endOffsetIndx)
    for line in txTranscript[kwargs.txOffset:endOffsetIndx]:
        print(line)
        twoline = txTranscript[txTranscript.index(line) - 1 ] + ' ' + line
        matches.append((txTranscript.index(line), fuzz.token_set_ratio(auTranscript, twoline)))
    #print(matches)
    for match in matches:
        print(match)
        emptyLineCheck = txTranscript[match[1]].split()
        if emptyLineCheck:
            if match[1] > bestMatch[1]:
                bestMatch = match
    syncLine = bestMatch[0]
    #print(syncLine)
    middleWord = get_middle_word(auTranscript)
    middleIndx = get_middle_word_line_index(middleWord, line)
    return (syncLine, middleIndx)

def make_sync_point(kwargs, txTranscript, auTranscript):
    '''
    matches auTranscript to line in txTranscript
    returns a tuple of line and word number
    '''
    #walk through the transcript from after the last match
    for line in txTranscript[kwargs.txOffset:]:
        #if first line then skip
        if txTranscript.index(line) == 0:
            continue
        else:
            #concatenate lines before current line
            twoline = txTranscript[txTranscript.index(line) - 1 ] + ' ' + line
        #run the comparison btw transcribed segment and two-line text
        if fuzz.token_set_ratio(auTranscript, twoline) > kwargs.fuzzRatio:
            #get the middle word in the line, roughly the word at 00:00
            print('Transcribed Audio: ' + auTranscript)
            print('Transcribed Texto: ' + twoline)
            print('Fuzz Ratio: ' + str(fuzz.token_set_ratio(auTranscript, twoline)))
            syncPointIsGood = verify_sync_point(kwargs, txTranscript.index(line))
            print(syncPointIsGood)
            if not syncPointIsGood:
                syncPoint = make_sync_point_backup(kwargs, txTranscript, auTranscript)
                return syncPoint
            else:
                middleWord = get_middle_word(auTranscript)
                middleIndx = get_middle_word_line_index(middleWord, line)
                return (str(txTranscript.index(line)), str(middleIndx))
        else:
            continue
    syncPoint = make_sync_point_backup(kwargs, txTranscript, auTranscript)
    return syncPoint

def walk(kwargs):
    '''
    "walks" through audio file in 60s intervals
    '''
    #get transcript text as array of lines
    print(kwargs.fuzzRatio)
    txTranscript = msoffice.get_text_lines_docx(kwargs.txPath, kwargs)
    #while audio start time (in s) less than duration of audio file (in s)
    while kwargs.auOffset < kwargs.auDuration:
        #transcribe the audio
        auTranscript = make_auTranscript(kwargs)
        if not auTranscript:
            print("there was an error transcribing the audio")
            auOffsetOrig = kwargs.auOffset
            kwargs.auOffset = kwargs.auOffset - 4
            kwargs.duration = 15
            auTranscript = make_auTranscript(kwargs)
            if not auTranscript:
                kwargs.auOffset = kwargs.auOffset - 7
                auTranscript = make_auTranscript(kwargs)
                if not auTranscript:
                    kwargs.auOffset = kwargs.auOffset + 15
                    auTranscript = make_auTranscript(kwargs)
                    if not auTranscript:
                        syncPoint = (kwargs.syncData[-1][0] + 11, 2)
                        kwargs.syncData.append(syncPoint)
                        kwargs.auOffset = kwargs.auOffset - 19
                        kwargs.duration = 8
                        continue
            syncPoint = make_sync_point_backup(kwargs, txTranscript, auTranscript)
            kwargs.syncData.append(syncPoint)
            kwargs.auOffset = auOffsetOrig + 60
            ohmsSyncData = format_sync_data(kwargs)
            print(ohmsSyncData)
            print('')
            continue
        else:
            syncPoint = make_sync_point(kwargs, txTranscript, auTranscript)
            kwargs.syncData.append(syncPoint)
            #increment offsets
            kwargs.auOffset = kwargs.auOffset + 60
            kwargs.txOffset = int(syncPoint[0]) + 1
            print("txOffset: " + str(kwargs.txOffset))
            ohmsSyncData = format_sync_data(kwargs)
            print(ohmsSyncData)
            print('')
    return kwargs

def init():
    '''
    initialize variable container object
    '''
    parser = argparse.ArgumentParser(description="Sync a transcript for OHMS using Google Speech Recognition")
    parser.add_argument('--transcript',dest='t',help="path to the transcript which you would like to sync")
    parser.add_argument('--audio',dest='a',help="path to audio file which matches transcript")
    parser.add_argument('--language',dest='l',help="Google language code: cloud.google.com/speech-to-text/docs/languages")
    parser.add_argument('--fuzz-ratio',dest='fr',default=65,help="the fuzz ratio threshold for match (generally 55-75 is ok)")
    parser.add_argument('--charsInLine',dest='charsInLine',default=90,help="the numbers of characters in a line of the transcript")
    args = parser.parse_args()
    kwargs = ut.dotdict({}) #keyword arguments wrapper
    '''
    #next three lines grip API_KEY for Google Cloud Speech from hidden 'passwords' file
    config = configparser.ConfigParser()
    config.read('passwords')
    kwargs.GCSapiKey = config['Google Cloud Speech']['API_KEY']
    '''
    #next two lines grip paths to wave file and transcript
    kwargs.auPath = args.a
    kwargs.txPath = args.t
    kwargs.auFile = sr.load_audio(kwargs.auPath) #load audio file
    kwargs.duration = 8 #default duration for each transcription request, in seconds
    kwargs.auOffset = 56 #default offset for first transcription request, in seconds
    kwargs.txOffset = 0 # defaultoffset for transcription text file, in lines
    kwargs.charsInLine = args.charsInLine
    kwargs.language = args.l #default language for each transcription request
    kwargs.fuzzRatio = args.fr #default fuzzy matching ratio, 0-100, 55-75 generally ok
    kwargs.fakeSyncWordIndex = 2 #default word index for sync points which have to be faked
    kwargs.auDuration = get_auDuration(kwargs.auPath) #get the duration of the audio file
    kwargs.syncData = [] #init object to contain sync data
    return kwargs

def main():
    '''
    do the thing
    '''
    print('OHMS Transcript AutoSync is starting...')
    kwargs = init()
    kwargs = walk(kwargs)
    ohmsSyncData = format_sync_data(kwargs)
    print(ohmsSyncData)

if __name__ == "__main__":
    main()
