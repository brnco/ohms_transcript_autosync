#!/usr/bin/env python
'''
OHMS Transcript AutoSync

returns OHMS transcript sync data for audio file + transcript
'''
import sr
import tr
import wave
import util as ut
import msoffice
import subprocess
import argparse
import logging as log
from fuzzywuzzy import fuzz
from fuzzywuzzy import process

def format_sync_data(kwargs):
    '''
    takes the sync data list of tuples and returns string on sync data formatted for OHMS
    '''
    ohmsSyncString = '1:'
    for syncPoint in kwargs.syncData:
        ohmsSyncString = ohmsSyncString + "|" + str(syncPoint[0]) + "(" + str(syncPoint[1]) + ")"
    print(ohmsSyncString)
    log.info('ohmsSyncString %s', ohmsSyncString)
    return ohmsSyncString

def get_auDuration(path):
    '''
    returns duration in s for audio file at path
    '''
    f = wave.open(path)
    frames = f.getnframes()
    log.debug('frames %s', frames)
    rate = f.getframerate()
    log.debug('rate %s', rate)
    duration = frames / float(rate)
    f.close()
    return duration

def get_middle_word(phrase):
    '''
    returns the middle word of a string
    '''
    phraseList = phrase.split()
    log.debug('phraseList %s', phraseList)
    middleIndx = float(len(phraseList) -1) / 2
    log.debug('middleIndx %s', middleIndx)
    middleIndx = int(middleIndx)
    return phraseList[middleIndx]

def get_middle_word_line_index(middleWord, line):
    '''
    takes areguments for a word and a line
    finds the word in the line
    if not, returns middle word of line
    middleWord is int
    '''
    lineList = line.split()
    log.debug('lineList %s', lineList)
    if lineList:
        if middleWord in lineList:
            log.debug('middle word "%s" in lineList', middleWord)
            return lineList.index(middleWord)
        else:
            log.debug('middle word "%s" not in lineList', middleWord)
            log.debug('getting the middle word from the retrieved line instead')
            middleOfLine = get_middle_word(line)
            log.debug('middleOfLine %s', middleOfLine)
            log.debug('middleOfLine index %s', lineList.index(middleOfLine))
            return lineList.index(middleOfLine)
    else:
        return 0

def make_auTranscript(kwargs):
    '''
    returns a string of the transcribed audio
    auTranscript is either string or False
    '''
    kwargs.audio = sr.r_audio(kwargs) #parse audio into segment
    try:
        auTranscript = sr.transcribe(kwargs) #transcribe the audio segment
    except:
        auTranscript = False #if transcription doesn't work
    return auTranscript

def get_average_syncpoint_distance(kwargs):
    '''
    returns the average number of lines between sync points
    avg is int
    '''
    totalSyncPoints = len(kwargs.syncData)
    log.debug('totalSyncPoints %s', totalSyncPoints)
    totalLines = kwargs.txOffset
    log.debug('totalLines %s', totalLines)
    if totalSyncPoints < 5:
        avg = 11
    else:
        try:
            avg = totalLines / totalSyncPoints
        except ZeroDivisionError:
            avg = 11
    log.debug("average distance between sync points: %s", avg)
    return avg

def verify_sync_point(kwargs, newline):
    '''
    if sync point deviates markedly from expected out put we need to fix it
    e.g. 958(2)|1858(2)  <- that's not good sync data!
    returns boolean value
    '''
    if len(kwargs.syncData) <=1: #if it's the first sync point I'm sure it's fine
        log.debug('first sync point verifies True')
        return True
    #avg = get_average_syncpoint_distance(kwargs)
    if int(kwargs.syncData[-1][0]): #if there's a previous sync point
        if int(newline) <= int(kwargs.syncData[-1][0]): #if the new sync line is before the previous sync line
            log.debug('new sync point is on line previous to last sync point, syncPoint validation False')
            return False
    if (int(newline) - int(kwargs.syncData[-1][0])) > (kwargs.syncavg *2.5): #if the new sync line is way too far down the page
        log.debug('new sync point is more than 2.5x farther down the transcript, syncPoint validation is False')
        return False
    else:
        log.debug("new sync point validates to True")
        return True


def make_sync_point(kwargs, txTranscript, auTranscript):
    '''
    matches auTranscript to line in txTranscript
    moves through next 25 lines of txTranscript
    matches each line
    finds best match
    returns a tuple of line and word number - both strings
    '''
    matches = []
    bestMatch = (0,0)
    #kwargs.syncavg = get_average_syncpoint_distance(kwargs)
    for line in txTranscript[kwargs.txOffset:kwargs.txOffset + 35]: #walk through the transcript from after the last match
        if txTranscript.index(line) == 0:#if first line then skip
            continue
        else:
            twoline = txTranscript[txTranscript.index(line) - 1 ] + ' ' + line #concatenate lines before current line
            matches.append((txTranscript.index(line), fuzz.token_set_ratio(auTranscript, twoline))) #append (line, match %) to list of matches
    for match in matches: #llop through matches and find the best match
        log.debug("match:")
        log.debug("line: %s", match[0])
        log.debug("fuzz: %s", match[1])
        log.debug("text: %s", txTranscript[match[0]])
        if match[1] > bestMatch[1]:
            bestMatch = match
    log.debug("bestMatch: %s", bestMatch)
    syncLine = int(float(bestMatch[0]))
    syncPointIsGood = verify_sync_point(kwargs, syncLine)
    if not syncPointIsGood: #if the sync line isn't right, faek it
        log.info('there was a problem with the expected sync point')
        log.info('new sync point will be estimated from previous data')
        newline = int(kwargs.syncData[-1][0]) + kwargs.syncavg
        syncPoint = (str(newline), "0")
        log.debug('estimated sync point %s', syncPoint)
        return syncPoint
    else:
        middleWord = get_middle_word(auTranscript)
        middleIndx = get_middle_word_line_index(middleWord, txTranscript[syncLine])
        return (str(syncLine), str(middleIndx))


def walk(kwargs):
    '''
    "walks" through audio file in 60s intervals
    handler for:
    send / receive transcribed audio
    generate corresponding sync point
    increment to next segment 60s further
    syncData is list of tuples of strings, [("lineNumber", "wordNumber")]
    '''
    txTranscript = msoffice.get_text_lines_docx(kwargs.txPath, kwargs) #get transcript text as array of lines
    while kwargs.auOffset < kwargs.auDuration: #while audio start time (in s) less than duration of audio file (in s)
        print('')
        auTranscript = make_auTranscript(kwargs) #transcribe the audio
        kwargs.syncavg = get_average_syncpoint_distance(kwargs)
        if not auTranscript: #if there was a problem transcribing it, try again
            auOffsetOrig = kwargs.auOffset
            kwargs.auOffset = kwargs.auOffset - 4 #set the offset back 4s and try again
            auTranscript = make_auTranscript(kwargs)
            if not auTranscript: #if there was still a problem, faek it
                syncPoint = (str(int(kwargs.syncData[-1][0]) + kwargs.syncavg), "2")
                print("SyncPoint for Minute " + str(len(kwargs.syncData)) + ":")
                print(str(syncPoint))
                kwargs.syncData.append(syncPoint)
                kwargs.auOffset = auOffsetOrig + 60
                continue
            if auTranscript:
                if kwargs.translate:
                    trsReq = {"textO":auTranscript,"src":kwargs.language[:2],"dest":kwargs.translate,"returnTextOnly":False}
                    log.debug("trsReq %s", trsReq)
                    translation = tr.translate(trsReq)
                    log.debug("translation %s", translation)
                    if translation.text:
                        auTranscript = translation.text
                syncPoint = make_sync_point(kwargs, txTranscript, auTranscript)
                log.info("SyncPoint for Minute " + str(len(kwargs.syncData)) + ": %s", syncPoint)
                kwargs.syncData.append(syncPoint)
                kwargs.auOffset = auOffsetOrig + 60
                kwargs.txOffset = int(syncPoint[0]) + 1
                ohmsSyncData = format_sync_data(kwargs)
                continue
        else:
            if kwargs.translate:
                trsReq = {"textO":auTranscript,"src":kwargs.language[:2],"dest":kwargs.translate,"returnTextOnly":False}
                log.debug("trsReq %s", trsReq)
                translation = tr.translate(trsReq)
                log.debug("translation %s", translation)
                if translation.text:
                    auTranscript = translation.text
            syncPoint = make_sync_point(kwargs, txTranscript, auTranscript)
            print("SyncPoint for Minute " + str(len(kwargs.syncData)) + ":")
            print(str(syncPoint))
            print("AudioTranscript: " + auTranscript)
            print("Text Transcript: " + txTranscript[int(syncPoint[0]) - 1] + " " + txTranscript[int(syncPoint[0])])
            kwargs.syncData.append(syncPoint)
            kwargs.auOffset = kwargs.auOffset + 60
            kwargs.txOffset = int(syncPoint[0]) + 1
            log.debug("txOffset %s", kwargs.txOffset)
            log.debug("auOffset %s", kwargs.auOffset)
            ohmsSyncData = format_sync_data(kwargs)
    return kwargs

def convert_syncdata_to_tuples(kwargs):
    '''
    converts user-input sync data to list of tuples
    all vars in here are str
    '''
    syncdata = kwargs.syncData.replace("1:","")
    sd = syncdata.split("|")
    sdt = []
    for sp in sd:
        if not sp:
            continue
        tup = sp.split("(")
        lineNum = tup[0]
        wordNum = tup[1].replace(")","")
        sdt.append((lineNum,wordNum))
    return sdt

def syncdata_convert(kwargs):
    '''
    transformations for sync data
    '''
    sdt = convert_syncdata_to_tuples(kwargs)
    newSyncData = []
    if kwargs.syncdataprint:
        for sp in kwargs.syncdata.split("|"):
            return []
    if kwargs.convert.startswith("+"):
        linediff = int(kwargs.convert.replace("+",""))
        totalLines = sdt[-1]
        totalLines = int(totalLines[0])
        percent = float(linediff / totalLines)
        percentLines = int(float(1/percent))
        i = 1
        for sp in sdt:
            newline = str(int(sp[0]) + int(float(int(sp[0]) / percentLines)))
            newSyncData.append((newline, sp[1]))
    elif kwargs.convert.startswith("-"):
        linediff = int(kwargs.convert.replace("-",""))
        totalLines = sdt[-1]
        totalLines = int(totalLines[0])
        percent = float(linediff / totalLines)
        percentLines = int(float(1/percent))
        i = 1
        for sp in sdt:
            newline = str(int(sp[0]) - int(float(int(sp[0]) / percentLines)))
            newSyncData.append((newline, sp[1]))
    return newSyncData

def init():
    '''
    initialize variable container object
    '''
    parser = argparse.ArgumentParser(description="Sync a transcript for OHMS using Google Speech Recognition")
    parser.add_argument('--transcript',dest='t',help="path to the transcript which you would like to sync")
    parser.add_argument('--audio',dest='a',help="path to audio file which matches transcript")
    parser.add_argument('--language',dest='l',help="Google language code: cloud.google.com/speech-to-text/docs/languages")
    parser.add_argument('--charsInLine',dest='charsInLine',default=77,help="the numbers of characters in a line of the transcript")
    parser.add_argument('--export',dest='export',action='store_true',default=False,help="export txt file with sync data")
    parser.add_argument('--syncdata',dest='syncdata',help="input sync data for transformation, surround with double-quotes")
    parser.add_argument('--convert',dest='convert',help="number of lines to add (+) or subtract (-) from sync data")
    parser.add_argument('--syncdata-print',dest='syncdataprint',action='store_true',default=False, help="print sync data to console on separate lines (for debugging)")
    parser.add_argument('--translate',dest='translate',help="translation language, 2-letter code. Use to sync a translation")
    parser.add_argument('--verbose',dest='loglevel',action='store_const',const=log.INFO,help="display more info during operation")
    parser.add_argument('--debug',dest='loglevel',action='store_const',const=log.DEBUG,default=log.WARNING,help="print everything to console")
    args = parser.parse_args()
    kwargs = ut.dotdict({}) #keyword arguments wrapper
    kwargs.auPath = args.a
    kwargs.txPath = args.t
    kwargs.duration = 8 #default duration for each transcription request, in seconds
    kwargs.auOffset = 56 #default offset for first transcription request, in seconds
    kwargs.txOffset = 0 # defaultoffset for transcription text file, in lines
    kwargs.charsInLine = int(args.charsInLine)
    kwargs.language = args.l #default language for each transcription request
    kwargs.fakeSyncWordIndex = 2 #default word index for sync points which have to be faked
    if args.syncdata:
        kwargs.syncData = args.syncdata
    else:
        kwargs.syncData = []
    kwargs.convert = args.convert
    kwargs.export = args.export
    kwargs.syncdataprint = args.syncdataprint
    kwargs.translate = args.translate
    log.basicConfig(level=args.loglevel)
    log.debug('args %s', args)
    log.debug('kwargs %s', kwargs)
    return kwargs

def main():
    '''
    do the thing
    '''
    print('OHMS Transcript AutoSync is starting...')
    kwargs = init()
    if kwargs.syncData:
        kwargs.syncData = syncdata_convert(kwargs)
        if kwargs.syncData:
            ohmsSyncData = format_sync_data(kwargs)
            print(ohmsSyncData)
    else:
        kwargs.syncData.append(("0","0"))
        kwargs.auFile = sr.load_audio(kwargs.auPath) #load audio file
        kwargs.auDuration = get_auDuration(kwargs.auPath) #get the duration of the audio file
        kwargs = walk(kwargs)
        kwargs.syncData.pop(0)
        ohmsSyncData = format_sync_data(kwargs)
        print(ohmsSyncData)
        if kwargs.export:
            txtfile = open(kwargs.auPath + ".txt", "w")
            txtfile.write(ohmsSyncData)
            txtfile.close()

if __name__ == "__main__":
    main()
