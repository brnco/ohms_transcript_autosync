# ohms_transcript_autosync
sync transcripts for OHMS using Google Speech Recognition

developed to support the Oral History Projects department at the Academy of Motion Picture Arts and Sciences

See examples at their [site for the Pacific Standard Time Cultural Festival](pstlala.oscars.org)

For more info about OHP see [their website](https://oscars.org/oral-history)

For more info about OHMS see [their website](http://libraries.uky.edu/libpage.php?lweb_id=11&llib_id=13&ltab_id=1370)


##install

clone or download zip from this repo
tested on python 3.6.x only
external packages can be installed with pip/ pip3:
`pip install SpeechRecognition`
`pip install python-docx`
`pip install wave`
`pip install fuzzywuzzy`


##usage

`python3 ohms_transcript_autosync.py --transcript JaneDoe.docx --audio JaneDoe.wav --language en-US`


##help

`python3 ohms_transcript_autosync.py -h`
