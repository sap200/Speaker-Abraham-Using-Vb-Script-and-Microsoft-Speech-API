'Using Microsoft speech API (SAPI) and VBS to make a computer talk 

Dim textToBeNarrated, speakerAbraham 

textToBeNarrated = InputBox("Text to be read out aloud ", "Speaker Abraham")

Set speakerAbraham = CreateObject("sapi.spvoice")

speakerAbraham.speak textToBeNarrated
