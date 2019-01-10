;Start;
FileDelete SLIPS.DOC

;Set vars
version = 4.7

;Check for TEMPLATE.DOC
IfNotExist, TEMPLATE.DOC 
{
	msgbox Cannot find TEMPLATE.DOC
	exit
}

;Get input file
FileSelectFile, xlsFile_original,,C:\Users\%A_UserName%\Downloads\, CleanSlips %version% - Select File, *.xls*

;Check for input file or cancel to exit
If xlsFile_original =
{
	exit
}

;Copy XLS to script folder
xlsFile_copy = %A_ScriptDir%\LendingRequestReport_copy.xls
FileCopy, %xlsFile_original%, %xlsFile_copy%, 1

;Open XLS file ----------------------------------------------------------------
Progress, zh0 fs12, Tidying up fields...,,CleanSlips %version%
xl         := ComObjCreate("Excel.Application")
xl.Visible := False
book       := xl.Workbooks.Open(xlsFile_copy)
ws         := book.Worksheets(1)
rows       := ws.UsedRange.Rows.Count

;Parse XLS fields
loop, %rows%
{	
	;Set original columns
	title_col            = A%A_Index%
	avail_col            = G%A_Index%
	ship_col             = I%A_Index%
    
    ;Create new columns
    norm_call_number_col = S%A_Index%
    avail_original_col   = T%A_Index% ;Preserve original availability field for testing
    comments_col         = U%A_Index%

	;Set headers for new columns
	if A_Index = 1
	{
        ws.Range(norm_call_number_col).Value := "Normalized Call Number"
        ws.Range(avail_original_col).Value   := "Availability_Original"
        ws.Range(comments_col).Value         := "Comments"
        
        continue
	}
    
    ;Title
	title := ws.Range(title_col).Value
    title := SubStr(title, 1, 40) ;Shortens title to a max length of 40 chars
	
    ;Shipping note
    shippingNote := ws.Range(ship_col).Value
    RegExMatch(shippingNote, "\|\|(.*)\|\|.*\|\|", sn) ;Pulls out requestor's name from shippingNote
    shippingNote := sn1
    
    ;Comments
    comments := ws.Range(ship_col).Value
    RegExMatch(comments, "(.*)\|\|.*\|\|.*\|\|.*", cm) ;Pulls out requestor's name from shippingNote
    comment := cm1
    
    ;Availability
    full_availability  :=
    availability_array :=
    normalized_call_number_full :=
    
    availability       := ws.Range(avail_col).Value
    availability_array := StrSplit(availability, "||")
    
    for index, element in availability_array
    {
        ;Skip if on loan
        if element contains Resource Sharing Long Loan,Resource Sharing Short Loan
        {
            continue
        }
        
        ;Library, location, and call number
        RegExMatch(element, "(.*?),(.*?)\.(.*).*(\(\d{1,3} copy,\d{1,3} available\))", e)
        library                     := e1
        location                    := e2
        call_number                 := e3
        inventory                   := e4
        
        ;Normalize call number and add to ongoing string
        normalized_call_number      := normalize(call_number)
        normalized_call_number_full := normalized_call_number_full . " | " . location . " | " . normalized_call_number
        
        
        ;Add availability note to ongoing string
        full_availability := full_availability . "[" . location . " - " . call_number "] "
        full_availability := RegExReplace(full_availability, " ]", "] ") ;Removes extra spaces around brackets
    }
	
	;Make final changes to XLS sheet
	ws.Range(title_col).Value            := title 
	ws.Range(avail_col).Value            := full_availability 
	ws.Range(ship_col).Value             := shippingNote
    ws.Range(norm_call_number_col).Value := normalized_call_number_full
    ws.Range(avail_original_col).Value   := availability
    ws.Range(comments_col).Value         := comment
}

;Sort spreadsheet by normalized call number
xlAscending := 1
xlYes       := 1
ws.UsedRange.Sort(Key1 := xl.Range("S2"), Order1 := xlAscending,,,,,, Header := xlYes)

;Save and quit XLS file
book.Save()
book.Close
xl.Quit

;Open DOC file ----------------------------------------------------------------
Progress, zh0 fs12, Performing MailMerge...,,CleanSlips %version%
template     = %A_ScriptDir%\TEMPLATE.DOC
saveFile     = %A_ScriptDir%\SLIPS.DOC
wrd         := ComObjCreate("Word.Application")
wrd.Visible := False

;Perform mail merge
doc := wrd.Documents.Open(template)
doc.MailMerge.OpenDataSource(xlsFile_copy,,,,,,,,,,,,,"SELECT * FROM [Sheet0$]")
doc.MailMerge.Execute

;Save and quit DOC file
wrd.ActiveDocument.SaveAs(saveFile)
wrd.DisplayAlerts := False
doc.Close
wrd.Quit

;Delete spreadsheets
FileDelete %xlsFile_original%
FileDelete %xlsFile_copy%

;Finish
Progress, zh0 fs12, Sending to Word...,,CleanSlips %version%
IfNotExist, SLIPS.DOC 
{
	msgbox Cannot find SLIPS.DOC
	exit
}
run winword.exe SLIPS.DOC




; FUNCTIONS -------------------------------------------------------------------

;Takes a string and pads it with a number of trailing characters
add_trailing(str, count, pad)
{
    loop, %count%
    {
        str := str . pad
    }
    
    return str
}

;Takes a call number string and normalizes for alpha numeric sorting
;*Credit to Bill Dueber at http://robotlibrarian.billdueber.com)
normalize(call_number)
{
    ;Normalize call number (*Credit to Bill Dueber at http://robotlibrarian.billdueber.com)
    RegExMatch(call_number, "x)^\s*([A-Z]{1,3})\s*(\d+(?:\s*\.\s*\d+)?)?\s*(?:\.?\s*([A-Z]+)\s*(\d+)?)?(?:\.? \s*([A-Z]+)\s*(\d+)?)?\s*(.*?)\s*$", m)
    
    alpha   := Format("{:-3s}", m1)
    num     := Format("{:09.4f}", m2)
    c1alpha := Format("{:-2s}", m3)
    c1num   := add_trailing(m4, 4 - StrLen(m4), 0)
    c1num   := Format("{:4s}", c1num)
    c2alpha := Format("{:-2s}", m5)
    c2num   := add_trailing(m6, 4 - StrLen(m6), 0)
    c2num   := Format("{:4s}", c2num)
    extra   := Format(" {:s}", m7)
    normalized_call_number = %alpha%%num%%c1alpha%%c1num%%c2alpha%%c2num%%extra%

    return normalized_call_number
}