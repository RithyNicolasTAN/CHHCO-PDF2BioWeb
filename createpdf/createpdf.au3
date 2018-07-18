#include <File.au3>
#include <Array.au3>
#include "MPDF_UDF.au3"
_SetTitle("Pdf2HM")
_SetSubject("")
_SetKeywords("pdf")
_SetUnit($PDF_UNIT_CM)
_SetPaperSize("a4")
_SetZoomMode($PDF_ZOOM_CUSTOM, 90)
_SetOrientation($PDF_ORIENTATION_PORTRAIT)
_SetLayoutMode($PDF_LAYOUT_CONTINOUS)
_InitPDF(@ScriptDir & "\a.pdf")
Local $hFileOpenppdf = FileOpen(@ScriptDir & "\pdf.ntr", $FO_read)
$i = 0
$txt = FileReadLine($hFileOpenppdf)
While $txt <> ""
	_LoadResImage("img" & $i, $txt)
	$txt = FileReadLine($hFileOpenppdf)
	$i = $i + 1
WEnd
$imax = $i - 1
For $i = 0 To $imax
	_BeginPage()
	_InsertImage("img" & $i, 0, 0, _GetPageWidth() / _GetUnit(), _GetPageHeight() / _GetUnit())
	_EndPage()
Next
_ClosePDFFile()
FileClose($hFileOpenppdf)
