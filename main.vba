Option Explicit

Sub do_all()
  Dim fn As Long
  Dim i As Long

  fn = FreeFile
  i = 0
  
  Dim out As String
  out = Range("out_file").Value
  
  Dim row As String
  Dim scolumn As String
  Dim sheader As String
  Dim tcolumn As String
  Dim paragraph As String
  
  row = Range("row_class").Value
  scolumn = "  " & Range("scolumn").Value
  sheader = "    " & Range("sheader").Value
  tcolumn = "  " & Range("tcolumn").Value
  paragraph = "    " & Range("paragraph").Value
  
  Open out For Append As #fn
    While Range("contents").Offset(i, 1).Value <> ""
      If Range("contents").Offset(i, 0).Value <> "" Then
        Print #fn, row
        Print #fn, scolumn
        Print #fn, sheader
        Print #fn, "      " & Replace(Trim(Range("contents").Offset(i, 0).Value), ":", "")
        Print #fn, tcolumn
      End If

      Print #fn, paragraph
      Print #fn, "      " & Trim(Range("contents").Offset(i, 1).Value)

      i = i + 1
    Wend
  Close #fn
  
End Sub
