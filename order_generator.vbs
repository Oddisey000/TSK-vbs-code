ETE_Program_PDCR = ""
ETE_Program_PDCR = ETE_Program_Name & "_PDCR"
header_ST2 = ""
connectors_list = ""
mGroup_ST2 = mGroup
GetFiles_ST2 = "config_file.ini"
SetFiles_ST2 = GetFiles_ST2 & Left(mGroup,9) & "\"
order_file = "path_to_original_order_file"

header1s_ = "32:" & vbCrLf & "Zn:" & ETE_Program_PDCR & vbCrLf & "Ec:PDCR" & vbCrLf & "Zm:" & vbCrLf & "Em:" & vbCrLf & "Rn:PDCR" & vbCrLf & "Rr:TSK" & vbCrLf
header2s_ = "Lk:" & mGroup_ST2 & vbCrLf & "Pt:Second" & vbCrLf 
headerUV_ = ""
For i = 1 To 256
    headerUV_ = headerUV_ & "Uv:" & vbCrLf
Next 
header3s_ = "Da:" & date() & vbCrLf & "Tn:" & vbCrLf & "Vr:" & vbCrLf & "Af:" & vbCrLf & "Bb:" & vbCrLf & "Mo:1" & vbCrLf & "Gr:" & vbCrLf & "Ic:{`CAT`}mainprgcomponent{`SUB`}0" & vbCrLf & "K4:10.00" & vbCrLf & "K2:70,00" & vbCrLf & "Ks:10000,00" & vbCrLf & "Kt:1,00" & vbCrLf
header4s_ = "Co:0" & vbCrLf & "Cn:0" & vbCrLf & "Cb:0" & vbCrLf & "Tp:1" & vbCrLf & "On:" & vbCrLf & "Oc:{24CC0F72-817B-11D3-9593-0080AD50603A}" & vbCrLf & "Tk:1" & vbCrLf
header5s_ = "Td:1" & vbCrLf & "Ts:1" & vbCrLf & "Te:1" & vbCrLf & "Pb:1" & vbCrLf & "Ab:1" & vbCrLf & "Fo:1" & vbCrLf
header_ST2 = header1s_ & header2s_ & headerUV_ & header3s_ & header4s_ & header5s_ & vbCrLf

Set HFSO = CreateObject("Scripting.FileSystemObject")
If Not HFSO.FolderExists(SetFiles_ST2) Then
    Dialogs.TimeOut=0'sets timeout value
	    Result = Dialogs.Message(g_Parent, "Відсутня структура папок" & vbCrLf & SetFiles_ST2 & vbCrLf & Err_msg, "Увага!", 48 )
    If Result < 0 Then
      Exit Function
    End If
Else
  strArr = split(SetFiles_ST2, "\")
  fileName = "PDCR_electrical_config"

  Set extractedOrderFile = CreateObject("Scripting.FileSystemObject").OpenTextFile(order_file,1)
  order_data = extractedOrderFile.ReadAll
  extractedOrderFile.Close
  Set extractedOrderFile = Nothing

  Set objFileToRead = CreateObject("Scripting.FileSystemObject").OpenTextFile(SetFiles_ST2 & fileName & ".ini",1)
  Do While Not objFileToRead.AtEndOfStream
    strLine = objFileToRead.ReadAll
    strLineArr = Split(strLine, vbCrLf)

    ReDim dataArr(Ubound(strLineArr),13)
    For counter = LBound(strLineArr) To UBound(strLineArr)
      strLine = strLineArr(counter)
	    strDataArr = Split(strLine, ";")
	    For element = LBound(strDataArr) To UBound(strDataArr)
	      dataArr(counter,element) = strDataArr(element)
	    Next
    Next
  Loop

  objFileToRead.Close
  Set objFileToRead = Nothing

  For counter = LBound(dataArr) To UBound(dataArr)
    If InStr(order_data, dataArr(counter, 0)) > 0 Then
      connectors_list = connectors_list & "Vb:" & dataArr(counter,1) & vbCrLf & "Vn:" & dataArr(counter,2) & vbCrLf & "Vq:" & dataArr(counter,3) & vbCrLf & "Vf:" & dataArr(counter,4) & vbCrLf & "T1:" & dataArr(counter,5) & vbCrLf & "T2:" & dataArr(counter,6) & vbCrLf & "X1:" & dataArr(counter,7) & vbCrLf & "X2:" & dataArr(counter,10) & vbCrLf & "Tx:" & dataArr(counter,13) & vbCrLf & vbCrLf   
      connectors_list = connectors_list & "Xc:" & dataArr(counter,7) & vbCrLf & "Sc:" & dataArr(counter,8) & vbCrLf & "Sk:" & dataArr(counter,9) & vbCrLf & vbCrLf
      connectors_list = connectors_list & "Xc:" & dataArr(counter,10) & vbCrLf & "Sc:" & dataArr(counter,11) & vbCrLf & "Sk:" & dataArr(counter,12) & vbCrLf & vbCrLf
    End If
  Next
End If

Set FSO=CreateObject("Scripting.FileSystemObject")
If connectors_list <> "" Then
  Set order_ST2 = FSO.OpenTextFile("C:\CSWin\Order\" & ETE_Program_PDCR & ".ord", 2, True)
  order_ST2.write header_ST2
  order_ST2.write connectors_list
  order_ST2.close
End If

dTime = Right("0" & Datepart("h", Time), 2) & Right("0" & Datepart("n", Time), 2) & Right("0" & Datepart("s", Time), 2)

ETE_PDCR = "C:\CSWin\Order\" & ETE_Program_PDCR & ".ord"
FSO.CopyFile ETE_PDCR, BackUpFolder & "\" & dTime & "_" & ETE_Program_PDCR & ".ord"