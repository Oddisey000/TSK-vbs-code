'/**
'* @var configFile require for path of configuration file with times and additional data in
'* @var isSecondStep should be the global variable, depending on condition second step should be done or not
'* @var currentTime contaign current time in HH:MM:SS format
'* @var timeSlotArrPosition stored position of data inside array where changes need to be done
'*/
configFile = "C:\CSWIN\Order\foamed_grommet_test.ini"
isSecondStep = False
currentTime = cDate(FormatDateTime(now, 4))
timeSlotArrPosition = 0

'// Read confing file and make array of data to work with later on
Set extractedFile = CreateObject("Scripting.FileSystemObject").OpenTextFile(configFile,1)
configDataString = ""
configDataString = extractedFile.ReadAll()
configDataArr = Split(configDataString, vbCrLf)
extractedFile.Close
Set extractedFile = Nothing

'/**
'* Unpack data array and split each element to peaces
'* Check in wich time slot @var currentTime is
'* Define @var endTimeSlot to use it in condition where test during end of the shift should be performed
'* Store in @var timeSlotArrPosition position from text file where times and additional metadata stored
'* If test on shift start not performed yet (metadata less then 1) change this metadata and @var isSecondStep as well
'* If @var currentTime not in any available time slot - change all the positions with metadata to default value of 0
'*/
For i = LBound(configDataArr) To UBound(configDataArr)
  strLineArr = Split(configDataArr(i), "--")
  If (currentTime >= cDate(strLineArr(0)) And currentTime <= cDate(strLineArr(1))) Then
    endTimeSlot = cDate(strLineArr(1))
    timeSlotArrPosition = i
    If strLineArr(2) < 1 Then
      configDataArr(i) = Left(configDataArr(i), Len(configDataArr(i)) - 1) & 1
      isSecondStep = True
    End If
  Else
    configDataArr(i) = Left(configDataArr(i), Len(configDataArr(i)) - 1) & 0
  End If
Next

'// If @var isSecondStep Is True open config file And write all content from array
If isSecondStep Then
  Set objFileToWrite = CreateObject("Scripting.FileSystemObject").OpenTextFile(configFile, 2, True)
  For i = LBound(configDataArr) To UBound(configDataArr)
    objFileToWrite.Write configDataArr(i)
    If Not i = UBound(configDataArr) Then
      objFileToWrite.Write vbCrLf
    End If
  Next
  objFileToWrite.Close
  Set objFileToWrite = Nothing
End If

'// Read confing file and make array of data to work with later on
Set extractedFile = CreateObject("Scripting.FileSystemObject").OpenTextFile(configFile,1)
configDataString = ""
configDataString = extractedFile.ReadAll()
configDataArr = Split(configDataString, vbCrLf)
extractedFile.Close
Set extractedFile = Nothing

'// If @var timeSlotArrPosition not defined (it happens when harness tested not durring time slots) take position with metadata higher then 0 it should be only 1 position
For i = LBound(configDataArr) To UBound(configDataArr)
  strLineArr = Split(configDataArr(i), "--")
  If (strLineArr(2) > 0) Then
    timeSlotArrPosition = i
  End If
Next

'/**
'* If @var currentTime higher then @var endTimeSlot
'* Check if @var timeSlotArrPosition is last element in array
'* If @var timeSlotArrPosition last element then take first one if not - take next element
'* If @var endTimeSlot already passed and time of the next shift has not come - check if metadata not changed yet to default 0 value
'* Open config file and write content from array of data with metadata replacement to default 0 for each position, add empty line if this is not last element
'* Change @var isSecondStep to true, make sure second step will be created and processed
'*/
If (currentTime >= endTimeSlot) Then
  If timeSlotArrPosition = UBound(configDataArr) Then
    strLineArr = Split(configDataArr(0), "--")
  Else
    strLineArr = Split(configDataArr(timeSlotArrPosition + 1), "--")
  End If
  If (currentTime >= endTimeSlot And currentTime <= cDate(strLineArr(0))) Then
    If (Right(configDataArr(timeSlotArrPosition), 1) = 1) Then
      Set objFileToWrite = CreateObject("Scripting.FileSystemObject").OpenTextFile(configFile, 2, True)
      For i = LBound(configDataArr) To UBound(configDataArr)
        configDataArr(i) = Left(configDataArr(i), Len(configDataArr(i)) - 1) & 0
        objFileToWrite.Write configDataArr(i)
        If Not i = UBound(configDataArr) Then
          objFileToWrite.Write vbCrLf
        End If
      Next
      objFileToWrite.Close
      Set objFileToWrite = Nothing
      isSecondStep = True
    End If
  End If
End If