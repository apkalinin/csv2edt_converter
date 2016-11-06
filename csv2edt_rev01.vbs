' ����������� ������� ����� ������ �� csv � edt
const c_strScriptVer = "1.1"
'------------------------------------------------------------------------------

' �������� ��������������� ��������
const c_strLogsSubFolder = "logs"
const c_strSourceSubFolder = "source\"
'const c_strSourceSubFolder = ""
const c_strResultsSubFolder = "results"

' ���� � �������� ������ ��� �������
c_strSourceFile = "__test01"

' ���������� ������
const c_strLogExtension = ".log"
const c_strResultsExtension = ".edt"
const c_strSourceExtension = ".csv"

' �������������� ������ csv �����
' �� c_WinUserSetting_ListSeparator
'const c_CSVSplitSymbol = ";"

' ������ ������� �����
const c_EdtDecimalSymbol = "."

' ������� ����� ��� ����������� ������ �����
const c_strNewBlock = "strain(%)"

' ������ ����� �������� ������
const c_nOutputWidth = 80
const c_nElementLength = 10
const c_nNumberLength = 5
c_nElementsInOneRow = c_nOutputWidth / c_nElementLength
c_nNumbersInOneRow = c_nOutputWidth / c_nNumberLength

' ������ ��������������� ���� � ���������� ���������
const c_nHeaderLenth_NumberOfElements = 5

' ��������� ������������ ���������
const c_nBlockHeader01_FullLength = 16
const c_nBlockHeader01_LeadingSpaces = 4
const c_nBlockHeader02_FullLength = 59	' ��� ������-�� ������� ������� ����� ���������

' ����� ����������
const c_strBlockHeader01_01 = "Mat"
const c_strBlockHeader01_02 = "_Busher"

const c_strBlockHeader02_01 = "Modulus values for Material No. "
const c_strBlockHeader02_02 = "Damping values for Material No. "

const c_strGlogalHeader01 = "SHAKE2000 EDT File - SI"
const c_strGlogalHeader02 = "Option 1 - Dynamic material properties"
const c_strGlogalHeader03 = "    1"

' �� ������ ������� ����������� ����������
const c_nPlaceToRound = 5

const c_nBlockTypeModulus = 0
const c_nBlockTypeDamping = 1

' ���������� ���������� 
set g_FSO = CreateObject("Scripting.FileSystemObject")
set g_Shell = WScript.CreateObject("WScript.Shell")
const c_strLocalSettingRegPath = "HKCU\Control Panel\International\"

dim c_WinUserSetting_DecimalSymbol
dim c_WinUserSetting_ListSeparator

' ������� ���� ����� �������� ������
dim g_LogName

' �������� ������ ������������ �����, 
' �������� ������ ���������� ��� ��������� ����������
dim g_arrResultArray()
g_nResultArrayIndex = 3
 '------------------------------------------------------------------------------

' ����� ������� �������
if (WScript.Arguments.Count) then
  startConversion(Wscript.Arguments(0))
else
  startConversion("")
end if

'------------------------------------------------------------------------------

function startConversion(astrSourcePath)  
  startConversion = -1
  
  dim dtBuildStart
  dtBuildStart = Now()  
  
  ' ��������� �������� ������
  if (astrSourcePath <> "") then
    strSourceFilePath = astrSourcePath
    c_strSourceFile = g_FSO.GetBaseName(astrSourcePath)
  else 
    strSourceFilePath = getLocalPath() & c_strSourceSubFolder & c_strSourceFile & c_strSourceExtension
  end if
  
	' ������� ���� �����
  g_LogName = createFile(getLocalPath(), c_strLogsSubFolder, c_strLogExtension)
  logOut "Starting csv2edt conversion, script version: " & c_strScriptVer
	logOut "created log file: " & g_LogName
 
  logOut "trying to load source file: " & strSourceFilePath
  if not g_FSO.fileExists(strSourceFilePath) then
		logOut "couldn't find source file, exiting script"
		startConversion = -1
		exit function	
  end if
  
	' �������� ��� � ����� ���� ������
  dim instructionsFile
  set instructionsFile = g_FSO.OpenTextFile(strSourceFilePath, 1, false, 0)
	if instructionsFile.AtEndOfStream then
    logOut "source file is empty"
    exit function
  end if
  
  logOut "reading windows user local regional setting"
  readWinUserRegionalSetting()
  logOut "decimal symbol: " & c_WinUserSetting_DecimalSymbol 
  logOut "string separator: " & c_WinUserSetting_ListSeparator
    
  redim preserve g_arrResultArray(g_nResultArrayIndex)
  ' ����� ����� ��� ������ 
  g_arrResultArray(0) = c_strGlogalHeader01
  g_arrResultArray(1) = c_strGlogalHeader02
  g_arrResultArray(2) = c_strGlogalHeader03
  ' 3 ��������� � �����, ����� ���� ��� ������ �������� ����� ���������� ���������
  
	' ������ �������� ����� ������
	dim arrCurrentBlock
  arrCurrentBlock = Array()
	
	' �������� ������ ������
	dim arrCurrentLine
	arrCurrentLine = array()
	
	dim nSourceLineNumber
	dim nBlockLineNumber
	dim nResultLinesNumber
	
	dim fNewBlockFlag
	fNewBlockFlag = false
	
	dim nCurrentLineElementsNumber
	dim nMaterialNumber
	
	dim nBlockType ' 0 - modulus, 1 - damping
	nBlockType = c_nBlockTypeModulus
	
  logOut "initializing complete"
  logOut ""
  while not instructionsFile.AtEndOfStream
    nSourceLineNumber = nSourceLineNumber + 1
		
    ' ���� � ������� ���������� ��� ��������� ������������ �� ������� ����� �������
    ' ����������� �� ���� ������
    
    startConversion = parseLine(_
			nSourceLineNumber,_
			instructionsFile.ReadLine(),_
			arrCurrentLine,_
			nCurrentLineElementsNumber,_
			nMaterialNumber,_
			fNewBlockFlag) 
    
    if (fNewBlockFlag) then   
      
      extend = extendResultArray()
      g_arrResultArray(g_nResultArrayIndex) = createBlockHeader(_
        nBlockType, _
        nMaterialNumber, _
        nCurrentLineElementsNumber)
      
      logOut ""
      logOut "block header: " & vbNewLine & g_arrResultArray(g_nResultArrayIndex)

      ' �������� ���� ������ ���, ��� � csv �� ������� ����
      if (nBlockType = c_nBlockTypeModulus) then
        nBlockType = c_nBlockTypeDamping
      else
        nBlockType = c_nBlockTypeModulus
      end if
      
      fNewBlockFlag = 0
    end if
    
    startConversion = fillArrayWithData(_
      arrCurrentLine, _ 
      nCurrentLineElementsNumber)    
  wend

  logOut ""
  logOut "information was converted to edt format"
  g_arrResultArray(3) = addLeadingSpaces(_
    nMaterialNumber, _
    c_nHeaderLenth_NumberOfElements, _
    len(nMaterialNumber))
   
  logOut "total number of materials: " & g_arrResultArray(3)
   
  logOut "creating last string"
  
  strUnlock = g_arrResultArray(3)
  startConversion = createLastString(_
    nMaterialNumber, _
    strUnlock)
  
  dim strEdtFilePath
	startConversion = exportResultsToEdt(strEdtFilePath)
  logOut "finishing csv2edt "
  
  logOutMsgBox _
    "csv2edt finished for file: " & vbNewLine & strSourceFilePath & vbNewLine & vbNewLine & _ 
    "result file: " & vbNewLine & strEdtFilePath & vbNewLine & vbNewLine & _ 
    "log file: " & vbNewLine & g_LogName & vbNewLine & vbNewLine & _
    "runtime: " & DateDiff("s", dtBuildStart, Now()) & " seconds"
  
  startConversion = 0
end function
'------------------------------------------------------------------------------

sub readWinUserRegionalSetting()
  c_WinUserSetting_DecimalSymbol = g_Shell.RegRead(c_strLocalSettingRegPath & "sDecimal")
  c_WinUserSetting_ListSeparator = g_Shell.RegRead(c_strLocalSettingRegPath & "sList")
end sub
'------------------------------------------------------------------------------

function createLastString(anMaterialNumber, astrMaterialNumber)
  dim realIterator
	realIterator = 0
  
  extend = extendResultArray()
  
  g_arrResultArray(g_nResultArrayIndex) = astrMaterialNumber
  
  for i = 1 to anMaterialNumber
    if (realIterator = c_nNumbersInOneRow - 1) then
      extend = extendResultArray()
      realIterator = 0
    end if
      
    g_arrResultArray(g_nResultArrayIndex) = g_arrResultArray(g_nResultArrayIndex) & _
      addLeadingSpaces(_
        i, _
        c_nHeaderLenth_NumberOfElements, _
        len(i))

    realIterator = realIterator + 1
  next
  
  
end function
'------------------------------------------------------------------------------

function getArgsPath()
  
  const c_strPathName = "-path="

  for nArgIndex = 0 to WScript.Arguments.Count - 1
    if InStr(WScript.Arguments(nArgIndex), c_strPathName) = 1 then
      getArgsPath = Mid(WScript.Arguments(nArgIndex), Len(c_strPathName) + 1)
    end if
  next

end function
'------------------------------------------------------------------------------

function exportResultsToEdt(astrEdtFilePath)
  ' ������� ���� �����
  exportFile = createFile(getLocalPath(), c_strResultsSubFolder, c_strResultsExtension)
  logOut "exporting results to: " & exportFile
	
  dim edtFile
  set edtFile = g_FSO.OpenTextFile(exportFile, 8, true, 0)
  
  for i = 0 to ubound(g_arrResultArray)
    edtFile.WriteLine(g_arrResultArray(i))
  next
  
  edtFile.close()
  
  logOut "edt file created"
  astrEdtFilePath = exportFile
  
end function
'------------------------------------------------------------------------------

function extendResultArray()
  redim preserve g_arrResultArray(g_nResultArrayIndex + 1)
  g_nResultArrayIndex = ubound(g_arrResultArray)
end function
'------------------------------------------------------------------------------

function fillArrayWithData(aarrCurrentLine, anCurrentLineElementsNumber)
	 
	dim realIterator
	realIterator = 0

  extend = extendResultArray()
	
  logOut "initial array size: " & g_nResultArrayIndex  

  dim maxIter
      
  ' 2 ������ ����������  
  for i = 2 to anCurrentLineElementsNumber + 1
		'if (i > 1) then
      'logOut "i: " & i & "; realIterator: " & realIterator
    if (realIterator = c_nElementsInOneRow) then
      extend = extendResultArray()
      realIterator = 0
    end if
    
    'logOut "current line: " & aarrCurrentLine(i)  
    g_arrResultArray(g_nResultArrayIndex) = _
      g_arrResultArray(g_nResultArrayIndex) & aarrCurrentLine(i)
    'logOut "ubound(aarrCurrentLine): " & ubound(aarrCurrentLine) & "; i: " & i
    logOut "current line:" & vbNewLine & g_arrResultArray(g_nResultArrayIndex)
    
    realIterator = realIterator + 1
    
    'end if
	next
  
  logOut "new array size: " & g_nResultArrayIndex 
  
end function

'------------------------------------------------------------------------------

function createBlockHeader(anBlockType, anMaterialNumber, anCurrentLineElementsNumber)
	
	dim strNumberOfElems
	strNumberOfElems = addLeadingSpaces(_
		anCurrentLineElementsNumber,_
		c_nHeaderLenth_NumberOfElements,_
		len(anCurrentLineElementsNumber))
	
	dim strHeader01
	strHeader01 = string(c_nBlockHeader01_LeadingSpaces, " ") &_
		c_strBlockHeader01_01 &_ 
		anMaterialNumber &_ 
		c_strBlockHeader01_02
	
	if (len(strHeader01) > c_nBlockHeader01_FullLength) then
		strHeader01 = left(strHeader01, c_nBlockHeader01_FullLength)
	else
		strHeader01 = strHeader01 &_
			string(c_nBlockHeader01_FullLength - len(strHeader01), " ")
	end if
	
	dim strHeader02
	if (anBlockType = c_nBlockTypeModulus) then
    strHeader02 = c_strBlockHeader02_01 & anMaterialNumber
  elseif (anBlockType = c_nBlockTypeDamping) then
    strHeader02 = c_strBlockHeader02_02 & anMaterialNumber
  end if
	if (len(strHeader02) > c_nBlockHeader02_FullLength) then
		strHeader02 = left(strHeader02, c_nBlockHeader02_FullLength)
	end if
	
	createBlockHeader = strNumberOfElems & strHeader01 & strHeader02
end function
'------------------------------------------------------------------------------

function parseLine(anSourceLineNumber,_ 
	aLine,_ 
	arrCurrentLine,_
	anCurrentLineElementsNumber,_
	anMaterialNumber,_
	afNewBlockFlag)
	
  parseLine = -1
	
	arrCurrentLine = Split(aLine, c_WinUserSetting_ListSeparator, -1, 1)
	nSourceLineElementsNumber = ubound(arrCurrentLine)
	
  logOut ""
	logOut "parsing line: " & anSourceLineNumber & _
    "; number of elements is source line: " & nSourceLineElementsNumber + 1
	
	if ((nSourceLineElementsNumber < 1) or (nSourceLineElementsNumber > 21)) then  
    logOutMsgBox "Bad line: " & sourceLineNumber & _
      " - inappropriate number of elements: " & nSourceLineElementsNumber
    exit function
  end if
	
	if (arrCurrentLine(1) = c_strNewBlock) then
		afNewBlockFlag = true
		anMaterialNumber = arrCurrentLine(0)
		
		logOut "current line begins new data block, material number: " & anMaterialNumber
	end if
	
	parseLine = convertValues(arrCurrentLine,_ 
    nSourceLineElementsNumber, _
    anCurrentLineElementsNumber)
end function
'------------------------------------------------------------------------------

function convertValues(arrCurrentLine, anSourceLineElementsNumber, anCurrentLineElementsNumber)
	convertValues = -1
	anCurrentLineElementsNumber = 0
	
	for i = 0 to anSourceLineElementsNumber  
		logOut "converting element # " & i + 1
		logOut "csv form: '" & arrCurrentLine(i) & "'"
		
		if i < 2 then
			logOut "element skipped"
    else
			dim nElementLength
			nElementLength = len(arrCurrentLine(i))
			
			if (nElementLength = 0) then
				logOut "element skipped, data length: " & nElementLength
			else 
        
        arrCurrentLine(i) = convertAndRound(arrCurrentLine(i))
        
        arrCurrentLine(i) = replace(arrCurrentLine(i), ",", ".")

				dim nNumberOfCharactersToCopy
        nNumberOfCharactersToCopy = 0
				nNumberOfCharactersToCopy = calcNumOfChars(arrCurrentLine(i))
				
				' 0 ����� �� ����� �� �������
				if (left(arrCurrentLine(i), 1) = "0") then
					' ���������� ������ 0
          if (arrCurrentLine(i) <> "0") then
            nNumberOfCharactersToCopy = nNumberOfCharactersToCopy - 1
          end if
				end if
				
				' ��������� ���������� �������      
				arrCurrentLine(i) = addLeadingSpaces(arrCurrentLine(i), _
          c_nElementLength, _
          nNumberOfCharactersToCopy)
				logOut "edt form: '" & arrCurrentLine(i) & "'"
				
				anCurrentLineElementsNumber = anCurrentLineElementsNumber + 1
			end if
		end if
	next	
	
	logOut "number of converted elements: " & anCurrentLineElementsNumber
	
	if (anCurrentLineElementsNumber > 0) then
		convertValues = 0
	'end if
	else
		logOut "source string had no elements to convert. script will continue for now"
	end if
	
end function
'------------------------------------------------------------------------------

function calcNumOfChars(aSourceStr)
  pointPos = InStr(aSourceStr,".")
 
  if (pointPos > 0) then
    if (len(aSourceStr) >= pointPos + c_nPlaceToRound) then
      calcNumOfChars = pointPos + c_nPlaceToRound
    else
      calcNumOfChars = len(aSourceStr)
    end if
  else
    calcNumOfChars = len(aSourceStr)
  end if
 
end function
'------------------------------------------------------------------------------

function addLeadingSpaces(astrInputString, anNeededLength, anNumOfCharactersToCopy)
	
	addLeadingSpaces = _
		string(anNeededLength - anNumOfCharactersToCopy, " ") & _
		right(astrInputString, anNumOfCharactersToCopy)
		
end function
'------------------------------------------------------------------------------

function convertAndRound(aValue)
  
  if not (InStr(aValue, c_WinUserSetting_DecimalSymbol)) then
    aValue = Replace(aValue, ".", ",")
  end if 
  
  convertAndRound = round(CDbl(aValue), 5)
  'convertAndRound = round(CDbl(Replace(aValue,".",",")), 5)
end function
'------------------------------------------------------------------------------

' ���������� �� c_nPlaceToRound ������� � ����������� ��
' (c_nPlaceToRound + 1) �������
function roughRounding(aValue)
 
	pointPos = InStr(aValue,".")
	sixthCharacter = mid(aValue, pointPos + c_nPlaceToRound + 1, 1)
	fifthCharacter = mid(aValue, pointPos + c_nPlaceToRound, 1)
	
	if (sixthCharacter > "4") then
		fifthCharacter = fifthCharacter + 1
	end if

	roughRounding = left(aValue, pointPos + c_nPlaceToRound - 1) & fifthCharacter
end function
'------------------------------------------------------------------------------

' ������� ����������� ���� � �������
function getLocalPath()
  getLocalPath = Mid(WScript.ScriptFullName, 1, _
    Len(WScript.ScriptFullName) - Len(WScript.ScriptName))
end function
'------------------------------------------------------------------------------

' ���� ��� ���� �������
function getDateFormat()
  dtNow = Now()

  dim strResult
  
  strResult = Year(dtNow) '& "_"

  if Month(dtNow) < 10 then
    strResult = strResult & "0"
  end if
  strResult = strResult & Month(dtNow) '& "_"

  if Day(dtNow) < 10 then
    strResult = strResult & "0"
  end if
  strResult = strResult & Day(dtNow) & "_"
  
  if Hour(dtNow) < 10 then
    strResult = strResult & "0"
  end if
  strResult = strResult & Hour(dtNow) & "_"
  
  if Minute(dtNow) < 10 then
    strResult = strResult & "0"
  end if 
  strResult = strResult & Minute(dtNow) & "_"
  
  if Second(dtNow) < 10 then
    strResult = strResult & "0"
  end if
  strResult = strResult & Second(dtNow)

  getDateFormat = strResult

end function
'------------------------------------------------------------------------------

' �������� ����������� ����� ����� �� ����
' astrPath - ���� � �����
' astrExtension - ���������� ������������ �����
function generateName(astrPath, astrExtension)
  dim strTemp, strResult
  strTemp = astrPath & "\" & "csv2edt__" &  c_strSourceFile & "_" & getDateFormat()
  strResult = strTemp
  
  dim fileNameIterName
  filenameIterator = 0
  while g_FSO.fileExists(strResult & fileNameIterName & astrExtension)
    filenameIterator = fileIterator + 1
    fileNameIterName = "_" & filenameIterator
  wend
  if (filenameIterator) then
    strResult = strTemp & fileNameIterName
  end if

  generateName = strResult & astrExtension
  
end function
'------------------------------------------------------------------------------

' ����� � ���� � � ����������� ����
sub logOutMsgBox(astrMsg)
  on error resume next  
  msgbox astrMsg
  logOut astrMsg
end sub
'------------------------------------------------------------------------------

' ����� � ����
sub logOut(astrMsg)
  on error resume next  
  WScript.StdOut.WriteLine astrMsg
  logOutToFile astrMsg
end sub
'------------------------------------------------------------------------------

' �������� �������� ����� ��� ����������� �������� ����������
' astrPath - ���� � �����
function createFile(astrPath, astrSubFolder, astrExtension)
  strFinalPath = astrPath & astrSubFolder
  if not g_FSO.FolderExists(strFinalPath) then
    g_FSO.CreateFolder strFinalPath
  end if
  createFile = generateName(strFinalPath, c_strLogExtension)
end function
'------------------------------------------------------------------------------

' ����� ������ �� ������� ����
sub logOutToFile(astr)
  dim ts
  set ts = g_FSO.OpenTextFile(g_LogName, 8, true, 0)
  ts.WriteLine(Now() & " >> " & astr)
  ts.close()
end sub
'------------------------------------------------------------------------------