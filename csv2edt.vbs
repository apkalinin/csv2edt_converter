' конвертация первого блока данных из csv в edt
const c_strScriptVer = "1.2"

' 1.2:
' - имя очередного блока берется из второй ячейки строки, содержащей c_strNewBlock
'   поменялся формат исходных файлов на номер;название;тип строки;значения...
'   (см __test03.csv)
'------------------------------------------------------------------------------

' названия соответствующих подпапок
const c_strLogsSubFolder = "logs"
const c_strSourceSubFolder = "source\"
'const c_strSourceSubFolder = ""
const c_strResultsSubFolder = "results"

' путь к исходным данным для отладки
c_strSourceFile = "__test03"

' расширения файлов
const c_strLogExtension = ".log"
const c_strResultsExtension = ".edt"
const c_strSourceExtension = ".csv"

' разделительный символ csv файла
' см c_WinUserSetting_ListSeparator
'const c_CSVSplitSymbol = ";"

' формат дробных чисел
const c_EdtDecimalSymbol = "."

' кодовое слово для обозначения нового блока
const c_strNewBlock = "strain(%)"

' размер одной числовой записи
const c_nOutputWidth = 80
const c_nElementLength = 10
const c_nNumberLength = 5
c_nElementsInOneRow = c_nOutputWidth / c_nElementLength
c_nNumbersInOneRow = c_nOutputWidth / c_nNumberLength

' размер информационного поля о количестве элементов
const c_nHeaderLenth_NumberOfElements = 5

' параметры формирования заголоков
const c_nBlockHeader01_FullLength = 16
const c_nBlockHeader01_LeadingSpaces = 4
const c_nBlockHeader02_FullLength = 59	' вот почему-то столько отводит места программа

' текст заголовков
const c_strBlockHeader01_01 = "Mat"
const c_strBlockHeader01_02 = "_Busher"

const c_strBlockHeader02_01 = "Modulus values for Material No. "
const c_strBlockHeader02_02 = "Damping values for Material No. "

const c_strGlogalHeader01 = "SHAKE2000 EDT File - SI"
const c_strGlogalHeader02 = "Option 1 - Dynamic material properties"
const c_strGlogalHeader03 = "    1"

' до какого разряда производить округление
const c_nPlaceToRound = 5

const c_nBlockTypeModulus = 0
const c_nBlockTypeDamping = 1

' глобальные переменные 
set g_FSO = CreateObject("Scripting.FileSystemObject")
set g_Shell = WScript.CreateObject("WScript.Shell")
const c_strLocalSettingRegPath = "HKCU\Control Panel\International\"

dim c_WinUserSetting_DecimalSymbol
dim c_WinUserSetting_ListSeparator

' внешний файл логов процесса сборки
dim g_LogName

' итоговый массив обработанных строк, 
' придется делать глобальным для избежания блокировок
dim g_arrResultArray()
g_nResultArrayIndex = 3
 '------------------------------------------------------------------------------

' вызов главной функции
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
  
  ' загружаем исходные данные
  if (astrSourcePath <> "") then
    strSourceFilePath = astrSourcePath
    c_strSourceFile = g_FSO.GetBaseName(astrSourcePath)
  else 
    strSourceFilePath = getLocalPath() & c_strSourceSubFolder & c_strSourceFile & c_strSourceExtension
  end if
  
	' создаем файл логов
  g_LogName = createFile(getLocalPath(), c_strLogsSubFolder, c_strLogExtension)
  logOut "Starting csv2edt conversion, script version: " & c_strScriptVer
	logOut "created log file: " & g_LogName
 
  logOut "trying to load source file: " & strSourceFilePath
  if not g_FSO.fileExists(strSourceFilePath) then
		logOut "couldn't find source file, exiting script"
		startConversion = -1
		exit function	
  end if
  
	' проверка что в файле есть данные
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
  ' точно знаем эти строки 
  g_arrResultArray(0) = c_strGlogalHeader01
  g_arrResultArray(1) = c_strGlogalHeader02
  g_arrResultArray(2) = c_strGlogalHeader03
  ' 3 заполняем в конце, после того как станет известно общее количество элементов
  
	' строки текущего блока данных
	dim arrCurrentBlock
  arrCurrentBlock = Array()
	
	' элементы строки строки
	dim arrCurrentLine
	arrCurrentLine = array()
	
	dim nSourceLineNumber
	dim nBlockLineNumber
	dim nResultLinesNumber
	
	dim fNewBlockFlag
	fNewBlockFlag = false
	
	dim nCurrentLineElementsNumber
	dim nMaterialNumber
  dim strMaterialName
	
	dim nBlockType ' 0 - modulus, 1 - damping
	nBlockType = c_nBlockTypeModulus
	
  logOut "initializing complete"
  logOut ""
  while not instructionsFile.AtEndOfStream
    nSourceLineNumber = nSourceLineNumber + 1
		
    ' идея с блочной структурой при детальном рассмотрении не кажется столь удобной
    ' переделываю на один массив
    
    startConversion = parseLine(_
			nSourceLineNumber,_
			instructionsFile.ReadLine(),_
			arrCurrentLine,_
			nCurrentLineElementsNumber,_
			nMaterialNumber,_
      strMaterialName,_
			fNewBlockFlag) 
    
    if (fNewBlockFlag) then   
      
      extend = extendResultArray()
      g_arrResultArray(g_nResultArrayIndex) = createBlockHeader(_
        nBlockType, _
        nMaterialNumber, _
        strMaterialName, _
        nCurrentLineElementsNumber)
      
      logOut ""
      logOut "block header: " & vbNewLine & g_arrResultArray(g_nResultArrayIndex)

      ' чередуем типы блоков так, ибо в csv не указаны типы
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
      
  ' 2 первых пропускаем
  ' 26.11 3 пропускаем 3
  for i = 3 to anCurrentLineElementsNumber + 1
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

function createBlockHeader(anBlockType, anMaterialNumber, astrMaterialName, anCurrentLineElementsNumber)
	
	dim strNumberOfElems
	strNumberOfElems = addLeadingSpaces(_
		anCurrentLineElementsNumber,_
		c_nHeaderLenth_NumberOfElements,_
		len(anCurrentLineElementsNumber))
	
	dim strHeader01
	strHeader01 = string(c_nBlockHeader01_LeadingSpaces, " ") &_
		c_strBlockHeader01_01 &_ 
		anMaterialNumber &_ 
		"_" & astrMaterialName
	
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
  astrMaterialName,_
	afNewBlockFlag)
	
  parseLine = -1
	
	arrCurrentLine = Split(aLine, c_WinUserSetting_ListSeparator, -1, 1)
	nSourceLineElementsNumber = ubound(arrCurrentLine)
	
  logOut ""
	logOut "parsing line: " & anSourceLineNumber & _
    "; number of elements is source line: " & nSourceLineElementsNumber + 1
	
	if ((nSourceLineElementsNumber < 2) or (nSourceLineElementsNumber > 22)) then  
    logOutMsgBox "Bad line: " & sourceLineNumber & _
      " - inappropriate number of elements: " & nSourceLineElementsNumber
    exit function
  end if
	
	if (arrCurrentLine(2) = c_strNewBlock) then
		afNewBlockFlag = true
		anMaterialNumber = arrCurrentLine(0)
    astrMaterialName = arrCurrentLine(1)
    
		
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
		
    ' 26.11 пропускаем первые 3 служебных элемента
		if i < 3 then
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
				
				' 0 слева от точки не пишется
				if (left(arrCurrentLine(i), 1) = "0") then
					' обработаем полный 0
          if (arrCurrentLine(i) <> "0") then
            nNumberOfCharactersToCopy = nNumberOfCharactersToCopy - 1
          end if
				end if
				
				' добавляем лидирующие пробелы      
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

' округление по c_nPlaceToRound разряду в зависимости от
' (c_nPlaceToRound + 1) разряда
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

' функция определения пути к скрипту
function getLocalPath()
  getLocalPath = Mid(WScript.ScriptFullName, 1, _
    Len(WScript.ScriptFullName) - Len(WScript.ScriptName))
end function
'------------------------------------------------------------------------------

' Дать имя даты текущее
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

' Создание уникального имени файла по дате
' astrPath - путь к файлу
' astrExtension - расширение создаваемого файла
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

' вывод в логи и в всплавающее окно
sub logOutMsgBox(astrMsg)
  on error resume next  
  msgbox astrMsg
  logOut astrMsg
end sub
'------------------------------------------------------------------------------

' вывод в логи
sub logOut(astrMsg)
  on error resume next  
  WScript.StdOut.WriteLine astrMsg
  logOutToFile astrMsg
end sub
'------------------------------------------------------------------------------

' создание внешнего файла для логирования процесса автосборки
' astrPath - путь к файлу
function createFile(astrPath, astrSubFolder, astrExtension)
  strFinalPath = astrPath & astrSubFolder
  if not g_FSO.FolderExists(strFinalPath) then
    g_FSO.CreateFolder strFinalPath
  end if
  createFile = generateName(strFinalPath, c_strLogExtension)
end function
'------------------------------------------------------------------------------

' вывод текста во внешний файл
sub logOutToFile(astr)
  dim ts
  set ts = g_FSO.OpenTextFile(g_LogName, 8, true, 0)
  ts.WriteLine(Now() & " >> " & astr)
  ts.close()
end sub
'------------------------------------------------------------------------------