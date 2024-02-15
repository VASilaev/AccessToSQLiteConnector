'https://tablacus.github.io/scriptcontrol_en.html - x64

dim objDynamicWrapperX, objPrintProvider, dbgLevel
set objPrintProvider = Nothing

sub Err_Raise (message)
'Возбуждает исключение с указанным текстом
  err.raise 1000, , message
end sub

function GetPrintProvider ()
  set GetPrintProvider = objPrintProvider
end function

Function SetPrintProvider(pEngine)
'Сохрняет ссылку на провайдера отладочных сообщений
'#param pEngine - Объект с методом Output и единственным параметром
  set objPrintProvider = pEngine
  if not objPrintProvider is nothing then dbg "Лог начат " & now(), 0
end function

function GetFileLogger(spFileName)
'Возвращает провайдера отладочных сообщений с выводом сообщений в файл
'#param spFileName - имя файла с логом
  dim fl
  set fl = new FileLogger
  set fl.TextStream = CreateObject("Scripting.FileSystemObject").OpenTextFile(spFileName, 2, True)
  set GetFileLogger=fl
end function

Sub dbg(spDbgText, npLevel)
'Выводит отладочное сообщение через провайдер
'#param spDbgText - Теккстовое сообщение
'#param npLevel - Уровень сдвига. Если положительное то следующие сообщения будут выведены с добавление пробелов. Если отрицательное, то начиная с текущего сообщения сдвигается влево
  
  if not objPrintProvider is Nothing then
    
    if npLevel < 0 then dbgLevel = dbgLevel + npLevel
    if dbgLevel < 0 then dbgLevel = 0
    Dim Line, sPrefix
    sPrefix = Space(dbgLevel)
    on error resume next
    For each Line in split(spDbgText, vbCrLf)
      objPrintProvider.Output sPrefix & Line
      if err then 
        on error goto 0
        set objPrintProvider = nothing
        exit sub
      end if
    Next
    on error goto 0
    if npLevel >0 then dbgLevel = dbgLevel + npLevel
  end if
end Sub


Function InitDLL (spDllPath)
'Инициализация динамически подключаемой библиотеки
'#param spDllPath - Абсолютный путь до sqlite3.dll

  objDynamicWrapperX.Register spDllPath, "sqlite3_errmsg", "i=p", "r=p"
  objDynamicWrapperX.Register spDllPath, "sqlite3_prepare16_v2", "i=pwlPp", "r=l"
  objDynamicWrapperX.Register spDllPath, "sqlite3_step", "i=p", "r=l"
  objDynamicWrapperX.Register spDllPath, "sqlite3_reset", "i=p", "r=l"
  objDynamicWrapperX.Register spDllPath, "sqlite3_finalize", "i=p", "r=l"
  objDynamicWrapperX.Register spDllPath, "sqlite3_open16", "i=wP", "r=l"
  objDynamicWrapperX.Register spDllPath, "sqlite3_close", "i=p", "r=l"
  objDynamicWrapperX.Register spDllPath, "sqlite3_column_count", "i=p", "r=l"

  objDynamicWrapperX.Register spDllPath, "sqlite3_column_type", "i=pl", "r=l"
  objDynamicWrapperX.Register spDllPath, "sqlite3_column_name16", "i=pl", "r=w"
  objDynamicWrapperX.Register spDllPath, "sqlite3_column_double", "i=pl", "r=d"
  objDynamicWrapperX.Register spDllPath, "sqlite3_column_int", "i=pl", "r=l"
  objDynamicWrapperX.Register spDllPath, "sqlite3_column_text16", "i=pl", "r=w"

  objDynamicWrapperX.Register spDllPath, "sqlite3_value_type", "i=p", "r=l"
  objDynamicWrapperX.Register spDllPath, "sqlite3_value_double", "i=p", "r=d"
  objDynamicWrapperX.Register spDllPath, "sqlite3_value_int", "i=p", "r=l"
  objDynamicWrapperX.Register spDllPath, "sqlite3_value_text16", "i=p", "r=w"

  objDynamicWrapperX.Register spDllPath, "sqlite3_last_insert_rowid", "i=p", "r=m"
  objDynamicWrapperX.Register spDllPath, "sqlite3_libversion", "r=s"
  objDynamicWrapperX.Register spDllPath, "sqlite3_bind_parameter_index", "i=ps", "r=l"
  objDynamicWrapperX.Register spDllPath, "sqlite3_bind_double", "i=pld", "r=l"
  objDynamicWrapperX.Register spDllPath, "sqlite3_bind_int", "i=pll", "r=l"
  objDynamicWrapperX.Register spDllPath, "sqlite3_bind_null", "i=pl", "r=l"
  objDynamicWrapperX.Register spDllPath, "sqlite3_bind_text16", "i=plwlp", "r=l"

  'Работа с памятью 
  objDynamicWrapperX.Register spDllPath, "sqlite3_malloc", "i=l", "r=p"
  objDynamicWrapperX.Register spDllPath, "sqlite3_free", "i=p"

  'Набор функций для возврата результата
  objDynamicWrapperX.Register spDllPath, "sqlite3_result_double", "i=pd"
  objDynamicWrapperX.Register spDllPath, "sqlite3_result_int", "i=pl"
  objDynamicWrapperX.Register spDllPath, "sqlite3_result_text16", "i=pwlp"
  objDynamicWrapperX.Register spDllPath, "sqlite3_result_null", "i=p"
  objDynamicWrapperX.Register spDllPath, "sqlite3_result_error16", "i=pwl"

  '

  objDynamicWrapperX.Register spDllPath, "sqlite3_declare_vtab", "i=pp", "r=l"
  objDynamicWrapperX.Register spDllPath, "sqlite3_create_module_v2", "i=ppppp", "r=l"

  set InitDLL = objDynamicWrapperX
end function

Function OpenDataBase(sDataBaseName, oWrapper, oApplication)
'Открывает базу данных SQLite и возвращает объект соединения
'#param sDataBaseName - Абсолютный путь до файла с БД. Передайте пусто, для создания БД в памяти
'#param oWrapper - Ссылка на класс SQLiteEngine. Пока жив хотя бы один экземпляр объект соединения класс SQLiteEngine не будет уничтожен
'#param oApplication - Ссылка на Access.Application. Нужен для созданитя объектов подключения для присоединенных БД

  dim oDB 
  set oDB = new SQLite_Connection
  oDB.Open sDataBaseName
  set odb.Wrapper = oWrapper
  set odb.Application =  oApplication
  set OpenDataBase = oDB
end function 




sub addInArray(byref spArray, byref pItem)
'Добавляет элемент в массив на месте
'#param spArray - Ссылка на массив в который нужно добавить элемент. Если пераметр пуст то создается пустой массив
'#param pItem - Добавляемый элемент
  if not isArray(spArray) then spArray = array()
  redim preserve spArray(ubound(spArray) + 1)
  spArray(ubound(spArray)) = pItem
end sub



function ShiftRightN(s, nCount)
'Сдвигает каждую строчку текста на указанное число пробелов
'#param s - Текст
'#param nCount - Число пробелов которое нужно включить перед строкой

  dim l, sSpace
  sSpace = space(nCount * 2)
  for each l in split(s, vbcrlf)
    if ShiftRightN <> "" then ShiftRightN = ShiftRightN & vbcrlf
    if l <> "" then ShiftRightN = ShiftRightN & sSpace & l
  next
end function



function dumpVar(byref a)
'Выводит содержимое переменной
'#param a - Исследуемая переменная

  'on Error resume next
  dim i, v

  v = VarType(a)
  if isArray(a) then
    if uBound(a) >= lBound(a) then
      dumpVar = dumpVar & "array("
      for i = lbound(a) to ubound(a)
       if i > lBound(a) then
        v = "," & vbCrlf
       else
        v = vbCrlf
       end if
       dumpVar = dumpVar & v & ShiftRightN(dumpVar(a(i)), 1)
      next
      dumpVar = dumpVar & vbCrlf & ")"
    else
      dumpVar = "array()"
    end if

  elseif lcase(TypeName(a)) = "dictionary" then
   dumpVar = "Dic {" & vbCrLf
   for each v in a.keys
    dumpVar = dumpVar & "  " & v & ": " & ShiftRightN(dumpVar(a(v)), 1) & vbCrLf
   next
   dumpVar = dumpVar & "}"
  elseif v = 2 or v = 3 or v = 4 or v = 5 then
    dumpVar = a & ""
  elseif isObject(a)  then
    dumpVar = lcase(TypeName(a))
    if dumpVar = "sqlite3_accessvirtualindex" then 
      dumpVar = a.dump
    else 
      dumpVar = "object(" & TypeName(a) &")"
    end if
  else 'Все остальные как строки
    if isNull(a) then
     dumpVar = "[NULL]"  
    elseif isEmpty(a) then
     dumpVar = "[Empty]"    
    else
      dumpVar = """" & replace(a, """", """""") & """"
      if instr(dumpVar,vbCr) > 0 or instr(dumpVar,vbLf) > 0 then
        dumpVar = replace(dumpVar,vbCr,chr(164))
        dumpVar = replace(dumpVar,vblf,chr(182))
        dumpVar = "sf(" & dumpVar & ")"
      end if
    end if

  end if
end function

Public Sub DumpMem(ByVal pAddress , ByVal nLength)
'Собирает дамп памяти
'#param pAddress - Стартовый адрес
'#param nLength - Размер дампа

  Dim b , s, l 
  
  if not objPrintProvider is Nothing then dbg "Dump memmory from " & Hex(pAddress), 0
  if not objPrintProvider is Nothing then dbg "00 01 02 03 04 05 06 07 08 09 0A 0B 0C 0D 0E 0F", 0
  
  Do While nLength > 0
    s = ""
    l = 16
    Do While nLength > 0 And l > 0
      b = objDynamicWrapperX.NumGet (pAddress, 0, "b")
      pAddress = pAddress + 1
      nLength = nLength - 1
      l = l - 1
      If s <> "" Then s = s & " "
      If b < 16 Then s = s & "0" & Hex(b) Else s = s & Hex(b)
    Loop
    if not objPrintProvider is Nothing then dbg s, 0
  Loop
  if not objPrintProvider is Nothing then dbg "", 0
End Sub

on Error resume next
set objDynamicWrapperX = CreateObject("DynamicWrapperX")
if err then 
 err.clear
 set objDynamicWrapperX = CreateObject("DynamicWrapperX.2")
 if err then 
   on Error goto 0   
   err_raise "Установите сначала DynamicWrapperX"
 end if
end if


on Error goto 0

const SQLITE_OK         = 0   ' Successful result
const SQLITE_READONLY   = 0
const SQLITE_ROW	= 100
const SQLITE_DONE 	= 101
const SQLITE_INTEGER	= 1
const SQLITE_FLOAT    	= 2
const SQLITE_TEXT     	= 3
const SQLITE_BLOB     	= 4
const SQLITE_NULL     	= 5

'const SQLITE_OK = 0
const SQLITE_Error = 1
const SQLITE_NOMEM = 7
const SQLITE_NOTFOUND = 12

Const CP_UTF8 		= 65001
Const dbAutoIncrField   = 16

const dbBigInt = 16
const dbBinary = 9
const dbBoolean = 1
const dbByte = 2
const dbChar = 18
const dbCurrency = 5
const dbDate = 8
const dbDecimal = 20
const dbDouble = 7
const dbFloat = 21
const dbGUID = 15
const dbInteger = 3
const dbLong = 4
const dbLongBinary = 11
const dbMemo = 12
const dbNumeric = 19
const dbSingle = 6
const dbText = 10
const dbTime = 22
const dbTimeStamp = 23
const dbVarBinary = 17


Const JULIANDAY_OFFSET = 2415018.5

const dbOpenTable = 1



const SQLITE_INDEX_CONSTRAINT_EQ = 2 
const SQLITE_INDEX_CONSTRAINT_GT = 4 
const SQLITE_INDEX_CONSTRAINT_LE = 8 
const SQLITE_INDEX_CONSTRAINT_LT = 16 
const SQLITE_INDEX_CONSTRAINT_GE = 32 
const SQLITE_INDEX_CONSTRAINT_MATCH = 64 
const SQLITE_INDEX_CONSTRAINT_LIKE = 65 
const SQLITE_INDEX_CONSTRAINT_GLOB = 66 
const SQLITE_INDEX_CONSTRAINT_REGEXP = 67 
const SQLITE_INDEX_CONSTRAINT_NE = 68 
const SQLITE_INDEX_CONSTRAINT_ISNOT = 69 
const SQLITE_INDEX_CONSTRAINT_ISNOTNULL = 70 
const SQLITE_INDEX_CONSTRAINT_ISNULL = 71 
const SQLITE_INDEX_CONSTRAINT_IS = 72 
const SQLITE_INDEX_CONSTRAINT_LIMIT = 73 
const SQLITE_INDEX_CONSTRAINT_OFFSET = 74 
const SQLITE_INDEX_CONSTRAINT_FUNCTION = 150 
const SQLITE_INDEX_SCAN_UNIQUE = 1

Function GetCursorOject(hCursor)
'Получает объект курсора поего указателю
'#param hCursor - Указатель на курсор

  GetCursorOject = objDynamicWrapperX.numGet (hCursor, ptr_size, "p")
  set GetCursorOject = objDynamicWrapperX.ObjGet(GetCursorOject)
end function

Function SQLite3_Value(ptr)
'Извлекает значение из структуры sqlite3_value 
'#param pPtr - Указатель на структуру sqlite3_value 

 Select Case objDynamicWrapperX.sqlite3_value_type(ptr)  
   Case SQLITE_NULL
     SQLite3_Value = NULL
   Case SQLITE_INTEGER
     SQLite3_Value = objDynamicWrapperX.sqlite3_value_int(ptr)        
   Case SQLITE_FLOAT
     SQLite3_Value = objDynamicWrapperX.sqlite3_value_double(ptr)
   Case SQLITE_TEXT
     SQLite3_Value = objDynamicWrapperX.sqlite3_value_text16(ptr)
   Case SQLITE_BLOB
     SQLite3_Value = "BLOB НЕ РЕАЛИЗОВАН"
   Case Else
     SQLite3_Value = "НЕ УАДЛОСЬ ОПРЕДЕЛИТЬ ТИП"
 End Select 

  if not objPrintProvider is Nothing then dbg " - SQLite3_Value Type = [" & typename(SQLite3_Value) & "], Value = [" & SQLite3_Value & "]", 0  
end function

Function GetVArr(pPtr, nIdx)
'Извлекает элемент из массива указателей
'#param pPtr - Указатель на начало массива
'#param nIdx - Индекс элемента
 GetVArr = objDynamicWrapperX.numGet (pPtr, nIdx * ptr_size, "p")
end function


class SQLite3_AccessVirtualTable
'Вспомогательный клас, содержит все необходимое для создания RecordSet
  public SelfLock, sTableName, oConnection, aFields, aIndexes, nEstimatedRows, nROWIDColumn, CanSeek, TableRecordSet, primaryKeyIndex, hVTab
  public  sub setError (sError)
    dim ptr 
    ptr = objDynamicWrapperX.numGet (hVTab, 2 * ptr_size, "p")
    if ptr <> 0 then objDynamicWrapperX.sqlite3_free(ptr)
    objDynamicWrapperX.numPut SQLite3mPrint(sError) ,hVTab, 2 * ptr_size, "p"
  end sub
end class 

class SQLite3_AccessVirtualIndex
'Вспомогательный класс, описывающий индекс
  public Name, Unique, Count, FullMask, Cost, EstimatedRows, columnUsage, ColumnToConstrain
  public function dump()
    dump = replace("{\n Name:"""& Name & """,\n Unique:" & Unique & ",\n Count:" & Count & ",\n FullMask:" & FullMask & ",\n Cost:" & Cost &  ",\n ColumnToConstrain:" & dumpvar(ColumnToConstrain) & "\n}","\n", vbcrlf)
  end function 
end class 


class SQLite3_AccessVTCursor
'Вспомогательный класс, текущего курсора
  public SelfLock, oRecordSet, oVirtualTable, FindByPK
end class 




class FileLogger
  public TextStream
  public sub Output(pText)
   TextStream.WriteLine pText
  end sub
  Private Sub Class_Terminate() 
    if isObject(TextStream) then
      if not TextStream is nothing then   
        TextStream.close()
        set TextStream = Nothing
      end if
    end if
  End Sub 
end class 

function SQLite3mPrint(sText)
    SQLite3mPrint = objDynamicWrapperX.sqlite3_malloc((len(sText) + 1) * 2)
    objDynamicWrapperX.StrPut sText, SQLite3mPrint, 0 , "UTF-8"
end function 




Function SQLite3AccessInit (byVal hDB, byVal pAux, byVal argc, byval argv, byval ppVtab, ByVal pzErr)
'Создание виртуальной таблицы

  dim oConnection, oAccessConnection, oTableDef
  if not objPrintProvider is Nothing then dbg ">SQLite3AccessInit hDB = [" & hDB & "], pAux = [" & pAux & "], argc = [" & argc & "], argv = [" & argv & "], ppVtab = [" & ppVtab & "], pzErr = [" & pzErr & "]", + 2
  set oConnection = objDynamicWrapperX.ObjGet(pAux)
  'msgbox "oAccess.Name = [" & oAccess.name & "]"

  if argc < 4 then 
    objDynamicWrapperX.numPut SQLite3mPrint("Укажите имя присоединяемой таблицы"), pzErr, 0, "p"
    SQLite3AccessInit = SQLITE_Error     
  else   
    dim sTableName
    sTableName = objDynamicWrapperX.strGet(objDynamicWrapperX.numGet (argv, 3 * ptr_size, "p"),0,"UTF-8")

    if not objPrintProvider is Nothing then dbg "Table [" & sTableName & "]", 0

    set oAccessConnection = oConnection.GetAccessDB("")
    if oAccessConnection is Nothing then 
      if not objPrintProvider is Nothing then dbg "No current DB", 0
      objDynamicWrapperX.numPut SQLite3mPrint("Не найдена база данных [По умолчанию]"), pzErr, 0, "p"
      SQLite3AccessInit = SQLITE_Error    
    else 
      set oTableDef = oAccessConnection.TableDefs(sTableName)     

      dim vTab, vFieldDef, sSchema
      set vTab = new SQLite3_AccessVirtualTable

'todo открывать новое подключение для внешних таблиц 
      vTab.CanSeek = true

      if oTableDef.Connect <> "" then 
        vTab.CanSeek = false
      end if 

      set vTab.TableRecordSet = Nothing



      set vTab.oConnection = oAccessConnection

      vTab.nEstimatedRows = oTableDef.RecordCount
      if vTab.nEstimatedRows < 0 then 
        vTab.nEstimatedRows = 10000000.0
      end if

      vTab.sTableName = sTableName


      if not objPrintProvider is Nothing then dbg "Fields count [" & oTableDef.Fields.Count & "]", 0

      vFieldDef = array()   
      if oTableDef.Fields.Count > 0 then      
        redim vFieldDef(oTableDef.Fields.Count - 1)
        vTab.aFields = vFieldDef
        
        i = 0
        For each item in oTableDef.Fields
          vTab.aFields(i) = array(lcase(item.Name), Empty, item.Attributes)
          i = i + 1
        next

      end if

      if not objPrintProvider is Nothing then dbg "Index count [" & oTableDef.Indexes.Count & "]", 0

      vFieldDef = array()  
      if oTableDef.Indexes.Count > 0 then       
        redim vFieldDef(oTableDef.Indexes.Count - 1)
        vTab.aIndexes = vFieldDef
        i = 0
        For each index in oTableDef.Indexes
          dim ind, ColumnToConstrain
          set ind = new SQLite3_AccessVirtualIndex
          ind.Name = index.Name
          ind.Unique = index.Unique
          ind.Count = index.Fields.count
          ind.FullMask = 2 ^ (ind.Count) - 1
          ColumnToConstrain = array()
          redim ColumnToConstrain(ind.Count - 1)
          ind.ColumnToConstrain = ColumnToConstrain 

          if ind.Unique then 
            ind.Cost = 1.0
            if vTab.nEstimatedRows = 10000000.0 then vTab.nEstimatedRows = index.DistinctCount
          else 
            ind.Cost = vTab.nEstimatedRows / index.DistinctCount
          end if
          set vTab.aIndexes(i) = ind

          x = 0      
          for each item in index.Fields         
            FieldName = lCase(item.Name)
            cf = ubound(vTab.aFields)
            for j = 0 to cf
              if vTab.aFields(j)(0) = FieldName then
                dim f
                f = vTab.aFields
                addinarray f(j)(1), array(ind, 2 ^ x, x)
                vTab.aFields = f

                if index.primary and ind.Count = 1 then 
                  vTab.nROWIDColumn = vTab.aFields(j)(0)
                  set vTab.primaryKeyIndex = ind
                end if 
                exit for
              end if
            next
            x = x + 1
          next 
          i = i + 1
        next
      end if

      if not objPrintProvider is Nothing then dbg "nEstimatedRows [" & vTab.nEstimatedRows & "]", 0

      'индексы нужно отсортировать от самых ценных к самым бесполезным 


      'dbg "Fields",0 
      'dbg dumpVar(vTab.aFields) & ubound(vTab.aFields),0   
      'dbg "Indexes",0  
      'dbg dumpVar(vTab.aIndexes) & ubound(vTab.aIndexes),0    

      'dbg "Collect Schema", 0

      

      For Each vFieldDef In oTableDef.Fields
        If sSchema <> "" Then sSchema = sSchema & ","
        sSchema = sSchema & vFieldDef.Name & " "
        
  
        If (vFieldDef.Attributes And dbAutoIncrField) = dbAutoIncrField Then
          sSchema = sSchema & "INTEGER" 'NOT NULL PRIMARY KEY            
        Else
          Select Case vFieldDef.Type
            Case dbBigInt, dbBoolean, dbByte, dbInteger, dbLong
                sSchema = sSchema & "INTEGER"
            Case dbCurrency, dbDecimal, dbNumeric
                sSchema = sSchema & "DECIMAL"
            Case dbDate, dbTimeStamp
                sSchema = sSchema & "DATETIME"
            Case dbTime
                sSchema = sSchema & "TIME"
            Case dbDouble, dbFloat, dbSingle
                sSchema = sSchema & "DOUBLE"
            Case dbGUID
                sSchema = sSchema & "GUID"
            Case dbLongBinary, dbVarBinary
                sSchema = sSchema & "BLOB"
            Case dbMemo:
                sSchema = sSchema & "MEMO"
            Case Else
                sSchema = sSchema & "TEXT"
          End Select
        End If
      Next
      sSchema = "CREATE TABLE x(" & sSchema & ")" 
      if isEmpty(vTab.nROWIDColumn) then sSchema = sSchema & "WITHOUT ROWID"         
      
      if not objPrintProvider is Nothing then dbg "Table Schema [" & sSchema & "]", 0

      dim rc
      rc = objDynamicWrapperX.sqlite3_declare_vtab (hDB, objDynamicWrapperX.StrPtr(sSchema,"UTF-8"))
      sSchema = empty
      
      if SQLITE_OK = 0 then 
        

        rc = objDynamicWrapperX.sqlite3_malloc(ptr_size * 4)

        if rc = 0 then 
          SQLite3AccessInit = SQLITE_NOMEM
        else 
          set vTab.SelfLock = vTab
          vTab.hVTab = rc
          if not objPrintProvider is Nothing then dbg "vTab obj ref [" & objDynamicWrapperX.ObjPtr(vTab) & "] struct address [" & rc & "]", 0
          objDynamicWrapperX.numPut rc, ppVtab, 0, "p"
          objDynamicWrapperX.numPut objDynamicWrapperX.ObjPtr(vTab), rc, 3 * ptr_size, "p"
          SQLite3AccessInit = SQLITE_OK            
        end if
        
      else 
        objDynamicWrapperX.numPut SQLite3mPrint("Не удалось зарегистрировать схему таблицы " & sSchema), pzErr, 0, "p"
        SQLite3AccessInit = SQLITE_Error            
      end if   
    end if
  end if
  if not objPrintProvider is Nothing then dbg "<SQLite3AccessInit " & SQLite3AccessInit, -2
end function

Function SQLite3AccessConnect (byVal hDB, byVal pAux, byVal argc, byval argv, byval ppVtab, ByVal pzErr)
'Virtual table method [xConnect](https://www.sqlite.org/vtab.html#xconnect)
  if not objPrintProvider is Nothing then dbg ">SQLite3AccessConnect " , 2
  dim sError
  on Error resume next 
  SQLite3AccessConnect = SQLite3AccessInit(hDB,pAux,argc,argv,ppVtab,pzErr)
  if err then sError = err.Description
  on Error Goto 0
  if not isEmpty(sError) then 
    objDynamicWrapperX.numPut SQLite3mPrint(sError), pzErr, 0, "p"
    SQLite3AccessConnect = SQLITE_Error       
  end if
  if not objPrintProvider is Nothing then dbg "<SQLite3AccessConnect " & SQLite3AccessConnect, -2
end function


Function SQLite3AccessCreate (byVal hDB, byVal pAux, byVal argc, byval argv, byval ppVtab, ByVal pzErr)
'Virtual table method [xCreate](https://www.sqlite.org/vtab.html#xcreate)
  if not objPrintProvider is Nothing then dbg ">SQLite3AccessConnect " , 2
  dim sError
  on Error resume next 
  SQLite3AccessCreate = SQLite3AccessInit(hDB,pAux,argc,argv,ppVtab,pzErr)
  if err then sError = err.Description
  on Error Goto 0
  if not isEmpty(sError) then 
    objDynamicWrapperX.numPut SQLite3mPrint(sError), pzErr, 0, "p"
    SQLite3AccessCreate = SQLITE_Error       
  end if
  if not objPrintProvider is Nothing then dbg "<SQLite3AccessCreate " & SQLite3AccessCreate, -2
end function


Function SQLite3AccessBestIndex (byVal hVTab, byVal hIdxInfo)
'Virtual table method [xBestIndex](https://www.sqlite.org/vtab.html#xbestindex)
  dim vTab, bFullScan
  bFullScan = true 
  vTab = objDynamicWrapperX.numGet (hVTab, 3 * ptr_size, "p")
  if not objPrintProvider is Nothing then dbg ">SQLite3AccessBestIndex hVTab = [" & hVTab & "] ovTab = [" & vTab & "] hIdxInfo =[" & hIdxInfo & "]", +2
  set vTab = objDynamicWrapperX.ObjGet(vTab)

  'DumpMem  hIdxInfo, 16 * ptr_size
  
  dim nConstraint, aConstraint, aConstraintUsage, CurrentCost
  'Число ограничений
  nConstraint = objDynamicWrapperX.numGet (hIdxInfo, 0, "l")
  'Описание ограничения
  aConstraint = objDynamicWrapperX.numGet (hIdxInfo, ptr_size, "p")
  'Сюда будем писать какие ограничения нам нужны
  aConstraintUsage = objDynamicWrapperX.numGet (hIdxInfo, ptr_size * 4, "p")
 
  CurrentCost = vTab.nEstimatedRows
  if CurrentCost < 0 then CurrentCost = 1000000 

  dim indexUasage, indexIdx, sWhere, idxNum
  indexUsage = array()
  redim indexUsage(ubound(vTab.aIndexes))
  for each indexIdx in vTab.aIndexes
    indexIdx.columnUsage = 0
  next 

  idxNum = 0

  if not objPrintProvider is Nothing then dbg "nConstraint = [" & nConstraint & "] aConstraint = [" & aConstraint & "]" , 0

  if nConstraint > 0 then 
    dim idxConstraint, Usage, nConst
    nConst = 0
    
    for idxConstraint  = 0 TO  nConstraint - 1
      op = objDynamicWrapperX.numGet (aConstraint, idxConstraint * 3 * ptr_size + ptr_size, "b") 'Op
      column = objDynamicWrapperX.numGet (aConstraint, idxConstraint * 3 * ptr_size , "l") 'Номер столбца
      if objDynamicWrapperX.numGet (aConstraint, idxConstraint * 3 * ptr_size + ptr_size + 1, "b") = 1 then 'usage = 0 
        'ограничение можем обработать только если usage = true

        if not objPrintProvider is Nothing then dbg idxConstraint & ": column [" & (vTab.aFields(column)(0)) & "] op = [" & op & "]" , 0
        
        if op = SQLITE_INDEX_CONSTRAINT_EQ or op = SQLITE_INDEX_CONSTRAINT_GT or op = SQLITE_INDEX_CONSTRAINT_LE or _
           op = SQLITE_INDEX_CONSTRAINT_LT or op = SQLITE_INDEX_CONSTRAINT_GE or op = SQLITE_INDEX_CONSTRAINT_LIKE or _ 
           op = SQLITE_INDEX_CONSTRAINT_NE or op = SQLITE_INDEX_CONSTRAINT_ISNOTNULL or op = SQLITE_INDEX_CONSTRAINT_ISNULL then 
 
           if sWhere <> "" then sWhere = sWhere & " and "  
           sWhere =  sWhere & "[" & vTab.aFields(column)(0) & "]"
           select case op
             case SQLITE_INDEX_CONSTRAINT_EQ:        sWhere = sWhere & " = "
             case SQLITE_INDEX_CONSTRAINT_GT:        sWhere = sWhere & " > "
             case SQLITE_INDEX_CONSTRAINT_LE:        sWhere = sWhere & " <= "
             case SQLITE_INDEX_CONSTRAINT_LT:        sWhere = sWhere & " < "
             case SQLITE_INDEX_CONSTRAINT_GE:        sWhere = sWhere & " >= "
             case SQLITE_INDEX_CONSTRAINT_LIKE:      sWhere = sWhere & " like "
             case SQLITE_INDEX_CONSTRAINT_NE:        sWhere = sWhere & " <> "
             case SQLITE_INDEX_CONSTRAINT_ISNOTNULL: sWhere = sWhere & " is not null "
             case SQLITE_INDEX_CONSTRAINT_ISNULL:    sWhere = sWhere & " is null "
           end select

           if op <> SQLITE_INDEX_CONSTRAINT_ISNOTNULL and op <> SQLITE_INDEX_CONSTRAINT_ISNULL then 
             nConst = nConst + 1
             sWhere = sWhere & "{:" & (nConst) & "}"
             objDynamicWrapperX.numPut nConst, aConstraintUsage, idxConstraint * ptr_size * 2, "l"    'argvIndex
             objDynamicWrapperX.numPut 1, aConstraintUsage, idxConstraint * ptr_size * 2 + ptr_size, "b"  'omit
           end if

          if op = SQLITE_INDEX_CONSTRAINT_EQ then 
            bFullScan = false
            if not isEmpty(vTab.aFields(column)(1)) then              
              dim iIndex: iIndex = 1  'порядковый номер индекса 
              for each indexIdx in vTab.aFields(column)(1)
                indexIdx(0).columnUsage = indexIdx(0).columnUsage or indexIdx(1) 
                indexIdx(0).ColumnToConstrain(indexIdx(2)) = idxConstraint
                if indexIdx(0).columnUsage and indexIdx(0).FullMask = indexIdx(0).FullMask then 
                  CurrentCost = indexIdx(0).Cost                  
                  if indexIdx(0).Unique then 
                    if not objPrintProvider is Nothing then dbg "Вернет 0 или 1 строку" , 0
                    objDynamicWrapperX.numPut SQLITE_INDEX_SCAN_UNIQUE, hIdxInfo, 14 * ptr_size, "l"

                    if vTab.CanSeek then 
                      'Запоминаем номер индекса
                      idxNum = iIndex
                      'Условие больше не нужно  
                      sWhere = ""
                      for idxConstr = 0 to idxConstraint
                        objDynamicWrapperX.numPut 0, aConstraintUsage, idxConstr * ptr_size * 2, "l"    'argvIndex
                        objDynamicWrapperX.numPut 0, aConstraintUsage, idxConstr * ptr_size * 2 + ptr_size, "b"  'omit
                      next
                      for idxConstr = 0 to indexIdx(0).Count - 1
                        objDynamicWrapperX.numPut idxConstr + 1, aConstraintUsage, indexIdx(0).ColumnToConstrain(idxConstr) * ptr_size * 2, "l"    'argvIndex
                        objDynamicWrapperX.numPut 1, aConstraintUsage, indexIdx(0).ColumnToConstrain(idxConstr) * ptr_size * 2 + ptr_size, "b"  'omit
                      next
                      exit for 'Лучше уже не будет
                    end if
                  end if
                end if
                iIndex = iIndex + 1
              next 
              if idxNum <> 0 then exit for
            end if
          end if 

        elseif op = SQLITE_INDEX_CONSTRAINT_OFFSET then 

        elseif op = SQLITE_INDEX_CONSTRAINT_LIMIT then 
          
        end if
      else 
        if not objPrintProvider is Nothing then dbg idxConstraint & ": column [" & (vTab.aFields(column)(0)) & "] op = [" & op & "] no usage" , 0
      end if
    next
  end if



  if not objPrintProvider is Nothing then dbg "Cost = [" & CurrentCost & "] FullScan = [" & bFullScan & "] idxNum = [" & idxNum & "] where " & sWhere, 0




  
  if idxNum <> 0 then 
    objDynamicWrapperX.numPut idxNum, hIdxInfo, 5 * ptr_size, "l" 'idxNum
  else
    dim nOrderBy, sOrderBy, aOrderBy
    nOrderBy = objDynamicWrapperX.numGet (hIdxInfo, 2 * ptr_size, "l")
  
    if nOrderBy > 0 then 
      aOrderBy = objDynamicWrapperX.numGet (hIdxInfo, 3 * ptr_size, "l")
      sOrderBy = ""
      for nConstraint = 0 to nOrderBy - 1
        column = objDynamicWrapperX.numGet (aOrderBy, nConstraint * 2 * ptr_size , "l") ' iColumn 
  
        if sOrderBy <> "" then sOrderBy = sOrderBy & ","
        sOrderBy = sOrderBy & "[" & (vTab.aFields(column)(0)) & "]"
        if objDynamicWrapperX.numGet (aOrderBy, nConstraint * 2 * ptr_size + ptr_size, "b") = 1 then sOrderBy = sOrderBy & " desc"
      next
      if not objPrintProvider is Nothing then dbg "nOrderBy = [" & nOrderBy & "] sOrderBy " & sOrderBy, 0
     
      sWhere = sWhere & " order by " & sOrderBy
      objDynamicWrapperX.numPut 1, hIdxInfo, 8 * ptr_size, "l" 'orderByConsumed
    end if

    if sWhere <> "" then 
      objDynamicWrapperX.numPut SQLite3mPrint(sWhere), hIdxInfo, 6 * ptr_size, "p" 'idxStr
      objDynamicWrapperX.numPut 1, hIdxInfo, 7 * ptr_size, "p" 'needToFreeIdxStr
    end if
  end if


  if bFullScan = false then
    objDynamicWrapperX.numPut clng(CurrentCost), hIdxInfo, 12 * ptr_size, "q"       ' estimatedRows
    if CurrentCost > 1 then CurrentCost = CurrentCost * 1000
    objDynamicWrapperX.numPut CurrentCost * 10.0, hIdxInfo, 10 * ptr_size, "d" ' estimatedCost
  else 
    objDynamicWrapperX.numPut clng(CurrentCost), hIdxInfo, 12 * ptr_size, "q"       ' estimatedRows  
  end if

  if not objPrintProvider is Nothing then dbg "Rows = [" & objDynamicWrapperX.numGet (hIdxInfo, 12 * ptr_size, "l") & "] Cost = [" & objDynamicWrapperX.numGet (hIdxInfo, 10 * ptr_size, "d")  & "]" , 0

  'DumpMem  hIdxInfo, 16 * ptr_size

  SQLite3AccessBestIndex = SQLITE_OK

  on error goto 0
  
  if not objPrintProvider is Nothing then dbg "<SQLite3AccessBestIndex", -2
end function

Function SQLite3AccessDisconnect (byVal hVTab)
'Virtual table method [xDisconnect](https://www.sqlite.org/vtab.html#xdisconnect)
  dim vTab 
  vTab = objDynamicWrapperX.numGet (hVTab, 3 * ptr_size, "p")
  if not objPrintProvider is Nothing then dbg ">SQLite3AccessDisconnect hVTab = [" & hVTab & "] ovTab = [" & vTab & "]", +2
  set vTab = objDynamicWrapperX.ObjGet(vTab)
  set vTab.SelfLock = Nothing
  if not vtab.TableRecordSet is nothing then
     vTab.TableRecordSet.Close 
     set vTab.TableRecordSet = Nothing
  end if 
  objDynamicWrapperX.sqlite3_free(hVTab)
  SQLite3AccessDisconnect = SQLITE_OK
  if not objPrintProvider is Nothing then dbg "<SQLite3AccessDisconnect", -2
end function

Function SQLite3AccessOpen (byVal hVTab, byVal hCursor)
'Virtual table method [xOpen](https://www.sqlite.org/vtab.html#xopen)
  if not objPrintProvider is Nothing then dbg ">SQLite3AccessOpen hVTab = [" & hVTab & "] hCursor = [" & hCursor & "]", +2
  
  dim vTab 
  vTab = objDynamicWrapperX.numGet (hVTab, 3 * ptr_size, "p")
'  dbg "ovTab = [" & vTab & "]",0
  set vTab = objDynamicWrapperX.ObjGet(vTab)

  dim rc 
  rc = objDynamicWrapperX.sqlite3_malloc(ptr_size * 2)
  if rc = 0 then 
    SQLite3AccessOpen = SQLITE_NOMEM
  else 
    dim oCursor 
    set oCursor = new SQLite3_AccessVTCursor
    set oCursor.SelfLock = oCursor
    set oCursor.oVirtualTable = vTab
'    dbg "set vCursor ptr = [" & objDynamicWrapperX.ObjPtr(oCursor) & "]", 0
    objDynamicWrapperX.numPut 0, rc, 0, "p"
    objDynamicWrapperX.numPut objDynamicWrapperX.ObjPtr(oCursor), rc, ptr_size, "p"
    objDynamicWrapperX.numPut rc, hCursor, 0, "p"
    SQLite3AccessOpen = SQLITE_OK
  end if
  if not objPrintProvider is Nothing then dbg "<SQLite3AccessOpen", -2

end function


Function SQLite3AccessClose (byVal hCursor)
'Virtual table method [xClose](https://www.sqlite.org/vtab.html#xclose)

  if not objPrintProvider is Nothing then dbg ">SQLite3AccessClose hCursor = [" & hCursor & "]", +2

  dim vCursor: set vCursor = GetCursorOject(hCursor)
  set vCursor.SelfLock = Nothing
  set vCursor.oVirtualTable = nothing
  if not isEmpty(vCursor.oRecordSet) then
    'Закрываем только если открыли как Query
    if vCursor.FindByPK = 0 then vCursor.oRecordSet.close 
    set vCursor.oRecordSet = Nothing
    vCursor.oRecordSet = Empty
  end if
  objDynamicWrapperX.sqlite3_free(hCursor)
  SQLite3AccessClose = SQLITE_OK

  if not objPrintProvider is Nothing then dbg "<SQLite3AccessClose", -2
end function

Function SQLite3AccessFilter (byVal hCursor, byval idxNum, byVal pStr, byVal argc, byval argv)
'Virtual table method [xFilter](https://www.sqlite.org/vtab.html#xfilter)    

  dim vCursor, sWhere, sSQL, iIdxConstraint
  dim sError: sError = ""

  if pStr <> 0 then 
    sWhere = objDynamicWrapperX.strGet(pStr,0,"UTF-8")
  end if

  if not objPrintProvider is Nothing then dbg ">SQLite3AccessFilter hCursor = [" & hCursor & "] idxNum = [" & idxNum & "] pStr = [" & pStr & ":" & sWhere & "] argc = [" & argc & "] argv = [" & argv & "]", +2

  set vCursor = GetCursorOject(hCursor)

  if idxNum = 0 then 
    for iIdxConstraint = 0 to argc - 1
      ptr = objDynamicWrapperX.numGet (argv, iIdxConstraint * ptr_size, "p")
       
  
      on error resume next
      value = SQLite3_Value(ptr)
      if err then dbg "error " & err.description, 0   
      on error goto 0
      if varType(value)=vbString then 
        value = "'"  & replace(value, "'", "''") & "'"
      elseif isNull(value) then 
        value = "NULL"
      end if
      
      sWhere = replace(sWhere,"{:" & (iIdxConstraint + 1) & "}",value)
    next
  
    sSQL = "select * from [" & vCursor.oVirtualTable.sTableName & "]"
    if sWhere <> "" then
      if left(sWhere,10) <> " order by " then  sSQL = sSQL & " where "  
      sSQL = sSQL & sWhere
    end if 
  
    if not objPrintProvider is Nothing then dbg "SQL: " & sSQL, 0

    vCursor.FindByPK = 0
  
    on error resume next 
    set vCursor.oRecordSet = vCursor.oVirtualTable.oConnection.openrecordset(sSQL)

    if err then sError = err.description
    on error goto 0
  else 
    'Поиск по уникальному индексу
    dim Index
    set Index = vCursor.oVirtualTable.aIndexes(idxNum - 1)

    if vCursor.oVirtualTable.TableRecordSet is nothing then set vCursor.oVirtualTable.TableRecordSet = vCursor.oVirtualTable.oConnection.openrecordset(vCursor.oVirtualTable.sTableName, dbOpenTable) 
     
    set vCursor.oRecordSet = vCursor.oVirtualTable.TableRecordSet
    vCursor.FindByPK = 1
    vCursor.oRecordSet.index = Index.Name

    if not objPrintProvider is Nothing then dbg "Используется метод Seek по индексу " & Index.Name , 0

    on error resume next 
    select case Index.Count
      case 1: vCursor.oRecordSet.seek  "=", SQLite3_Value(GetVArr(argv, 0))
      case 2: vCursor.oRecordSet.seek  "=", SQLite3_Value(GetVArr(argv, 0)), SQLite3_Value(GetVArr(argv, 1))
      case 3: vCursor.oRecordSet.seek  "=", SQLite3_Value(GetVArr(argv, 0)), SQLite3_Value(GetVArr(argv, 1)), SQLite3_Value(GetVArr(argv, 2))
      case 4: vCursor.oRecordSet.seek  "=", SQLite3_Value(GetVArr(argv, 0)), SQLite3_Value(GetVArr(argv, 1)), SQLite3_Value(GetVArr(argv, 2)), SQLite3_Value(GetVArr(argv, 3))
      case 5: vCursor.oRecordSet.seek  "=", SQLite3_Value(GetVArr(argv, 0)), SQLite3_Value(GetVArr(argv, 1)), SQLite3_Value(GetVArr(argv, 2)), SQLite3_Value(GetVArr(argv, 3)), SQLite3_Value(GetVArr(argv, 4))
      case 6: vCursor.oRecordSet.seek  "=", SQLite3_Value(GetVArr(argv, 0)), SQLite3_Value(GetVArr(argv, 1)), SQLite3_Value(GetVArr(argv, 2)), SQLite3_Value(GetVArr(argv, 3)), SQLite3_Value(GetVArr(argv, 4)), SQLite3_Value(GetVArr(argv, 5))
      case 7: vCursor.oRecordSet.seek  "=", SQLite3_Value(GetVArr(argv, 0)), SQLite3_Value(GetVArr(argv, 1)), SQLite3_Value(GetVArr(argv, 2)), SQLite3_Value(GetVArr(argv, 3)), SQLite3_Value(GetVArr(argv, 4)), SQLite3_Value(GetVArr(argv, 5)), SQLite3_Value(GetVArr(argv, 6))
      case 8: vCursor.oRecordSet.seek  "=", SQLite3_Value(GetVArr(argv, 0)), SQLite3_Value(GetVArr(argv, 1)), SQLite3_Value(GetVArr(argv, 2)), SQLite3_Value(GetVArr(argv, 3)), SQLite3_Value(GetVArr(argv, 4)), SQLite3_Value(GetVArr(argv, 5)), SQLite3_Value(GetVArr(argv, 6)), SQLite3_Value(GetVArr(argv, 7))
      case 9: vCursor.oRecordSet.seek  "=", SQLite3_Value(GetVArr(argv, 0)), SQLite3_Value(GetVArr(argv, 1)), SQLite3_Value(GetVArr(argv, 2)), SQLite3_Value(GetVArr(argv, 3)), SQLite3_Value(GetVArr(argv, 4)), SQLite3_Value(GetVArr(argv, 5)), SQLite3_Value(GetVArr(argv, 6)), SQLite3_Value(GetVArr(argv, 7)), SQLite3_Value(GetVArr(argv, 8))
      case 10: vCursor.oRecordSet.seek  "=", SQLite3_Value(GetVArr(argv, 0)), SQLite3_Value(GetVArr(argv, 1)), SQLite3_Value(GetVArr(argv, 2)), SQLite3_Value(GetVArr(argv, 3)), SQLite3_Value(GetVArr(argv, 4)), SQLite3_Value(GetVArr(argv, 5)), SQLite3_Value(GetVArr(argv, 6)), SQLite3_Value(GetVArr(argv, 7)), SQLite3_Value(GetVArr(argv, 8)), SQLite3_Value(GetVArr(argv, 9))
      case 11: vCursor.oRecordSet.seek  "=", SQLite3_Value(GetVArr(argv, 0)), SQLite3_Value(GetVArr(argv, 1)), SQLite3_Value(GetVArr(argv, 2)), SQLite3_Value(GetVArr(argv, 3)), SQLite3_Value(GetVArr(argv, 4)), SQLite3_Value(GetVArr(argv, 5)), SQLite3_Value(GetVArr(argv, 6)), SQLite3_Value(GetVArr(argv, 7)), SQLite3_Value(GetVArr(argv, 8)), SQLite3_Value(GetVArr(argv, 9)), SQLite3_Value(GetVArr(argv, 10))
      case 12: vCursor.oRecordSet.seek  "=", SQLite3_Value(GetVArr(argv, 0)), SQLite3_Value(GetVArr(argv, 1)), SQLite3_Value(GetVArr(argv, 2)), SQLite3_Value(GetVArr(argv, 3)), SQLite3_Value(GetVArr(argv, 4)), SQLite3_Value(GetVArr(argv, 5)), SQLite3_Value(GetVArr(argv, 6)), SQLite3_Value(GetVArr(argv, 7)), SQLite3_Value(GetVArr(argv, 8)), SQLite3_Value(GetVArr(argv, 9)), SQLite3_Value(GetVArr(argv, 10)), SQLite3_Value(GetVArr(argv, 11))
      case 13: vCursor.oRecordSet.seek  "=", SQLite3_Value(GetVArr(argv, 0)), SQLite3_Value(GetVArr(argv, 1)), SQLite3_Value(GetVArr(argv, 2)), SQLite3_Value(GetVArr(argv, 3)), SQLite3_Value(GetVArr(argv, 4)), SQLite3_Value(GetVArr(argv, 5)), SQLite3_Value(GetVArr(argv, 6)), SQLite3_Value(GetVArr(argv, 7)), SQLite3_Value(GetVArr(argv, 8)), SQLite3_Value(GetVArr(argv, 9)), SQLite3_Value(GetVArr(argv, 10)), SQLite3_Value(GetVArr(argv, 11)), SQLite3_Value(GetVArr(argv, 12))
    end select
    if err then sError = err.description
    on error goto 0
    
  end if
  
  if sError = "" then 
    SQLite3AccessFilter = SQLITE_OK
  else 
    SQLite3AccessFilter = SQLITE_ERROR
    if not objPrintProvider is Nothing then dbg "Error: " & sError, 0
  end if
 

  'vCursor.oRecordSet.MoveFirst

  if not objPrintProvider is Nothing then dbg "<SQLite3AccessFilter", -2
end function

Function SQLite3AccessNext (byVal hCursor)
'Virtual table method [xNext](https://www.sqlite.org/vtab.html#xnext)    
  if not objPrintProvider is Nothing then dbg ">SQLite3AccessNext hCursor = [" & hCursor & "]", +2
  dim vCursor: set vCursor = GetCursorOject(hCursor)

  SQLite3AccessNext = SQLITE_OK

  if vCursor.FindByPK = 1 then 
    vCursor.FindByPK = -1
  elseif vCursor.FindByPK = 0 then
    if vCursor.oRecordSet.eof then 
      SQLite3AccessNext = SQLITE_Error
    else 
      vCursor.oRecordSet.MoveNext      
    end if
  end if
  if not objPrintProvider is Nothing then dbg "<SQLite3AccessNext [" & SQLite3AccessNext & "]", -2  
end function


Function SQLite3AccessEOF (byVal hCursor)
'Virtual table method [xEOF](https://www.sqlite.org/vtab.html#xeof)    
  if not objPrintProvider is Nothing then dbg ">SQLite3AccessEOF hCursor = [" & hCursor & "]", +2
  dim vCursor: set vCursor = GetCursorOject(hCursor)

  SQLite3AccessEOF = 0
  
  if vCursor.FindByPK = -1 then 
    SQLite3AccessEOF = 1
  elseif vCursor.FindByPK = 1 then 
    if vCursor.oRecordSet.NoMatch  then 
      SQLite3AccessEOF = 1
      vCursor.FindByPK = -1
    end if
  else 
    if vCursor.oRecordSet.EOF then SQLite3AccessEOF = 1
  end if

  if not objPrintProvider is Nothing then dbg "<SQLite3AccessEOF " & SQLite3AccessEOF , -2
end function

Function SQLite3AccessColumn (byVal hCursor, byval hContext, byval i)
'Virtual table method [xColumn](https://www.sqlite.org/vtab.html#xcolumn)    
  dim vCursor: set vCursor = GetCursorOject(hCursor)

  dim value: set value = vCursor.oRecordSet.fields(i)
  dim sDbg 

  sDbg = "hCursor = [" & hCursor & "] hContext = [" & hContext & "] i = [" & i & "], type = [" & value.Type & "], value = [" & value.value & "] "

  If IsNull(value.value) Then
    sDbg = sDbg & " result null"
    objDynamicWrapperX.sqlite3_result_null hContext
  Else
    Select Case value.Type
      Case dbBigInt, dbBoolean, dbByte, dbInteger, dbLong
        If value.Type = dbBoolean Then
          if value.value then value = 1 else value = 0  
        else 
          value = value.value           
        end if 
        sDbg = sDbg & "result int"
        objDynamicWrapperX.sqlite3_result_int hContext, value
      Case dbDate, dbTimeStamp, dbTime
        sDbg = sDbg & "result date"
        objDynamicWrapperX.sqlite3_result_double hContext, CDbl(value.value) '+ JULIANDAY_OFFSET 
      Case dbDouble, dbFloat, dbSingle
        sDbg = sDbg & "result double"
        objDynamicWrapperX.sqlite3_result_double hContext, value.value
      Case Else
        value = value.value & ""
        sDbg = sDbg & "result text"
        objDynamicWrapperX.sqlite3_result_text16 hContext, value, Len(value) * 2, -1
    End Select
  End If

  SQLite3AccessColumn = SQLITE_OK
  if not objPrintProvider is Nothing then dbg " - SQLite3AccessColumn " & sDbg, -2  
end function

Function SQLite3AccessROWID (byVal hCursor, byval hROWID)
'Virtual table method [xRowid ](https://www.sqlite.org/vtab.html#xrowid)    
  if not objPrintProvider is Nothing then dbg ">SQLite3AccessROWID hCursor = [" & hCursor & "] hCursor = [" & hROWID & "]", +2
  dim vCursor: set vCursor = GetCursorOject(hCursor)

  dim id 
  id = vCursor.oRecordSet.fields(vCursor.oVirtualTable.nROWIDColumn).value

  if not objPrintProvider is Nothing then dbg "ROWID = [" & id & "]", 0 

  objDynamicWrapperX.numPut id, hROWID, 0, "m"

  SQLite3AccessROWID = SQLITE_OK
  if not objPrintProvider is Nothing then dbg "<SQLite3AccessROWID", -2  
end function



Function SQLite3AccessUpdate  (byVal hVTab, byval argc, byval argv, byval pRowid)
'Virtual table method [xUpdate ](https://www.sqlite.org/vtab.html#xupdate )   
   
  if not objPrintProvider is Nothing then dbg ">SQLite3AccessUpdate hVTab = [" & hVTab & "] argc = [" & argc & "] argv = [" & argv & "] pRowid = [" & pRowid & "]", +2
  dim vTab, objRecordSet, sSQL, value, sError
  sError = ""
  vTab = objDynamicWrapperX.numGet (hVTab, 3 * ptr_size, "p")
'  dbg "ovTab = [" & vTab & "]",0
  set vTab = objDynamicWrapperX.ObjGet(vTab)

  SQLite3AccessUpdate = SQLITE_OK

  if isEmpty(vTab.nROWIDColumn) or argc = 0 then 
    SQLite3AccessUpdate = SQLITE_READONLY

    if not objPrintProvider is Nothing then dbg "Строка не может быть обновлена", 0

  else  
  
    rowid = SQLite3_Value(GetVArr(argv, 0))

    if not objPrintProvider is Nothing then dbg " CanSeek = [" & vTab.CanSeek & "] rowid = [" & rowid & "] vTab.sTableName = [" & vTab.sTableName & "]", 0

    if vTab.CanSeek then 
      on error resume next
      if vTab.TableRecordSet is nothing then set vTab.TableRecordSet = vTab.oConnection.openrecordset(vTab.sTableName, dbOpenTable)
      if err then sError = err.description 
      on error goto 0
      if sError = "" then
        set objRecordSet = vTab.TableRecordSet
        if isNull(rowid) then 
          if not objPrintProvider is Nothing then dbg "Открыта таблица", 0
        else 
          if not objPrintProvider is Nothing then dbg "Поиск строки через индекс " & vTab.primaryKeyIndex.Name & ": " & rowid, 0
          objRecordSet.index = vTab.primaryKeyIndex.Name
          objRecordSet.seek "=", rowid  
          if objRecordSet.NoMatch then SQLite3AccessUpdate = SQLITE_NOTFOUND 
        end if
      end if
    else 
      if isNull(rowid) then 
        if not objPrintProvider is Nothing then dbg "Открыта таблица", 0
        on error resume next
        set objRecordSet = vTab.oConnection.openRecordset(vTab.sTableName)
        if err then sError = err.description 
        on error goto 0
      else 
        if varType(rowid)=vbString then rowid = "'"  & replace(rowid, "'", "''") & "'"
        sSQL = "select * from [" & vTab.sTableName & "] where [" & vTab.nROWIDColumn & "] = " & rowid
  
        if not objPrintProvider is Nothing then dbg "Поиск строки через запрос " & sSQL, 0
        on error resume next
        if not isNull(rowid) and objRecordSet.EOF then SQLite3AccessUpdate = SQLITE_NOTFOUND
        if err then sError = err.description 
        on error goto 0
      end if
    end if

    if sError <> "" then 
      vTab.setError sError
      if not objPrintProvider is Nothing then dbg "Ошибка ["&sError&"]",0 
      SQLite3AccessUpdate = SQLITE_Error
    end if
  
    if SQLite3AccessUpdate = SQLITE_OK then 
      if argc = 1 then 
        if not isNull(rowid) then 
          objRecordSet.Delete
        end if  
      else 
        if isNull (rowid) then
          objRecordSet.AddNew()

          rowid = objRecordSet.Fields(vTab.nROWIDColumn).value

          objDynamicWrapperX.numPut rowid, pRowid, 0, "q" 

          if not objPrintProvider is Nothing then dbg "Создана новая строка: " & rowid, 0
        else 
          value = SQLite3_Value(GetVArr(argv, 1))

          objRecordSet.Edit

          if value <> rowid then  
            'dbg "Обновление ключевого поля: " & rowid & " -> " & value, 0
            'objRecordSet.Fields(vTab.nROWIDColumn).value = value
             SQLite3AccessUpdate = SQLITE_READONLY
          end if

        end if

        if SQLite3AccessUpdate = SQLITE_OK then 
          if not objPrintProvider is Nothing then dbg "Обновление строки: " & rowid, 0
  
          dim sColumn, i
          for i = 2 to argc - 1 
            sColumn = vTab.aFields(i - 2)(0)
  
            if not objPrintProvider is Nothing then dbg " #" & sColumn, 0
  
            'if sColumn <> vTab.nROWIDColumn then 
            If ((vTab.aFields(i - 2)(2)) And dbAutoIncrField) = dbAutoIncrField Then
              value = SQLite3_Value(GetVArr(argv, i))
              if not isNull (value) and not  objRecordSet.Fields(vTab.aFields(i - 2)(0)).value = value then sError = "Ключевое поле не обновляется"
            else 
              objRecordSet.Fields(vTab.aFields(i - 2)(0)).value = SQLite3_Value(GetVArr(argv, i))
            end if
            'elseif objRecordSet.Fields(vTab.aFields(i - 2)(0)).value <> SQLite3_Value(GetVArr(argv, i)) then
            '  dbg "Обновление ключевого поля не поддерживается", 0
            'end if
          next

          if sError = "" then 
            on error resume next
            objRecordSet.Update
            if err then sError = err.description 
            on error goto 0
          end if 

          if sError <> "" then 
            vTab.setError sError
            if not objPrintProvider is Nothing then dbg "Ошибка ["&sError&"]",0 
            SQLite3AccessUpdate = SQLITE_Error
          end if 
        end if
      end if
    end if

  end if
  if not objPrintProvider is Nothing then dbg "<SQLite3AccessUpdate " & SQLite3AccessUpdate, -2  
end function





'Собираем дескриптор модуля
dim AccessConnector, ptr_size, sDbgInfo, ptr
ptr_size = objDynamicWrapperX.Bitness() / 8
AccessConnector = objDynamicWrapperX.MemAlloc(4 * 24, 1)
pPtr = AccessConnector
'номер версии - 0
objDynamicWrapperX.NumPut 0, AccessConnector, 0, "l"
objDynamicWrapperX.NumPut objDynamicWrapperX.RegisterCallback(GetRef("SQLite3AccessCreate"), "i=pplppp", "r=l", "f=c"), AccessConnector, 1 * ptr_size, "p"
objDynamicWrapperX.NumPut objDynamicWrapperX.RegisterCallback(GetRef("SQLite3AccessConnect"), "i=pplppp", "r=l", "f=c"), AccessConnector, 2 * ptr_size, "p"
objDynamicWrapperX.NumPut objDynamicWrapperX.RegisterCallback(GetRef("SQLite3AccessBestIndex"), "i=pp", "r=l", "f=c"), AccessConnector, 3 * ptr_size, "p"
objDynamicWrapperX.NumPut objDynamicWrapperX.RegisterCallback(GetRef("SQLite3AccessDisconnect"), "i=p", "r=l", "f=c"), AccessConnector, 4 * ptr_size, "p"
objDynamicWrapperX.NumPut objDynamicWrapperX.RegisterCallback(GetRef("SQLite3AccessDisconnect"), "i=p", "r=l", "f=c"), AccessConnector, 5 * ptr_size, "p"
objDynamicWrapperX.NumPut objDynamicWrapperX.RegisterCallback(GetRef("SQLite3AccessOpen"), "i=pp", "r=l", "f=c"), AccessConnector, 6 * ptr_size, "p"
objDynamicWrapperX.NumPut objDynamicWrapperX.RegisterCallback(GetRef("SQLite3AccessClose"), "i=p", "r=l", "f=c"), AccessConnector, 7 * ptr_size, "p"
objDynamicWrapperX.NumPut objDynamicWrapperX.RegisterCallback(GetRef("SQLite3AccessFilter"), "i=plplp", "r=l", "f=c"), AccessConnector, 8 * ptr_size, "p"
objDynamicWrapperX.NumPut objDynamicWrapperX.RegisterCallback(GetRef("SQLite3AccessNext"), "i=p", "r=l", "f=c"), AccessConnector, 9 * ptr_size, "p"
objDynamicWrapperX.NumPut objDynamicWrapperX.RegisterCallback(GetRef("SQLite3AccessEOF"), "i=p", "r=l", "f=c"), AccessConnector, 10 * ptr_size, "p"
objDynamicWrapperX.NumPut objDynamicWrapperX.RegisterCallback(GetRef("SQLite3AccessColumn"), "i=ppl", "r=l", "f=c"), AccessConnector, 11 * ptr_size, "p"
objDynamicWrapperX.NumPut objDynamicWrapperX.RegisterCallback(GetRef("SQLite3AccessROWID"), "i=pp", "r=l", "f=c"), AccessConnector, 12 * ptr_size, "p"
objDynamicWrapperX.NumPut objDynamicWrapperX.RegisterCallback(GetRef("SQLite3AccessUpdate"), "i=plpp", "r=l", "f=c"), AccessConnector, 13 * ptr_size, "p"
'Остальные поля не заполняются

Sub FreeResource()
  objDynamicWrapperX.MemFree AccessConnector
end sub

'*DOC*:ORDER 10000 

class SQLite_Connection
'Основное подключение к Бд SQLite

  public hDB, Application, Wrapper
  private  oDataBases

  public default function Open(dbName)
  'Открывает новуое подключение к БД 
  '#param dbName - Путь до файла с базой данных. 
  ' {*} :memory: - Специальное имя, для открытия пустой БД в памяти

    dim idResult, sError
    if not objPrintProvider is Nothing then dbg ">SQLiteConnection. Open [" & dbName & "]", +2
    if not (isEmpty(hDB) or hDB = 0) then close
    idResult = objDynamicWrapperX.sqlite3_open16 (dbName & "", hDB)
    if idResult <> SQLITE_OK then 
      sError = errmsg
      objDynamicWrapperX.sqlite3_close(hDB)
      
      if not objPrintProvider is Nothing then dbg "Fail [" & sError & "]", 0
      if not objPrintProvider is Nothing then dbg "<SQLiteConnection", -2

      err_raise "SQLLite. Не удалось открыть базу данных " & dbName & vbCrlf & sError
    end if
    set open = Me
    if not objPrintProvider is Nothing then dbg "<SQLiteConnection", -2
  end function

  public Function AttachAccessDB(oAccessDB, byval spAlias)
  'Добавляет возможность подключать к БД SQLite таблицы Access как виртуальные, а так же сохраняет внутри ссылку на подключение
  '#param oAccessDB - Ссылка на объект БД 
  '#param spAlias - алиас для подключения, передайте пустую строку для основного подключения

    if not objPrintProvider is Nothing then dbg "> AttachAccessDB = meRef[" & objDynamicWrapperX.ObjPtr(Me) & "]", + 2

    if isEmpty(hDB) then err_raise "SQLLite. Сначала необходимо открыть БД"

    if IsEmpty(oDataBases) then 
      set oDataBases = CreateObject("Scripting.Dictionary")
      oDataBases.CompareMode = 1

      'DumpMem AccessConnector, 4 * 24

      dim rc
      rc = objDynamicWrapperX.sqlite3_create_module_v2 (hDB, objDynamicWrapperX.StrPtr("access","UTF-8"), AccessConnector, objDynamicWrapperX.ObjPtr(me), 0)
      if not objPrintProvider is Nothing then dbg "sqlite3_create_module_v2 = [" & rc & "]", 0
    end if

    if spAlias = "" then spAlias = "current"
 
    if oDataBases.exists(spAlias) then 
      if not objPrintProvider is Nothing then dbg "<AttachAccessDB. база данных с именем [" & spAlias & "] уже зарегистрирована", -2
      err_raise "SQLLite. база данных с именем [" & spAlias & "] уже зарегистрирована" 
    else 
      oDataBases.add spAlias, oAccessDB 
    end if
    if not objPrintProvider is Nothing then dbg "<AttachAccessDB", -2
  end function

  Public Function GetAccessDB(byval spAlias)
  'Возвращает объект подключения к БД Access по ранее заданному алиасу
  '#param spAlias - алиас для подключения

    if not objPrintProvider is Nothing then dbg "> GetAccessDB = [" & spAlias & "]", + 2
    if spAlias = "" then spAlias = "current"
    if IsEmpty(oDataBases) then 
      set GetAccessDB = Nothing
    elseif oDataBases.exists(spAlias) then 
      set GetAccessDB = oDataBases(spAlias)
    else 
      set GetAccessDB = Nothing
    end if
    if not objPrintProvider is Nothing then dbg "< GetAccessDB", - 2
  end function

  public function close()
  'Закрывает открытое подключение

    if not isEmpty(hDB) then 
      objDynamicWrapperX.sqlite3_close(hDB)
      hDB = Empty
    end if
    set Application = Nothing
    set Wrapper = Wrapper
    oDataBases.removeall
    set oDataBases = Nothing
  end function

  
  public sub Execute(sSQL)
 'Выполняет запрос на изменение или запускает инструкцию SQL
 '#param sSQL - Запрос 

    dim hStmt, idResult
    'hStmt = objDynamicWrapperX.MemAlloc(4,1)
    idResult = objDynamicWrapperX.sqlite3_prepare16_v2(hDB, sSQL, -1, hStmt, 0)
    if idResult <> SQLITE_OK then err_raise "SQLLite.Prepare. Не удалось выполнить запрос " & sSQL & vbCrlf & "[" & idResult & "]" & errmsg()
    idResult = objDynamicWrapperX.sqlite3_step(hStmt)
    if not (idResult = SQLITE_DONE or idResult = SQLITE_ROW) then err_raise "SQLLite.Step. Не удалось выполнить запрос " & sSQL & vbCrlf & errmsg
    idResult = objDynamicWrapperX.sqlite3_finalize(hStmt)
    if idResult <> SQLITE_OK then err_raise "SQLLite.Finalize. Не удалось выполнить запрос " & sSQL & vbCrlf & errmsg
  end sub  

  public function last_insert_rowid()
  'Возвращает Id последней команды insert
    last_insert_rowid = objDynamicWrapperX.sqlite3_last_insert_rowid(hDB)
  end function

  public function libversion()
  'Возвращает версию библиотеки DLL
    libversion = objDynamicWrapperX.sqlite3_libversion()
  end function

  Private Sub Class_Terminate() 
    close
  End Sub  

  public Function PrepareSQL (sSQL) 
  'Формирует подготовленный запрос для указанного выражения и возвращает экземпляр объекта [SQLite_Prepared](#class_SQLite_Prepared) 
  '#param sSQL - Запрос 
 
    dim oStmt
    set oStmt = new  SQLite_Prepared
    oStmt.init me, sSQL
    set PrepareSQL = oStmt
  end function 

  Public Function OpenRecordSet (spSQL, pDummy)
  'Открывает набор данных для указанного выражения и возвращает экземпляр объекта [SQLite_Recordset](#class_SQLite_Recordset) 
  '#param sSQL - Запрос 
  '#param pDummy - Не используется 

    dim oStmt
    if spSQL & "" = "" then 
      set OpenRecordSet = nothing
    else 
      set oStmt = new  SQLite_Recordset
      oStmt.Open  spSQL, me, Empty, Empty
      set OpenRecordSet = oStmt
    end if
  end function 
  
  Public function errmsg()
  'Возвращает последнее сообщение об ошибке
    errmsg = objDynamicWrapperX.strGet(objDynamicWrapperX.sqlite3_errmsg(hDB),0,"UTF-8")
  end function 

end class

class SQLite_Prepared
'Подготовленное выражение. За счет того что не формируется новый план, должен работать быстрее

  dim hStmt, hDB, SQL, oConnection


  public sub init(pConnection, sSQl)
  'Инициализщирует новое подготовленное подключение
  '#param pConnection - ссылка на подключение [SQLite_Connection](#class_SQLite_Connection)
  '#param sSQL - Запрос. для подстановки параметров используйте символ '$1', '$2', ...

    dim idResult
    if not isEmpty(hStmt) then close
    set oConnection = pConnection
    hDB = pConnection.hDB  
    SQL = sSQl 
    'hStmt = objDynamicWrapperX.MemAlloc(4,1)
    idResult = objDynamicWrapperX.sqlite3_prepare16_v2(hDB, sSQl, -1, hStmt, 0)
  end sub


  public sub ExecuteByDic(tpDic)
  'Выполняет запрос. Параметры передаются в виде словаря. 
  '#param tpDic - Словарь с параметрами. Принимаются только ключи, которые начинаются с символа `$`

    dim i, idResult, index, aParam, key, value
    aParam = tpDic.keys
    idResult = objDynamicWrapperX.sqlite3_reset(hStmt)
    for each key in aParam
      if left(trim(key),1) = "$" then 
        index = objDynamicWrapperX.sqlite3_bind_parameter_index(hStmt, trim(key))
        if index > 0 then 
          value = tpDic(key) 
          if isNull(value) or isEmpty(value) then 
            idResult = objDynamicWrapperX.sqlite3_bind_null(hStmt, index)
            if idResult <> SQLITE_OK then err_raise "SQLLite.Prepared Bind Null. Не удалось забиндить параметр " & (index) &  vbCrlf & oConnection.errmsg
          ElseIf vartype(value) = 11 then ''vbBoolean  
            if value then 
              idResult = objDynamicWrapperX.sqlite3_bind_int(hStmt, index, 1)
            else 
              idResult = objDynamicWrapperX.sqlite3_bind_int(hStmt, index, 0)
            end if
            if idResult <> SQLITE_OK then err_raise "SQLLite.Prepared Bind Boolean. Не удалось забиндить параметр " & (index) &  vbCrlf & oConnection.errmsg
          ElseIf vartype(value) = 2 or vartype(value) = 3 then ''vbInteger vbLong   
            idResult = objDynamicWrapperX.sqlite3_bind_int(hStmt, index, value)
            if idResult <> SQLITE_OK then err_raise "SQLLite.Prepared Bind Integer. Не удалось забиндить параметр " & (index) &  vbCrlf & oConnection.errmsg
          ElseIf vartype(value) = 4 or vartype(value) = 5 or vartype(value) = 6 or vartype(avalue) = 7 then ''vbSingle  vbDouble vbCurrency vbDate       
            idResult = objDynamicWrapperX.sqlite3_bind_double(hStmt, index, value)
            if idResult <> SQLITE_OK then err_raise "SQLLite.Prepared Bind Double. Не удалось забиндить параметр " & (index) &  vbCrlf & oConnection.errmsg
          else 
            idResult = objDynamicWrapperX.sqlite3_bind_text16(hStmt, index, value, -1)
            if idResult <> SQLITE_OK then err_raise "SQLLite.Prepared Bind Text. Не удалось забиндить параметр " & (index) &  vbCrlf & oConnection.errmsg
          end if
        end if
      end if
    next
    idResult = objDynamicWrapperX.sqlite3_step(hStmt)
    if not(idResult = SQLITE_ROW or  idResult = SQLITE_DONE) then err_raise "SQLLite.Execute. Не удалось выполнить запрос" & vbCrlf & oConnection.errmsg
  end sub
  
  public sub Execute(aParam)
  'Выполняет запрос. Параметры передаются в виде массива. 
  '#param aParam - Массив с параметрами

    dim i, idResult, index
    idResult = objDynamicWrapperX.sqlite3_reset(hStmt)
    for i = 0 to ubound(aParam)
      index = objDynamicWrapperX.sqlite3_bind_parameter_index(hStmt, "$" & (i+1))
      if index > 0 then 
        if isNull(aParam(i)) or isEmpty(aParam(i)) then 
          idResult = objDynamicWrapperX.sqlite3_bind_null(hStmt, index)
          if idResult <> SQLITE_OK then err_raise "SQLLite.Prepared Bind Null. Не удалось забиндить параметр " & (index) &  vbCrlf & oConnection.errmsg
        ElseIf vartype(aParam(i)) = 11 then ''vbBoolean  
          if aParam(i) then 
            idResult = objDynamicWrapperX.sqlite3_bind_int(hStmt, index, 1)
          else 
            idResult = objDynamicWrapperX.sqlite3_bind_int(hStmt, index, 0)
          end if
          if idResult <> SQLITE_OK then err_raise "SQLLite.Prepared Bind Boolean. Не удалось забиндить параметр " & (index) &  vbCrlf & oConnection.errmsg
        ElseIf vartype(aParam(i)) = 2 or vartype(aParam(i)) = 3 then ''vbInteger vbLong   
          idResult = objDynamicWrapperX.sqlite3_bind_int(hStmt, index, aParam(i))
          if idResult <> SQLITE_OK then err_raise "SQLLite.Prepared Bind Integer. Не удалось забиндить параметр " & (index) &  vbCrlf & oConnection.errmsg
        ElseIf vartype(aParam(i)) = 4 or vartype(aParam(i)) = 5 or vartype(aParam(i)) = 6 or vartype(aParam(i)) = 7 then ''vbSingle  vbDouble vbCurrency vbDate       
          idResult = objDynamicWrapperX.sqlite3_bind_double(hStmt, index, aParam(i))
          if idResult <> SQLITE_OK then err_raise "SQLLite.Prepared Bind Double. Не удалось забиндить параметр " & (index) &  vbCrlf & oConnection.errmsg
        else 
          idResult = objDynamicWrapperX.sqlite3_bind_text16(hStmt, index, aParam(i), -1)
          if idResult <> SQLITE_OK then err_raise "SQLLite.Prepared Bind Text. Не удалось забиндить параметр " & (index) &  vbCrlf & oConnection.errmsg
        end if
      end if
    next
    idResult = objDynamicWrapperX.sqlite3_step(hStmt)
    if not(idResult = SQLITE_ROW or  idResult = SQLITE_DONE) then err_raise "SQLLite.Execute. Не удалось выполнить запрос" & vbCrlf & oConnection.errmsg
  end sub

  public sub close()
    'Закрывает подготовленный запрос и освобождает ресурсы
    if not isEmpty(hStmt) then 
      idResult = objDynamicWrapperX.sqlite3_finalize(hStmt)
      if idResult <> SQLITE_OK then err_raise "SQLLite.Finalize. Не удалось завершить запрос " & sSQL & vbCrlf & oConnection.errmsg
      hStmt = Empty
    end if
    set oConnection = Nothing
  end sub

  Private Sub Class_Terminate() 
    close
  End Sub  

end class


class SQLite_Recordset
'Объект представляет записи, получаемые в результате выполнения запросов. 

  dim hStmt, oConnection
  public Fields
  
  public function Open(spSQL, pConnection, dummy1, dummy2)
  'Инициализирует набор даннных по запросу
  '#param sSQL - Запрос 
  '#param pConnection - ссылка на подключение [SQLite_Connection](#class_SQLite_Connection)
  '#param dummy1 - Не используется 
  '#param dummy2 - Не используется 
   
    dim idResult
    set oConnection = pConnection
    if not isEmpty(hStmt) then 
      close
    end if   

    'hStmt = objDynamicWrapperX.MemAlloc(4,1)    
    idResult = objDynamicWrapperX.sqlite3_prepare16_v2(pConnection.hDB, spSQL & "", -1, hStmt, 0)

    if idResult <> SQLITE_OK then err_raise "SQLLite.Prepare. Не удалось выполнить запрос " & vbCrlf & spSQL & " idResult = ["&idResult&"]" & vbCrlf & oConnection.errmsg

    idResult = objDynamicWrapperX.sqlite3_reset(hStmt)
    idResult = objDynamicWrapperX.sqlite3_step(hStmt)

    if idResult = SQLITE_ROW then
      dim nRowCount, i, field 
      Set Fields = CreateObject("System.Collections.ArrayList")
      nRowCount = objDynamicWrapperX.sqlite3_column_count(hStmt)       
      for i = 0 to nRowCount - 1
        set field= new SQLLite_Field
        field.Create hStmt, i, Me
        Fields.add field
      next         
    else 
      close
    end if

  end function

  Public Default function FieldByName(name)
  'Возвращает поле [SQLLite_Field](class_SQLLite_Field) по его имени
  '#param name - имя поля
  

    dim field
    name = ucase(name)
    for each field in Fields
      if uCase(field.name) = name then 
        set FieldByName = field 
        exit function
      end if
    next
    err_raise "SQLLite.FieldByName. В наборе данных нет поля с именем {" & name & "}" 
  end function

  public function EOF()
  'Возвращает значение, которое указывает, находится ли позиция текущей записи после последней записи
    EOF = isEmpty(hStmt)   
  end function
  
  public function MoveNext 
  'Выполняет перемещение к следующей записи и делает запись текущей записью
    dim idResult
    if not eof then 
      idResult = objDynamicWrapperX.sqlite3_step(hStmt)
      if idResult = SQLITE_ROW then 
        ''Все хорошо дальше через Field получить значения
      elseif idResult = SQLITE_DONE then 
        close'' Больше записей нет все закрываем
      else 
        close
        err_raise "SQLLite.MoveNext. Не удалось выполнить запрос " & sSQL & " idResult = ["&idResult&"]" & vbCrlf & oConnection.errmsg
      end if      
    end if
  end function


  public function close()
  'Закрывает открытый объект

    if not isEmpty(hStmt) then 
      'Уничтожим все поля
      if not isEmpty(Fields) then 
        dim Field 
        for each Field in Fields
          Field.close
        next
        Fields.clear
        set Fields = nothing
        Fields = empty
      end if
      'Уничтожим в памяти запрос
      idResult = objDynamicWrapperX.sqlite3_finalize(hStmt)
      if idResult <> SQLITE_OK then dbg "SQLLite.Finalize. Не удалось закрыть Recordset idResult = ["&idResult&"]" & vbCrlf & oConnection.errmsg, 0
      hStmt = empty
      set oConnection = Nothing
    end if
  end function

  Private Sub Class_Terminate() 
    close
  End Sub  
  
end class

class SQLLite_Field
'Объект представляет столбец данных с общим типом данных и общим набором свойств

  dim ColumnNumber, sColumnName, hStmt, oRS

  public function Create(phStmt, pColumnNumber, oRecordSet)
  'Инициализация объекта
  '#param phStmt - Дескриптор текущего выражения
  '#param pColumnNumber - Номер колонки
  '#param oRecordSet - Объект набора данных [SQLite_Recordset](#class_SQLite_Recordset) 

    ColumnNumber = pColumnNumber
    hStmt = phStmt    
    sColumnName = empty
    set oRs = oRecordSet
  end function

  Public Default Property Get Value
  'Возвращает значение поля

    if hStmt = 0 then
      Value = empty
    else 
      Select Case objDynamicWrapperX.sqlite3_column_type(hStmt, ColumnNumber)  
        Case SQLITE_NULL
          Value = Null
        Case SQLITE_INTEGER
          Value = objDynamicWrapperX.sqlite3_column_int(hStmt, ColumnNumber)
        Case SQLITE_FLOAT
          Value = objDynamicWrapperX.sqlite3_column_double(hStmt, ColumnNumber)
        Case SQLITE_TEXT
          Value = objDynamicWrapperX.sqlite3_column_text16(hStmt, ColumnNumber)
        Case SQLITE_BLOB
          err_raise "SQLLite.Field.Value. Тип BLOB не реализован "
        Case Else
          err_raise "SQLLite.Field.Value. Не удалось получить значение столбца " & ColumnNumber
      End Select
    end if
  End Property

  Public Property Get Name
  'Возвращает имя поля

    if hStmt = 0 then 
      Name = Empty
    else 
      if isEmpty(sColumnName) then sColumnName = objDynamicWrapperX.sqlite3_column_name16(hStmt, ColumnNumber)
      Name = sColumnName
    end if
  End Property 

  Public Property Get OrdinalPosition
  'Возвращает позицию поля в наборе данных
    OrdinalPosition = ColumnNumber
  End Property 

  Public Sub Close()
  'Закрывает объект и освобождает ресурсы
    hStmt = 0
    set oRs = Nothing
  end sub

  Private Sub Class_Terminate() 
    close
  End Sub  
end class