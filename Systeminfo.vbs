' Získání aktuálního adresáře, kde je umístěn skript
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("WScript.Shell")

strDesktopPath = objShell.SpecialFolders("Desktop")
strComputerName = objShell.ExpandEnvironmentStrings("%COMPUTERNAME%")

Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
Set colBIOS = objWMIService.ExecQuery("Select * from Win32_BIOS")
For Each objBIOS in colBIOS
    strSerialNumber = objBIOS.SerialNumber
Next

Set colAdapters = objWMIService.ExecQuery("Select * from Win32_NetworkAdapter")
Set colIPConfig = objWMIService.ExecQuery("Select * from Win32_NetworkAdapterConfiguration")

Set colInstalledPrinters = objWMIService.ExecQuery("Select * from Win32_Printer")
Set colInstalledPorts = objWMIService.ExecQuery("Select * from Win32_TCPIPPrinterPort")

Function GetIPAddress(interfaceIndex)
    On Error Resume Next
    For Each objConfig In colIPConfig
        If objConfig.Index = interfaceIndex Then
            If Not IsNull(objConfig.IPAddress) Then
                GetIPAddress = objConfig.IPAddress(0)
                Exit Function
            End If
        End If
    Next
    GetIPAddress = "N/A"
End Function

' Deklarace funkce pro odstranění diakritiky
Function RemoveDiacritics(strInput)
    Dim objRegExp, strOutput
    Set objRegExp = New RegExp
    objRegExp.IgnoreCase = True
    objRegExp.Global = True
    objRegExp.Pattern = "[^\u0000-\u007F]" ' regulární výraz pro nalezení znaků mimo rozsah ASCII
    strOutput = objRegExp.Replace(strInput, "") ' nahrazení znaků bez diakritiky
    RemoveDiacritics = strOutput
End Function


Set colProcessors = objWMIService.ExecQuery("Select * from Win32_Processor")
For Each objProcessor In colProcessors
    strCPUInfo = objProcessor.Name
Next

Set colVideoControllers = objWMIService.ExecQuery("Select * from Win32_VideoController")
For Each objVideoController In colVideoControllers
    strGPUInfo = objVideoController.Name
Next

Set colPhysicalMemory = objWMIService.ExecQuery("Select * from Win32_PhysicalMemory")
lngTotalRAM = 0
For Each objMemory In colPhysicalMemory
    lngTotalRAM = lngTotalRAM + objMemory.Capacity
Next
strRAMInfo = FormatNumber(lngTotalRAM / 1024^3, 2) & " GB"

Set colAdapters = objWMIService.ExecQuery("Select * from Win32_NetworkAdapter")
Set objNetworkAdapters = CreateObject("Scripting.Dictionary")

For Each objAdapter In colAdapters
    strAdapterName = objAdapter.Description
    strMacAddress = objAdapter.MACAddress
    strIPAddress = GetIPAddress(objAdapter.Index)

    ' Zkontrolujeme, zda klíč (název adaptéru) již existuje v kolekci
    If Not objNetworkAdapters.Exists(strAdapterName) Then
        objNetworkAdapters.Add strAdapterName, Array(strMacAddress, strIPAddress)
    End If
Next


Set colItems = objWMIService.ExecQuery("Select * from Win32_ComputerSystem")
For Each objItem In colItems
    strModel = objItem.Model
    strManufacturer = objItem.Manufacturer
Next

Const ForReading = 1, ForWriting = 2, ForAppending = 8
Const HKEY_CURRENT_USER = &H80000001
Const HKEY_LOCAL_MACHINE = &H80000002

strComputer = "."
 
Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _
    strComputer & "\root\default:StdRegProv")
 
strKeyPath = "SOFTWARE\TeamViewer"
strValueName = "ClientID"
oReg.GetDWORDValue HKEY_LOCAL_MACHINE,strKeyPath,strValueName,dwTVID


' Nastavení cesty k souboru v kořenovém adresáři C:\, relativní cesta k skriptu
'strJsonFilePath = objFSO.BuildPath(strScriptDirectory, "Systeminfo - " & strComputerName & ".json")

' Změna cesty k souboru na cestu k ploše
strJsonFilePath = objFSO.BuildPath(strDesktopPath, "Systeminfo - " & strComputerName & ".json")

Set objFileSystem = CreateObject("Scripting.FileSystemObject")
Set objJsonFile = objFileSystem.CreateTextFile(strJsonFilePath, True)

objJsonFile.WriteLine "{"
objJsonFile.WriteLine """ComputerName"": """ & strComputerName & ""","
objJsonFile.WriteLine """SerialNumber"": """ & strSerialNumber & ""","
objJsonFile.WriteLine """Manufacturer"": """ & strManufacturer & ""","
objJsonFile.WriteLine """Model"": """ & strModel & ""","
objJsonFile.WriteLine """CPU"": """ & strCPUInfo & ""","
objJsonFile.WriteLine """GPU"": """ & strGPUInfo & ""","
objJsonFile.WriteLine """RAM"": """ & strRAMInfo & ""","
objJsonFile.WriteLine """TeamviewerID"": """& dwTVID & ""","
objJsonFile.WriteLine """NetworkAdapters"": ["

Dim adapterCount
adapterCount = objNetworkAdapters.Count
Dim adapterIndex
adapterIndex = 0

For Each strAdapterName In objNetworkAdapters.Keys
    adapterIndex = adapterIndex + 1
    arrAdapterInfo = objNetworkAdapters.Item(strAdapterName)
    strMacAddress = arrAdapterInfo(0)
    strIPAddress = arrAdapterInfo(1)

    objJsonFile.WriteLine "  {"
    objJsonFile.WriteLine "    ""AdapterName"": """ & strAdapterName & ""","
    objJsonFile.WriteLine "    ""MacAddress"": """ & strMacAddress & ""","
    objJsonFile.WriteLine "    ""IPAddress"": """ & strIPAddress & """"
    If adapterIndex = adapterCount Then
        objJsonFile.WriteLine "}"
    Else
        objJsonFile.WriteLine "},"
    End If
    'objJsonFile.WriteLine "  }"
Next

objJsonFile.WriteLine "],"

objJsonFile.WriteLine """Printers"": ["

Dim printerCount
printerCount = colInstalledPrinters.Count
Dim printerIndex
printerIndex = 0

For Each objPrinter in colInstalledPrinters
    printerIndex = printerIndex + 1
    objJsonFile.WriteLine "  {"
    objJsonFile.WriteLine "    ""PrinterName"": """ & RemoveDiacritics(objPrinter.Name) & ""","
    objJsonFile.WriteLine "    ""PrinterLocation"": """ & objPrinter.Location & ""","
    objJsonFile.WriteLine "    ""PrinterPortName"": """ & objPrinter.PortName & ""","
    objJsonFile.WriteLine "    ""PrinterDefault"": """ & objPrinter.Default & """"
    If printerIndex = printerCount Then
        objJsonFile.WriteLine "}"
    Else
        objJsonFile.WriteLine "},"
    End If
    'objJsonFile.WriteLine "  }"
Next

objJsonFile.WriteLine "],"

objJsonFile.WriteLine """PortsTCPIP"": ["

Dim portCount
portCount = colInstalledPorts.Count
Dim portIndex
portIndex = 0

for Each objPort in colInstalledPorts
    portIndex = portIndex + 1
    objJsonFile.WriteLine "  {"
    objJsonFile.WriteLine "    ""PortName"": """ & objPort.Name & """"
    If portIndex = portCount Then
        objJsonFile.WriteLine "}"
    Else
        objJsonFile.WriteLine "},"
    End If
    'objJsonFile.WriteLine "  }"
Next

objJsonFile.WriteLine "]"

objJsonFile.WriteLine "}"

objJsonFile.Close

WScript.Echo "Data byla ulozena do " & strJsonFilePath
