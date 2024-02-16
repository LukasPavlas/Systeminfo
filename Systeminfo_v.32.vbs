' Získání aktuálního adresáře, kde je umístěn skript
Set objFSO = CreateObject("Scripting.FileSystemObject")
strScriptDirectory = objFSO.GetParentFolderName(WScript.ScriptFullName)

Set objShell = CreateObject("WScript.Shell")
strComputerName = objShell.ExpandEnvironmentStrings("%COMPUTERNAME%")

Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
Set colBIOS = objWMIService.ExecQuery("Select * from Win32_BIOS")
For Each objBIOS in colBIOS
    strSerialNumber = objBIOS.SerialNumber
Next

Set colAdapters = objWMIService.ExecQuery("Select * from Win32_NetworkAdapter")
Set colIPConfig = objWMIService.ExecQuery("Select * from Win32_NetworkAdapterConfiguration")

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


' Nastavení cesty k souboru v kořenovém adresáři C:\, relativní cesta k skriptu
strJsonFilePath = objFSO.BuildPath(strScriptDirectory, "Systeminfo - " & strComputerName & ".json")

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
objJsonFile.WriteLine """NetworkAdapters"": ["

For Each strAdapterName In objNetworkAdapters.Keys
    arrAdapterInfo = objNetworkAdapters.Item(strAdapterName)
    strMacAddress = arrAdapterInfo(0)
    strIPAddress = arrAdapterInfo(1)

    objJsonFile.WriteLine "  {"
    objJsonFile.WriteLine "    ""AdapterName"": """ & strAdapterName & ""","
    objJsonFile.WriteLine "    ""MacAddress"": """ & strMacAddress & ""","
    objJsonFile.WriteLine "    ""IPAddress"": """ & strIPAddress & """"
    objJsonFile.WriteLine "  },"
Next

objJsonFile.WriteLine ""
objJsonFile.WriteLine "]}"
objJsonFile.Close

WScript.Echo "Data byla ulozena do " & strJsonFilePath