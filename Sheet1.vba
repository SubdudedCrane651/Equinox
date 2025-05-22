Sub Python_bargraph()
    Dim objShell As Object
    Set objShell = VBA.CreateObject("WScript.Shell")
    objShell.Run "C:/Users/rchrd/AppData/Local/Programs/Python/Python312/python.exe C:\Users\rchrd\Documents\Python\Equinox\calculate.py"
    Set objShell = Nothing
End Sub

Sub Python_CreateMySql()
    Dim objShell As Object
    Set objShell = VBA.CreateObject("WScript.Shell")
    objShell.Run "c:/Users/rchrd/Documents/Python/Chargepoint+/.venv/Scripts/python.exe c:/Users/rchrd/Documents/Python/Chargepoint+/MyEVSqlxls.py"
    Set objShell = Nothing
End Sub