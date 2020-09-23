Attribute VB_Name = "Module1"


Public Declare Function GetAsyncKeyState Lib "USER32" (ByVal vKey As Long) As Integer

Declare Function SystemParametersInfo Lib "user32.dll" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uiParam As Long, pvParam As Any, ByVal fWinIni As Long) As Long

Public Const SPI_SCREENSAVERRUNNING = 97

