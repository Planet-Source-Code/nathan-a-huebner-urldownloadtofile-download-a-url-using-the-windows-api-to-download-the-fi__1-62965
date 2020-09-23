Attribute VB_Name = "DownloadBAS"
Public Declare Function URLDownloadToFile Lib "urlmon" Alias _
"URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal _
szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long

Sub DownloadFile(URL As String, SaveAsFile As String)
Dim b
b = URLDownloadToFile(0, URL, SaveAsFile, 0, 0)
End Sub

