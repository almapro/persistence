On Error Resume Next
path = Wscript.Arguments.Item(0)
Dim oFso : Set oFso = CreateObject("Scripting.FileSystemObject")  
Dim oSap : Set oSap = CreateObject("Shell.Application")
Dim count, root, base, add, out
If Not oFso.FileExists(path) Then
  ' Not found
  WScript.Quit -1
end if
root = oFso.GetParentFolderName(path)
base = oFso.GetBaseName(path)
out = root & "\" & base
If oFso.FolderExists(out) Then 
  add = 2 : Do
    out = root & "\" & base & "-" & add
    If Not oFso.FolderExists(out) Then Exit Do
    add = add + 1
  Loop
End If 
oFso.CreateFolder(out)   
oSap.NameSpace(out).CopyHere oSap.NameSpace(path).Items, 20 
If FileCount(oSap, path) = FileCount(oSap, out) Then
  ' Done
Else
  oFso.DeleteFolder out, True
end if