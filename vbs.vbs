Dim fso, txtFile
Set fso = CreateObject("Scripting.FileSystemObject")
Set txtFile = fso.CreateTextFile("C:\Users\Public\yayayaya.txt", True)
txtFile.WriteLine("yayayaya")
txtFile.Close
