Dim fso, txtFile
Set fso = CreateObject("Scripting.FileSystemObject")
Set txtFile = fso.CreateTextFile("C:\Users\Public\THISISINSANSE.txt", True)
txtFile.WriteLine("THISISINSANSE")
txtFile.Close
