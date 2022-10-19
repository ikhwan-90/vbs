Dim inp_prev_month, inp_current_month, prev_month, current_month, current_dir, count_changes

count_changes = 0
'User input for which text to replace
inp_prev_month=InputBox("Input current text to be changed: ","Step 1")

if IsEmpty(inp_prev_month) Then
	MsgBox "Operation Canceled!",0,"Notice"
Else
	inp_current_month=InputBox("Replace text with: ","Step 2")
	
	if IsEmpty(inp_current_month) Then
		MsgBox "Operation Canceled!",0,"Notice"
	Else
		'Capital first letter check
		prev_month = UCase(Left(inp_prev_month,1)) & Mid(inp_prev_month,2)
		current_month = UCase(Left(inp_current_month,1)) & Mid(inp_current_month,2)
		
		Set objFso = CreateObject("Scripting.FileSystemObject")
		current_dir = objFSO.GetParentFolderName(WScript.ScriptFullName)
		Set Folder = objFSO.GetFolder(current_dir)
		MsgBox ("Current Dir: " & current_dir & vbCrLf & "Current Text on FileName: " & prev_month & vbCrLf & "Replaced with: " & current_month), 0, "Change Summary"
		
		'Run Function
		Browsefolder objFso.GetFolder(current_dir)
		
		
	end if
end if

'Function recursive for subfolder and files
Sub Browsefolder(Folder)
	For Each Subfolder In Folder.Subfolders
		Browsefolder Subfolder
	Next
	For Each File In Folder.Files
		sNewFile = File.Name
		sNewFile = Replace(sNewFile,UCase(prev_month),UCase(current_month))
		sNewFile = Replace(sNewFile,prev_month,current_month)
		if (sNewFile<>File.Name) Then 
			File.Move(File.ParentFolder+"\"+sNewFile)
			On Error Resume Next
			count_changes = count_changes + 1
		end if
	Next
	MsgBox count_changes & " Document(s) changed!", 0, "Successful Process Summary"
End Sub

