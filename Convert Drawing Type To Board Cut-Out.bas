Sub Main()
	decision = MsgBox("Converted shapes will not be copied if it will be used in Multiple Boards" + Chr(13) + "Are you sure you want to proceed in converting the shapes into Cut-out?", 4 + 32)
	If decision = 6 Then
		selected_ctr = 0
		converted_ctr = 0
		unconverted_ctr = 0
		Set selected_drawings = ActiveDocument.GetObjects(ppcbObjectTypeDrawing, "", True)
		For Each drowing In selected_drawings
			flag = 0
			selected_ctr = selected_ctr + 1
			dname = drowing.Name
			ActiveDocument.SelectObjects(ppcbObjectTypeDrawing, "*", False)
			ActiveDocument.SelectObjects(ppcbObjectTypeDrawing, dname, True)
			f = "Application.ExecuteCommand(" + Chr(34) + "Properties" + Chr(34) + ")" & vbLf
			f = f & "DraftingPropertiesDlg.DraftingType = " + Chr(34) + "Board Cut Out" + Chr(34) & vbLf
			f = f & "DraftingPropertiesDlg.Ok.Click()" & vbLf
			f = f & "DlgPrompt.Question(" + Chr(34) + "Board Cut Out intersects the Board Outline or another Cut Out." + Chr(34) + ").Answer(mbOK)" & vbLf
			Application.RunMacro "",f
			ActiveDocument.SelectObjects(ppcbObjectTypeDrawing, "*", False)
			ActiveDocument.SelectObjects(ppcbObjectTypeDrawing, dname, True)
			For Each tmp In ActiveDocument.GetObjects(ppcbObjectTypeDrawing, "", True)
				flag = 1
			Next
			If flag = 0 Then
				converted_ctr = converted_ctr + 1
			End If
		Next
		unconverted_ctr = selected_ctr - converted_ctr
		MsgBox("Success:" + Str(converted_ctr) + "/" + Str(selected_ctr) + Chr(13) +"Failed:" + Str(unconverted_ctr) + "/" + Str(selected_ctr),mbOK,"Conversion Results")
	End If
End Sub
