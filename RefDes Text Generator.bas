Sub Main()
Dim layname,horientation,vorientation,flag,angle

not_in_valid_layer = ""
no_refdes_label = ""
not_valid_ctr = 0
norefdes_ctr = 0
lblctr = 0
refdesctr = 0
no_of_comp = 0


ActiveDocument.SelectObjects(ppcbObjectTypeLabel, "", True)
For Each comp In ActiveDocument.GetObjects(ppcbObjectTypeComponent)
	no_of_comp = no_of_comp + 1
	flag_refdes_label = 0
	For Each lbl In comp.Labels
		If lbl.Name = "Ref.Des." Then
			flag_refdes_label = 1
			refdesctr = refdesctr + 1
			flag = 0
			angle = 0
			For Each lay In ActiveDocument.Layers
				If lay.Number = lbl.layer Then
					If lay.Number = 26 Or lay.Number = 29 Or lay.Number = 1 Or lay.Name = "Bottom" Then
						If lay.Name = "Silkscreen Top" Or lay.Name = "Silkscreen Bottom" Or lay.Name = "Top" Or lay.Name = "Bottom" Then
							If comp.layer = 1 Then
								layname = lay.Name
							Else
								If lay.Number = 26 Then
									layname = "Silkscreen Bottom"
								ElseIf lay.Number = 29 Then
									layname = "Silkscreen Top"
								ElseIf lay.Number = 1 Then
									layname = "Bottom"
								Else
									layname = "Top"
								End If
							End If
							flag = 1
							lblctr = lblctr + 1
						Else
							MsgBox("The Layer Definition is not in English Language")
							End
						End If
					Else
						If not_valid_ctr = 0 Then
							not_in_valid_layer = lbl.Text
							not_valid_ctr = 1
						Else
							not_in_valid_layer = not_in_valid_layer + "," + lbl.Text
						End If
					End If
					GoTo proceed
				End If
			Next lay
proceed:
			Select Case lbl.HorzJustification 
				Case ppcbJustifyLeft 
					horientation = "Left" 
				Case ppcbJustifyHCenter 
					horientation = "Center" 
				Case ppcbJustifyRight 
					horientation = "Right" 
			End Select 
			Select Case lbl.VertJustification 
				Case ppcbJustifyBottom 
					vorientation = "Down" 
				Case ppcbJustifyVCenter 
					vorientation = "Center" 
				Case ppcbJustifyTop 
					vorientation = "Up" 
			End Select
			If flag = 1 Then
				f = "Application.ExecuteCommand(" + Chr(34) + "Create Text" + Chr(34) + ")" & vbLf
				f = f & "AddFreeTextDlg.FontFace = "  + Chr(34) + "<Romansim Stroke Font>"  + Chr(34) & vbLf
				f = f & "AddFreeTextDlg.TextString = " + Chr(34) + lbl.Text + Chr(34) & vbLf
				f = f & "AddFreeTextDlg.LayerName = " + Chr(34) + layname + Chr(34) & vbLf
				f = f & "AddFreeTextDlg.XCoord = " + Chr(34) + lbl.PositionX(ppcbUnitMetric,ppcbOriginTypeDesign) + Chr(34) & vbLf
				f = f & "AddFreeTextDlg.YCoord = " + Chr(34) + lbl.PositionY(ppcbUnitMetric,ppcbOriginTypeDesign) + Chr(34) & vbLf
				If lbl.Orientation > 360 Or lbl.Orientation < -360 Then
					If lbl.Orientation > 360 Then
						angle = lbl.Orientation - 360
					Else
						angle = 720 + lbl.Orientation
					End If
				Else
					angle = lbl.Orientation
				End If
			
				f = f & "AddFreeTextDlg.RotationAngle = " + Chr(34) + angle + Chr(34) & vbLf
				f = f & "AddFreeTextDlg.TextHeight = " + Chr(34) + lbl.Height + Chr(34) & vbLf
				f = f & "AddFreeTextDlg.LineWidth = " + Chr(34) + lbl.LineWidth + Chr(34) & vbLf
				f = f & "AddFreeTextDlg.HorizontalJustification =  " + Chr(34) + horientation + Chr(34) & vbLf
				f = f & "AddFreeTextDlg.VerticalJustification = "  + Chr(34) + vorientation + Chr(34) & vbLf
				f = f & "AddFreeTextDlg.Mirrored = " + lbl.Mirror & vbLf
				f = f & "AddFreeTextDlg.FontFace = "  + Chr(34) + "Lucida Sans Unicode"  + Chr(34) & vbLf
				f = f & "AddFreeTextDlg.Ok.Click()" & vbLf
				f = f & "AddFreeTextDlg.Cancel.Click()" & vbLf
				Application.RunMacro "", f
				lbl.Delete()
			End If
		End If
	Next lbl
	If flag_refdes_label = 0 Then
		If norefdes_ctr = 0 Then
			no_refdes_label = comp.Name
			norefdes_ctr = 1
		Else
			no_refdes_label = no_refdes_label + "," + comp.Name
		End If
	End If
Next comp

repsum_file = "C:\PADS Projects\silkscreen_report_summary.txt"
Open repsum_file For Output As #1
Print #1, Str(lblctr) + " out of " + Str(refdesctr) + " Ref.Des. Labels were changed to Text";
Print #1
If not_in_valid_layer <> "" Then
	Print #1, "The following Ref.Des. labels are not assigned to Valid Layers (Top, Bottom, or Silkscreen Top/Bottom): " + not_in_valid_layer;
	Print #1
End If
Print #1, Str(refdesctr) + " out of " + Str(no_of_comp) + " Components have Ref.Des.Labels";
Print #1
If no_refdes_label <> "" Then
	Print #1, "The following Components doesn't have Ref.Des. Labels: " + no_refdes_label;
End If

Close #1
Shell "Notepad " & repsum_file, 3
MsgBox("Ref.Des Text Generation Completed")
End Sub
