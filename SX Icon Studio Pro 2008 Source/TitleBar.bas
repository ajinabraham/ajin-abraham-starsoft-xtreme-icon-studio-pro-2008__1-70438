Attribute VB_Name = "TitleBar"
 
 'Created  by Sanjul.A.S
 'E-Mai : ssanjul@gmail.com
 
 Sub SetupTitlebar(Form As Form)
'-----------------------------------------
'Form.IconImg.Picture = Form.Icon
Form.TitleBar.Top = -2
Form.TitleBar.Left = -5
Form.TitleBar.Width = Form.Width
Dim buttonHeight As Integer
Dim buttonTop As Integer
Dim buttonWidth As Integer
buttonTop = Form.TitleBar.Height * (20 / 100)
buttonHeight = Form.TitleBar.Height * (75 / 100)
buttonWidth = buttonHeight
'-----------------------------------
Form.Closeb.Top = buttonTop
Form.Restore.Top = buttonTop
Form.Minimize.Top = buttonTop
Form.Maximize.Top = buttonTop
Form.IconImg.Top = buttonTop
'----------------------------------
Form.Closeb.Height = buttonHeight
Form.Restore.Height = buttonHeight
Form.Minimize.Height = buttonHeight
Form.Maximize.Height = buttonHeight
Form.IconImg.Height = buttonHeight
'-----------------------------------
Form.Closeb.Width = buttonWidth
Form.Restore.Width = buttonWidth
Form.Minimize.Width = buttonWidth
Form.Maximize.Width = buttonWidth
Form.IconImg.Width = buttonWidth
'-----------------------------------
Form.Closeb.Left = Form.TitleBar.Width - Form.Closeb.Width - 50
Form.Restore.Left = Form.Closeb.Left - Form.Restore.Width
Form.Maximize.Left = Form.Restore.Left
Form.Minimize.Left = Form.Maximize.Left - Form.Minimize.Width
Form.IconImg.Left = 80
Form.WindowCaption.Left = Form.IconImg.Left + Form.IconImg.Width + 80
'----------------------------------------------------

End Sub

Sub ButtonDown(Button As Image)
Button.BorderStyle = 1
End Sub
Sub buttonup(Button As Image)
Button.BorderStyle = 0
End Sub


Sub MVRestore(Form As Form)
buttonup Form.Closeb
buttonup Form.Minimize
buttonup Form.Maximize
ButtonDown Form.Restore
End Sub
Sub MVMaximize(Form As Form)
buttonup Form.Closeb
buttonup Form.Minimize
ButtonDown Form.Maximize
buttonup Form.Restore
End Sub

Sub MVMinimize(Form As Form)
buttonup Form.Closeb
ButtonDown Form.Minimize
buttonup Form.Maximize
buttonup Form.Restore
End Sub


Sub MVCloseb(Form As Form)
ButtonDown Form.Closeb
buttonup Form.Minimize
buttonup Form.Maximize
buttonup Form.Restore
End Sub

