' MP3 Tag Fixer and MP3 Renamer
' Copyright (c) Samuel Gomes, 1999-2020
' mailto: v_2samg@hotmail.com || gomes.samuel@gmail.com
'
' NameP3 uses the "UltraID3Lib: MP3 ID3 Tag Editor" 

Imports System.IO
Imports HundredMilesSoftware.UltraID3Lib

Public Class FrmMain
	Inherits Form

	' Win32 Shell About box
	Private Declare Function ShellAbout Lib "shell32" Alias "ShellAboutA" (ByVal hWnd As IntPtr, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As IntPtr) As Integer

#Region " Windows Form Designer generated code "

	Public Sub New()
		MyBase.New()

		'This call is required by the Windows Form Designer.
		InitializeComponent()

		'Add any initialization after the InitializeComponent() call

	End Sub

	'Form overrides dispose to clean up the component list.
	Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
		If disposing Then
			If Not components Is Nothing Then
				components.Dispose()
			End If
		End If
		MyBase.Dispose(disposing)
	End Sub

	'Required by the Windows Form Designer
	Private ReadOnly components As System.ComponentModel.IContainer

	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.  
	'Do not modify it using the code editor.
	Friend WithEvents TxtTitle As TextBox
	Friend WithEvents CmdFixTag As Button
	Friend WithEvents CmdClear As Button
	Friend WithEvents CmdAbout As Button
	Friend WithEvents CmdExit As Button
	Friend WithEvents TxtComment As TextBox
	Friend WithEvents TxtYear As TextBox
	Friend WithEvents TxtAlbum As TextBox
	Friend WithEvents TxtArtist As TextBox
	Friend WithEvents TxtTrackNumber As TextBox
	Public WithEvents Label1 As Label
	Friend WithEvents Label2 As Label
	Friend WithEvents Label3 As Label
	Friend WithEvents Label4 As Label
	Friend WithEvents Label5 As Label
	Friend WithEvents Label6 As Label
	Friend WithEvents Label7 As Label
	Friend WithEvents btnWrite As Button
	Friend WithEvents btnRead As Button
	Friend WithEvents ofdMain As OpenFileDialog
	Friend WithEvents cboGenre As ComboBox
	Friend WithEvents txtFileName As TextBox

	<DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As Resources.ResourceManager = New Resources.ResourceManager(GetType(FrmMain))
		TxtTitle = New TextBox()
		CmdFixTag = New Button()
		CmdClear = New Button()
		btnWrite = New Button()
		CmdAbout = New Button()
		CmdExit = New Button()
		btnRead = New Button()
		TxtComment = New TextBox()
		TxtYear = New TextBox()
		TxtAlbum = New TextBox()
		TxtArtist = New TextBox()
		TxtTrackNumber = New TextBox()
		Label1 = New Label()
		Label2 = New Label()
		Label3 = New Label()
		Label4 = New Label()
		Label5 = New Label()
		Label6 = New Label()
		Label7 = New Label()
		ofdMain = New OpenFileDialog()
		cboGenre = New ComboBox()
		txtFileName = New TextBox()
		SuspendLayout()
		'
		'txtTitle
		'
		TxtTitle.Location = New Point(100, 36)
		TxtTitle.Name = "txtTitle"
		TxtTitle.Size = New Size(209, 20)
		TxtTitle.TabIndex = 2
		TxtTitle.Text = ""
		'
		'cmdFixTag
		'
		CmdFixTag.FlatStyle = System.Windows.Forms.FlatStyle.System
		CmdFixTag.Location = New Point(244, 212)
		CmdFixTag.Name = "cmdFixTag"
		CmdFixTag.Size = New Size(65, 25)
		CmdFixTag.TabIndex = 18
		CmdFixTag.Text = "F&ix!"
		'
		'cmdClear
		'
		CmdClear.DialogResult = System.Windows.Forms.DialogResult.Cancel
		CmdClear.FlatStyle = System.Windows.Forms.FlatStyle.System
		CmdClear.Location = New Point(164, 212)
		CmdClear.Name = "cmdClear"
		CmdClear.Size = New Size(65, 25)
		CmdClear.TabIndex = 17
		CmdClear.Text = "&Clear"
		'
		'btnWrite
		'
		btnWrite.FlatStyle = System.Windows.Forms.FlatStyle.System
		btnWrite.Location = New Point(88, 212)
		btnWrite.Name = "btnWrite"
		btnWrite.Size = New Size(65, 25)
		btnWrite.TabIndex = 16
		btnWrite.Text = "&Write"
		'
		'cmdAbout
		'
		CmdAbout.FlatStyle = System.Windows.Forms.FlatStyle.System
		CmdAbout.Location = New Point(8, 244)
		CmdAbout.Name = "cmdAbout"
		CmdAbout.Size = New Size(145, 25)
		CmdAbout.TabIndex = 19
		CmdAbout.Text = "&About..."
		'
		'cmdExit
		'
		CmdExit.FlatStyle = System.Windows.Forms.FlatStyle.System
		CmdExit.Location = New Point(164, 244)
		CmdExit.Name = "cmdExit"
		CmdExit.Size = New Size(145, 25)
		CmdExit.TabIndex = 20
		CmdExit.Text = "E&xit"
		'
		'btnRead
		'
		btnRead.FlatStyle = System.Windows.Forms.FlatStyle.System
		btnRead.Location = New Point(8, 212)
		btnRead.Name = "btnRead"
		btnRead.Size = New Size(65, 25)
		btnRead.TabIndex = 15
		btnRead.Text = "&Read"
		'
		'txtComment
		'
		TxtComment.Location = New Point(100, 180)
		TxtComment.Name = "txtComment"
		TxtComment.Size = New Size(209, 20)
		TxtComment.TabIndex = 14
		TxtComment.Text = ""
		'
		'txtYear
		'
		TxtYear.Location = New Point(100, 132)
		TxtYear.Name = "txtYear"
		TxtYear.Size = New Size(33, 20)
		TxtYear.TabIndex = 10
		TxtYear.Text = ""
		'
		'txtAlbum
		'
		TxtAlbum.Location = New Point(100, 108)
		TxtAlbum.Name = "txtAlbum"
		TxtAlbum.Size = New Size(209, 20)
		TxtAlbum.TabIndex = 8
		TxtAlbum.Text = ""
		'
		'txtArtist
		'
		TxtArtist.Location = New Point(100, 84)
		TxtArtist.Name = "txtArtist"
		TxtArtist.Size = New Size(209, 20)
		TxtArtist.TabIndex = 6
		TxtArtist.Text = ""
		'
		'txtTrackNumber
		'
		TxtTrackNumber.Location = New Point(100, 60)
		TxtTrackNumber.Name = "txtTrackNumber"
		TxtTrackNumber.Size = New Size(25, 20)
		TxtTrackNumber.TabIndex = 4
		TxtTrackNumber.Text = ""
		'
		'Label1
		'
		Label1.AutoSize = True
		Label1.Location = New Point(64, 40)
		Label1.Name = "Label1"
		Label1.Size = New Size(29, 13)
		Label1.TabIndex = 1
		Label1.Text = "&Title:"
		'
		'Label2
		'
		Label2.AutoSize = True
		Label2.Location = New Point(48, 64)
		Label2.Name = "Label2"
		Label2.Size = New Size(45, 13)
		Label2.TabIndex = 3
		Label2.Text = "T&rack #:"
		'
		'Label3
		'
		Label3.AutoSize = True
		Label3.Location = New Point(60, 88)
		Label3.Name = "Label3"
		Label3.Size = New Size(33, 13)
		Label3.TabIndex = 5
		Label3.Text = "Art&ist:"
		'
		'Label4
		'
		Label4.AutoSize = True
		Label4.Location = New Point(52, 112)
		Label4.Name = "Label4"
		Label4.Size = New Size(39, 13)
		Label4.TabIndex = 7
		Label4.Text = "A&lbum:"
		'
		'Label5
		'
		Label5.AutoSize = True
		Label5.Location = New Point(60, 136)
		Label5.Name = "Label5"
		Label5.Size = New Size(31, 13)
		Label5.TabIndex = 9
		Label5.Text = "&Year:"
		'
		'Label6
		'
		Label6.AutoSize = True
		Label6.Location = New Point(52, 160)
		Label6.Name = "Label6"
		Label6.Size = New Size(39, 13)
		Label6.TabIndex = 11
		Label6.Text = "&Genre:"
		'
		'Label7
		'
		Label7.AutoSize = True
		Label7.Location = New Point(36, 184)
		Label7.Name = "Label7"
		Label7.Size = New Size(56, 13)
		Label7.TabIndex = 13
		Label7.Text = "Co&mment:"
		'
		'ofdMain
		'
		ofdMain.DefaultExt = "mp3"
		ofdMain.Filter = "MP3 Files (*.mp3)|*.mp3"
		'
		'cboGenre
		'
		cboGenre.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
		cboGenre.Location = New Point(100, 156)
		cboGenre.Name = "cboGenre"
		cboGenre.Size = New Size(209, 21)
		cboGenre.TabIndex = 12
		'
		'txtFileName
		'
		txtFileName.Location = New Point(8, 8)
		txtFileName.Name = "txtFileName"
		txtFileName.ReadOnly = True
		txtFileName.Size = New Size(300, 20)
		txtFileName.TabIndex = 0
		txtFileName.TabStop = False
		txtFileName.Text = ""
		txtFileName.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
		txtFileName.WordWrap = False
		'
		'frmMain
		'
		AutoScaleBaseSize = New Size(5, 13)
		ClientSize = New Size(316, 278)
		Controls.AddRange(New Control() {txtFileName, cboGenre, TxtTitle, CmdFixTag, CmdClear, btnWrite, CmdAbout, CmdExit, btnRead, TxtComment, TxtYear, TxtAlbum, TxtArtist, TxtTrackNumber, Label1, Label2, Label3, Label4, Label5, Label6, Label7})
		FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
		Icon = CType(resources.GetObject("$this.Icon"), Icon)
		MaximizeBox = False
		Name = "frmMain"
		StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
		Text = "NameP3"
		ResumeLayout(False)

	End Sub

#End Region

	' Formats a tag string
	Private Function FixString(ByVal strUserString As String, Optional ByVal boolFixCase As Boolean = True) As String
		Dim i, blankCount As Integer
		Dim myStr As String
		Dim finaleStr As String = ""

		For i = 1 To Len(strUserString)
			' Isolate each character
			myStr = Mid(strUserString, i, 1)

			If myStr = " " Or myStr = vbTab Then
				blankCount += 1

				If blankCount < 2 Then
					finaleStr &= " "
				End If
			Else
				blankCount = 0

				finaleStr &= myStr
			End If
		Next

		finaleStr = Trim(finaleStr)

		If boolFixCase Then
			finaleStr = StrConv(finaleStr, VbStrConv.ProperCase)
		End If

		Return finaleStr
	End Function

	' Removes illegal charaters from a file name
	Private Function FixFileName(ByVal sFName As String) As String
		Dim i As Integer
		Dim s As String = ""
		Dim c As String

		For i = 1 To Len(sFName)
			c = Mid(sFName, i, 1)
			Select Case c
				Case "\", "/", "*", "?", "|"
					s &= "_"
				Case ":"
					s &= "-"
				Case "<"
					s &= "{"
				Case ">"
					s &= "}"
				Case """"
					s &= "'"
				Case Else
					s &= c
			End Select
		Next

		Return s
	End Function

	Private Sub CmdExit_Click(ByVal sender As Object, ByVal e As EventArgs) Handles CmdExit.Click
		Close()
	End Sub

	Private Sub CmdAbout_Click(ByVal sender As Object, ByVal e As EventArgs) Handles CmdAbout.Click
		ShellAbout(Handle,
		 Application.ProductName,
		 Diagnostics.FileVersionInfo.GetVersionInfo(Reflection.Assembly.GetExecutingAssembly.Location).LegalCopyright &
		 vbNewLine & "Version " & Application.ProductVersion,
		 Icon.Handle)
	End Sub

	Private Sub CmdClear_Click(ByVal sender As Object, ByVal e As EventArgs) Handles CmdClear.Click
		' Clear all text boxes
		TxtTrackNumber.Text = String.Empty
		TxtTitle.Text = String.Empty
		TxtArtist.Text = String.Empty
		TxtAlbum.Text = String.Empty
		TxtYear.Text = String.Empty
		cboGenre.Text = "Other"
		TxtComment.Text = String.Empty
	End Sub

	' Reads in all MP3 information
	Private Sub BtnRead_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRead.Click
		Dim id3 As New UltraID3()

		' Get the MP3 file name
		Dim drOpen As DialogResult = ofdMain.ShowDialog()
		If drOpen <> DialogResult.OK Then
			Return
		End If

		' Read and parse the file
		id3.Read(ofdMain.FileName)

		' Save individual fields
		TxtTrackNumber.Text = CStr(id3.TrackNum)
		TxtTitle.Text = id3.Title
		TxtArtist.Text = id3.Artist
		TxtAlbum.Text = id3.Album
		TxtYear.Text = CStr(id3.Year)
		cboGenre.Text = id3.Genre
		TxtComment.Text = id3.Comments

		' Update file name as well
		txtFileName.Text = Path.GetFileName(id3.FileName)
	End Sub

	' We just populate the genre combo box here
	Private Sub FrmMain_Load(ByVal sender As Object, ByVal e As EventArgs) Handles MyBase.Load
		Dim i As Byte
		Dim s As String

		'cboGenre.Items.Clear()
		Do
			s = GenreInfoCollection.GetGenreName(i)
			If s = String.Empty Then Exit Do
			cboGenre.Items.Add(s)
			i += CByte(1)
		Loop

		CmdClear_Click(sender, e)
	End Sub

	' Fixes field spacing and case
	Private Sub CmdFixTag_Click(ByVal sender As Object, ByVal e As EventArgs) Handles CmdFixTag.Click
		' Fix all text
		TxtTrackNumber.Text = CStr(Val(TxtTrackNumber.Text))
		TxtTitle.Text = FixString(TxtTitle.Text)
		TxtArtist.Text = FixString(TxtArtist.Text)
		TxtAlbum.Text = FixString(TxtAlbum.Text)
		TxtYear.Text = CStr(Val(TxtYear.Text))
		TxtComment.Text = FixString(TxtComment.Text, False)

		' Fill empty ones with defaults ;)
		If Val(TxtTrackNumber.Text) < 1 Then
			TxtTrackNumber.Text = CStr(1)
		End If

		If TxtTitle.Text = String.Empty Then
			TxtTitle.Text = "Untitled"
		End If

		If TxtArtist.Text = String.Empty Then
			TxtArtist.Text = "Unknown"
		End If

		If TxtAlbum.Text = String.Empty Then
			TxtAlbum.Text = "Unknown"
		End If

		If Val(TxtYear.Text) < 1 Then
			TxtYear.Text = CStr(Year(Today))
		End If

		If TxtComment.Text = String.Empty Then
			TxtComment.Text = "mailto: v_2samg@hotmail.com"
		End If
	End Sub

	' Save all information to the MP3 file and also renames it
	Private Sub BtnWrite_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnWrite.Click
		Dim sFN As String

		If ofdMain.FileName = String.Empty Then Exit Sub

		' Build the new file name
		sFN = Path.GetDirectoryName(ofdMain.FileName) &
		 Path.DirectorySeparatorChar &
		 FixFileName(TxtArtist.Text & " - " & TxtAlbum.Text & " - " & TxtTitle.Text) &
		 Path.GetExtension(ofdMain.FileName)

		' Lets rename the file first
		Try
			Rename(ofdMain.FileName, sFN)
		Catch exp As Exception
			MsgBox(exp.Message, MsgBoxStyle.Critical)
		End Try

		' Update file name
		ofdMain.FileName = sFN
		txtFileName.Text = Path.GetFileName(sFN)

		' Read in first to get the defaults
		Dim id3 As New UltraID3()
		id3.Read(sFN)

		' Save individual fields
		Try
			id3.TrackNum = CShort(TxtTrackNumber.Text)
			id3.Title = TxtTitle.Text
			id3.Artist = TxtArtist.Text
			id3.Album = TxtAlbum.Text
			id3.Year = CShort(TxtYear.Text)
			id3.Genre = cboGenre.Text
			id3.Comments = TxtComment.Text
		Catch exp As Exception
			MsgBox(exp.Message, MsgBoxStyle.Critical)
		End Try

		' Save to disk
		Try
			id3.Write()
		Catch exp As Exception
			MsgBox(exp.Message, MsgBoxStyle.Critical)
		End Try
	End Sub
End Class
