Imports System.Diagnostics
Imports System.Drawing
Imports System.Globalization
Imports System.Linq
Imports System.Windows.Forms

'  FlexibleMessageBox – A flexible replacement for the .NET MessageBox
' * 
' *  Author:         Jörg Reichert (public@jreichert.de)
' *  Version:        1.3
' *  Published at:   http://www.codeproject.com/Articles/601900/FlexibleMessageBox
' *  
' ************************************************************************************************************
' * Features:
' *  - It can be simply used instead of MessageBox since all important static "Show"-Functions are supported
' *  - It is small, only one source file, which could be added easily to each solution 
' *  - It can be resized and the content is correctly word-wrapped
' *  - It tries to auto-size the width to show the longest text row
' *  - It never exceeds the current desktop working area
' *  - It displays a vertical scrollbar when needed
' *  - It does support hyperlinks in text
' * 
' *  Because the interface is identical to MessageBox, you can add this single source file to your project 
' *  and use the FlexibleMessageBox almost everywhere you use a standard MessageBox. 
' *  The goal was NOT to produce as many features as possible but to provide a simple replacement to fit my 
' *  own needs. Feel free to add additional features on your own, but please left my credits in this class.
' * 
' ************************************************************************************************************
' * Usage examples:
' * 
' *  FlexibleMessageBox.Show("Just a text");
' * 
' *  FlexibleMessageBox.Show("A text", 
' *                          "A caption"); 
' *  
' *  FlexibleMessageBox.Show("Some text with a link: www.google.com", 
' *                          "Some caption",
' *                          MessageBoxButtons.AbortRetryIgnore, 
' *                          MessageBoxIcon.Information,
' *                          MessageBoxDefaultButton.Button2);
' *  
' *  var dialogResult = FlexibleMessageBox.Show("Do you know the answer to life the universe and everything?", 
' *                                             "One short question",
' *                                             MessageBoxButtons.YesNo);     
' * 
' ************************************************************************************************************
' *  THE SOFTWARE IS PROVIDED BY THE AUTHOR "AS IS", WITHOUT WARRANTY
' *  OF ANY KIND, EXPRESS OR IMPLIED. IN NO EVENT SHALL THE AUTHOR BE
' *  LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY ARISING FROM,
' *  OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OF THIS
' *  SOFTWARE.
' *  
' ************************************************************************************************************
' * History:
' * Made the box not resizable by Furuchi
' *
' *  Version 1.3 - 11.March 2014
' *   - Converted to VB.NET
' *  Version 1.2 - 10.August 2013
' *   - Do not ShowInTaskbar anymore (original MessageBox is also hidden in taskbar)
' *   - Added handling for Escape-Button
' *   - Adapted top right close button (red X) to behave like MessageBox (but hidden instead of deactivated)
' * 
' *  Version 1.1 - 14.June 2013
' *   - Some Refactoring
' *   - Added internal form class
' *   - Added missing code comments, etc.
' *  
' *  Version 1.0 - 15.April 2013
' *   - Initial Version
' *  
'

Public Class FlexibleMessageBox
#Region "Public statics"

    Public Shared SIZEGRIP As SizeGripStyle = SizeGripStyle.Show 'If we make box non resizable, size grip need Not be shown
    Public Shared BORDERSTYLE As FormBorderStyle = FormBorderStyle.FixedDialog 'Default Is resizable

    ''' <summary>
    ''' Defines the maximum width for all FlexibleMessageBox instances in percent of the working area.
    ''' 
    ''' Allowed values are 0.2 - 1.0 where: 
    ''' 0.2 means:  The FlexibleMessageBox can be at most half as wide as the working area.
    ''' 1.0 means:  The FlexibleMessageBox can be as wide as the working area.
    ''' 
    ''' Default is: 70% of the working area width.
    ''' </summary>
    ''' 
    Public Shared MAX_WIDTH_FACTOR As Double = 0.5

    ''' <summary>
    ''' Defines the maximum height for all FlexibleMessageBox instances in percent of the working area.
    ''' 
    ''' Allowed values are 0.2 - 1.0 where: 
    ''' 0.2 means:  The FlexibleMessageBox can be at most half as high as the working area.
    ''' 1.0 means:  The FlexibleMessageBox can be as high as the working area.
    ''' 
    ''' Default is: 90% of the working area height.
    ''' </summary>
    Public Shared MAX_HEIGHT_FACTOR As Double = 0.5

    ''' <summary>
    ''' Defines the font for all FlexibleMessageBox instances.
    ''' 
    ''' Default is: SystemFonts.MessageBoxFont
    ''' </summary>    
    Public Shared mbFont As Font = New System.Drawing.Font("Times New Roman", 10.0F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CByte(0))
    'Public Shared mbFont As Font = SystemFonts.MessageBoxFont

#End Region

#Region "Public show functions"

    ''' <summary>
    ''' Shows the specified message box.
    ''' </summary>
    ''' <param name="text">The text.</param>
    ''' <returns>The dialog result.</returns>
    Public Shared Function Show(ByVal text As String) As DialogResult
        Return FlexibleMessageBoxForm.Show(Nothing, text, String.Empty, MessageBoxButtons.OK, MessageBoxIcon.None, MessageBoxDefaultButton.Button1)
    End Function

    ''' <summary>
    ''' Shows the specified message box.
    ''' </summary>
    ''' <param name="owner">The owner.</param>
    ''' <param name="text">The text.</param>
    ''' <returns>The dialog result.</returns>
    Public Shared Function Show(ByVal owner As IWin32Window, ByVal text As String) As DialogResult
        Return FlexibleMessageBoxForm.Show(owner, text, String.Empty, MessageBoxButtons.OK, MessageBoxIcon.None, MessageBoxDefaultButton.Button1)
    End Function

    ''' <summary>
    ''' Shows the specified message box.
    ''' </summary>
    ''' <param name="text">The text.</param>
    ''' <param name="caption">The caption.</param>
    ''' <returns>The dialog result.</returns>
    Public Shared Function Show(ByVal text As String, ByVal caption As String) As DialogResult
        Return FlexibleMessageBoxForm.Show(Nothing, text, caption, MessageBoxButtons.OK, MessageBoxIcon.None, MessageBoxDefaultButton.Button1)
    End Function

    ''' <summary>
    ''' Shows the specified message box.
    ''' </summary>
    ''' <param name="owner">The owner.</param>
    ''' <param name="text">The text.</param>
    ''' <param name="caption">The caption.</param>
    ''' <returns>The dialog result.</returns>
    Public Shared Function Show(ByVal owner As IWin32Window, ByVal text As String, ByVal caption As String) As DialogResult
        Return FlexibleMessageBoxForm.Show(owner, text, caption, MessageBoxButtons.OK, MessageBoxIcon.None, MessageBoxDefaultButton.Button1)
    End Function

    ''' <summary>
    ''' Shows the specified message box.
    ''' </summary>
    ''' <param name="text">The text.</param>
    ''' <param name="caption">The caption.</param>
    ''' <param name="buttons">The buttons.</param>
    ''' <returns>The dialog result.</returns>
    Public Shared Function Show(ByVal text As String, ByVal caption As String, ByVal buttons As MessageBoxButtons) As DialogResult
        Return FlexibleMessageBoxForm.Show(Nothing, text, caption, buttons, MessageBoxIcon.None, MessageBoxDefaultButton.Button1)
    End Function

    ''' <summary>
    ''' Shows the specified message box.
    ''' </summary>
    ''' <param name="owner">The owner.</param>
    ''' <param name="text">The text.</param>
    ''' <param name="caption">The caption.</param>
    ''' <param name="buttons">The buttons.</param>
    ''' <returns>The dialog result.</returns>
    Public Shared Function Show(ByVal owner As IWin32Window, ByVal text As String, ByVal caption As String, ByVal buttons As MessageBoxButtons) As DialogResult
        Return FlexibleMessageBoxForm.Show(owner, text, caption, buttons, MessageBoxIcon.None, MessageBoxDefaultButton.Button1)
    End Function

    ''' <summary>
    ''' Shows the specified message box.
    ''' </summary>
    ''' <param name="text">The text.</param>
    ''' <param name="caption">The caption.</param>
    ''' <param name="buttons">The buttons.</param>
    ''' <param name="icon">The icon.</param>
    ''' <returns></returns>
    Public Shared Function Show(ByVal text As String, ByVal caption As String, ByVal buttons As MessageBoxButtons, ByVal icon As MessageBoxIcon) As DialogResult
        Return FlexibleMessageBoxForm.Show(Nothing, text, caption, buttons, icon, MessageBoxDefaultButton.Button1)
    End Function

    ''' <summary>
    ''' Shows the specified message box.
    ''' </summary>
    ''' <param name="owner">The owner.</param>
    ''' <param name="text">The text.</param>
    ''' <param name="caption">The caption.</param>
    ''' <param name="buttons">The buttons.</param>
    ''' <param name="icon">The icon.</param>
    ''' <returns>The dialog result.</returns>
    Public Shared Function Show(ByVal owner As IWin32Window, ByVal text As String, ByVal caption As String, ByVal buttons As MessageBoxButtons, ByVal icon As MessageBoxIcon) As DialogResult
        Return FlexibleMessageBoxForm.Show(owner, text, caption, buttons, icon, MessageBoxDefaultButton.Button1)
    End Function

    ''' <summary>
    ''' Shows the specified message box.
    ''' </summary>
    ''' <param name="text">The text.</param>
    ''' <param name="caption">The caption.</param>
    ''' <param name="buttons">The buttons.</param>
    ''' <param name="icon">The icon.</param>
    ''' <param name="defaultButton">The default button.</param>
    ''' <returns>The dialog result.</returns>
    Public Shared Function Show(ByVal text As String, ByVal caption As String, ByVal buttons As MessageBoxButtons, ByVal icon As MessageBoxIcon, ByVal defaultButton As MessageBoxDefaultButton) As DialogResult
        Return FlexibleMessageBoxForm.Show(Nothing, text, caption, buttons, icon, defaultButton)
    End Function

    ''' <summary>
    ''' Shows the specified message box.
    ''' </summary>
    ''' <param name="owner">The owner.</param>
    ''' <param name="text">The text.</param>
    ''' <param name="caption">The caption.</param>
    ''' <param name="buttons">The buttons.</param>
    ''' <param name="icon">The icon.</param>
    ''' <param name="defaultButton">The default button.</param>
    ''' <returns>The dialog result.</returns>
    Public Shared Function Show(ByVal owner As IWin32Window, ByVal text As String, ByVal caption As String, ByVal buttons As MessageBoxButtons, ByVal icon As MessageBoxIcon, ByVal defaultButton As MessageBoxDefaultButton) As DialogResult
        Return FlexibleMessageBoxForm.Show(owner, text, caption, buttons, icon, defaultButton)
    End Function

#End Region

#Region "Internal form class"

    ''' <summary>
    ''' The form to show the customized message box.
    ''' It is defined as an internal class to keep the public interface of the FlexibleMessageBox clean.
    ''' </summary>
    Private Class FlexibleMessageBoxForm
        Inherits Form
#Region "Form-Designer generated code"

        ''' <summary>
        ''' Erforderliche Designervariable.
        ''' </summary>
        Private components As System.ComponentModel.IContainer = Nothing

        ''' <summary>
        ''' Verwendete Ressourcen bereinigen.
        ''' </summary>
        ''' <param name="disposing">True, wenn verwaltete Ressourcen gelöscht werden sollen; andernfalls False.</param>
        Protected Overrides Sub Dispose(ByVal disposing As Boolean)
            If disposing AndAlso (components IsNot Nothing) Then
                components.Dispose()
            End If
            MyBase.Dispose(disposing)
        End Sub

        ''' <summary>
        ''' Erforderliche Methode für die Designerunterstützung.
        ''' Der Inhalt der Methode darf nicht mit dem Code-Editor geändert werden.
        ''' </summary>
        Private Sub InitializeComponent()
            Me.components = New System.ComponentModel.Container()
            Me.button1 = New System.Windows.Forms.Button()
            Me.richTextBoxMessage = New System.Windows.Forms.RichTextBox()
            Me.FlexibleMessageBoxFormBindingSource = New System.Windows.Forms.BindingSource(Me.components)
            Me.panel1 = New System.Windows.Forms.Panel()
            Me.pictureBoxForIcon = New System.Windows.Forms.PictureBox()
            Me.button2 = New System.Windows.Forms.Button()
            Me.button3 = New System.Windows.Forms.Button()
            DirectCast(Me.FlexibleMessageBoxFormBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
            Me.panel1.SuspendLayout()
            DirectCast(Me.pictureBoxForIcon, System.ComponentModel.ISupportInitialize).BeginInit()
            Me.SuspendLayout()
            ' 
            ' button1
            ' 
            Me.button1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
            Me.button1.AutoSize = True
            Me.button1.DialogResult = System.Windows.Forms.DialogResult.OK
            Me.button1.Location = New System.Drawing.Point(11, 67)
            Me.button1.MinimumSize = New System.Drawing.Size(0, 24)
            Me.button1.Name = "button1"
            Me.button1.Size = New System.Drawing.Size(75, 24)
            Me.button1.TabIndex = 2
            Me.button1.Text = "OK"
            Me.button1.UseVisualStyleBackColor = True
            Me.button1.Visible = False
            ' 
            ' richTextBoxMessage
            ' 
            Me.richTextBoxMessage.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) Or System.Windows.Forms.AnchorStyles.Left) Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
            Me.richTextBoxMessage.BackColor = System.Drawing.Color.White
            Me.richTextBoxMessage.BorderStyle = System.Windows.Forms.BorderStyle.None
            Me.richTextBoxMessage.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.FlexibleMessageBoxFormBindingSource, "MessageText", True, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged))
            Me.richTextBoxMessage.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CByte(0))
            Me.richTextBoxMessage.Location = New System.Drawing.Point(50, 16)
            Me.richTextBoxMessage.Margin = New System.Windows.Forms.Padding(0)
            Me.richTextBoxMessage.Name = "richTextBoxMessage"
            Me.richTextBoxMessage.[ReadOnly] = True
            Me.richTextBoxMessage.ScrollBars = System.Windows.Forms.RichTextBoxScrollBars.Vertical
            Me.richTextBoxMessage.Size = New System.Drawing.Size(200, 30)
            Me.richTextBoxMessage.TabIndex = 0
            Me.richTextBoxMessage.Text = "<Message>"
            AddHandler Me.richTextBoxMessage.LinkClicked, New System.Windows.Forms.LinkClickedEventHandler(AddressOf Me.richTextBoxMessage_LinkClicked)
            ' 
            ' panel1
            ' 
            Me.panel1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) Or System.Windows.Forms.AnchorStyles.Left) Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
            Me.panel1.BackColor = System.Drawing.Color.White
            Me.panel1.Controls.Add(Me.pictureBoxForIcon)
            Me.panel1.Controls.Add(Me.richTextBoxMessage)
            Me.panel1.Location = New System.Drawing.Point(0, 0) '-3, -4)
            Me.panel1.Name = "panel1"
            Me.panel1.Size = New System.Drawing.Size(268, 59)
            Me.panel1.TabIndex = 1
            ' 
            ' pictureBoxForIcon
            ' 
            Me.pictureBoxForIcon.BackColor = System.Drawing.Color.Transparent
            Me.pictureBoxForIcon.Location = New System.Drawing.Point(15, 15)
            Me.pictureBoxForIcon.Name = "pictureBoxForIcon"
            Me.pictureBoxForIcon.Size = New System.Drawing.Size(32, 32)
            Me.pictureBoxForIcon.TabIndex = 8
            Me.pictureBoxForIcon.TabStop = False
            ' 
            ' button2
            ' 
            Me.button2.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
            Me.button2.DialogResult = System.Windows.Forms.DialogResult.OK
            Me.button2.Location = New System.Drawing.Point(92, 67)
            Me.button2.MinimumSize = New System.Drawing.Size(0, 24)
            Me.button2.Name = "button2"
            Me.button2.Size = New System.Drawing.Size(75, 24)
            Me.button2.TabIndex = 3
            Me.button2.Text = "OK"
            Me.button2.UseVisualStyleBackColor = True
            Me.button2.Visible = False
            ' 
            ' button3
            ' 
            Me.button3.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
            Me.button3.AutoSize = True
            Me.button3.DialogResult = System.Windows.Forms.DialogResult.OK
            Me.button3.Location = New System.Drawing.Point(173, 67)
            Me.button3.MinimumSize = New System.Drawing.Size(0, 24)
            Me.button3.Name = "button3"
            Me.button3.Size = New System.Drawing.Size(75, 24)
            Me.button3.TabIndex = 0
            Me.button3.Text = "OK"
            Me.button3.UseVisualStyleBackColor = True
            Me.button3.Visible = False
            ' 
            ' FlexibleMessageBoxForm
            ' 
            Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0F, 13.0F)
            Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
            Me.ClientSize = New System.Drawing.Size(260, 102)
            Me.Controls.Add(Me.button3)
            Me.Controls.Add(Me.button2)
            Me.Controls.Add(Me.panel1)
            Me.Controls.Add(Me.button1)
            Me.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.FlexibleMessageBoxFormBindingSource, "CaptionText", True))
            Me.MaximizeBox = False
            Me.MinimizeBox = False
            Me.MinimumSize = New System.Drawing.Size(276, 140)
            Me.Name = "FlexibleMessageBoxForm"
            Me.ShowIcon = False
            Me.SizeGripStyle = SIZEGRIP ' defualt is System.Windows.Forms.SizeGripStyle.Show
            Me.FormBorderStyle = BORDERSTYLE ' added
            Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
            Me.Text = "<Caption>"
            AddHandler Me.Shown, New System.EventHandler(AddressOf Me.FlexibleMessageBoxForm_Shown)
            DirectCast(Me.FlexibleMessageBoxFormBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
            Me.panel1.ResumeLayout(False)
            DirectCast(Me.pictureBoxForIcon, System.ComponentModel.ISupportInitialize).EndInit()
            Me.ResumeLayout(False)
            Me.PerformLayout()

        End Sub

        Private button1 As System.Windows.Forms.Button
        Private panel1 As System.Windows.Forms.Panel
        Private pictureBoxForIcon As System.Windows.Forms.PictureBox
        Private button2 As System.Windows.Forms.Button
        Private button3 As System.Windows.Forms.Button

        Public FlexibleMessageBoxFormBindingSource As System.Windows.Forms.BindingSource
        Public richTextBoxMessage As System.Windows.Forms.RichTextBox

#End Region

#Region "Private constants"

        Private Enum BUTTON_TEXT
            OK = 0
            CANCEL
            YES
            NO
            ABORT
            RETRY
            IGNORE
        End Enum
        Private Shared ReadOnly BUTTON_TEXTS_ENGLISH As [String]() = {"OK", "Cancel", "Yes", "No", "Abort", "Retry",
         "Ignore"}
        Private Shared ReadOnly BUTTON_TEXTS_GERMAN As [String]() = {"OK", "Abbrechen", "Ja", "Nein", "Abbrechen", "Wiederholen",
         "Ignorieren"}

#End Region

#Region "Private members"

        Private _defaultButton As MessageBoxDefaultButton
        Private _visibleButtonsCount As Integer
        Private _isCultureGerman As Boolean

#End Region

#Region "Private constructor"

        ''' <summary>
        ''' Initializes a new instance of the <see cref="FlexibleMessageBoxForm"/> class.
        ''' </summary>
        Private Sub New()
            InitializeComponent()

            Me._isCultureGerman = CultureInfo.InstalledUICulture.TwoLetterISOLanguageName.Contains("de")
        End Sub

#End Region

#Region "Private helper functions"

        ''' <summary>
        ''' Gets the string rows.
        ''' </summary>
        ''' <param name="message">The message.</param>
        ''' <returns>The string rows as 1-dimensional array</returns>
        Private Shared Function GetStringRows(ByVal message As String) As String()
            If String.IsNullOrEmpty(message) Then
                Return Nothing
            End If

            Dim messageRows = message.Split(New Char() {ControlChars.Lf}, StringSplitOptions.None)
            Return messageRows
        End Function

        ''' <summary>
        ''' Gets the button text for the current language (if not german, english is used as default always)
        ''' </summary>
        ''' <param name="buttonTextIndex">Index of the button text.</param>
        ''' <returns>The button text</returns>
        Private Function GetButtonText(ByVal buttonTextIndex As BUTTON_TEXT) As String
            Return If(Me._isCultureGerman, BUTTON_TEXTS_GERMAN(Convert.ToInt32(buttonTextIndex)), BUTTON_TEXTS_ENGLISH(Convert.ToInt32(buttonTextIndex)))
        End Function

        ''' <summary>
        ''' Ensure the given working area factor in the range of  0.2 - 1.0 where: 
        ''' 
        ''' 0.2 means:  Half as large as the working area.
        ''' 1.0 means:  As large as the working area.
        ''' </summary>
        ''' <param name="workingAreaFactor">The given working area factor.</param>
        ''' <returns>The corrected given working area factor.</returns>
        Private Shared Function GetCorrectedWorkingAreaFactor(ByVal workingAreaFactor As Double) As Double
            Const MIN_FACTOR As Double = 0.2
            Const MAX_FACTOR As Double = 1.0

            If workingAreaFactor < MIN_FACTOR Then
                Return MIN_FACTOR
            End If
            If workingAreaFactor > MAX_FACTOR Then
                Return MAX_FACTOR
            End If

            Return workingAreaFactor
        End Function

        ''' <summary>
        ''' Set the dialogs start position when given. 
        ''' Otherwise center the dialog on the current screen.
        ''' </summary>
        ''' <param name="fmbForm">The FlexibleMessageBox dialog.</param>
        ''' <param name="owner">The owner.</param>
        Private Shared Sub SetDialogStartPosition(ByVal fmbForm As FlexibleMessageBoxForm, ByVal owner As IWin32Window)
            'If no owner given: Center on current screen
            If owner Is Nothing Then
                fmbForm.CenterToScreen()
                'Dim screen__1 = Screen.FromPoint(Cursor.Position)
                'fmbForm.StartPosition = FormStartPosition.Manual
                'fmbForm.Left = screen__1.Bounds.Left + screen__1.Bounds.Width \ 2 - fmbForm.Width \ 2
                'fmbForm.Top = screen__1.Bounds.Top + screen__1.Bounds.Height \ 2 - fmbForm.Height \ 2
            End If
        End Sub

        ''' <summary>
        ''' Calculate the dialogs start size (Try to auto-size width to show longest text row).
        ''' Also set the maximum dialog size. 
        ''' </summary>
        ''' <param name="fmbForm">The FlexibleMessageBox dialog.</param>
        ''' <param name="text">The text (the longest text row is used to calculate the dialog width).</param>
        Private Shared Sub SetDialogSizes(ByVal fmbForm As FlexibleMessageBoxForm, ByVal text As String)
            'Set maximum dialog size
            fmbForm.MaximumSize = New Size(Convert.ToInt32(SystemInformation.WorkingArea.Width * FlexibleMessageBoxForm.GetCorrectedWorkingAreaFactor(MAX_WIDTH_FACTOR)), Convert.ToInt32(SystemInformation.WorkingArea.Height * FlexibleMessageBoxForm.GetCorrectedWorkingAreaFactor(MAX_HEIGHT_FACTOR)))

            'Calculate dialog start size: Try to auto-size width to show longest text row
            Dim rowSize As Size
            Dim maxTextRowWidth = 0
            Dim maxTextRowHeight = 0.0F
            Dim stringRows = GetStringRows(text)
            Using graphics = fmbForm.CreateGraphics()
                maxTextRowHeight = graphics.MeasureString("}]|)", mbFont).Height * 1.131 ' 1.131 give a good approximation for the space between lines

                For Each textForRow As String In stringRows
                    rowSize = graphics.MeasureString(textForRow, mbFont, fmbForm.MaximumSize.Width).ToSize()
                    If rowSize.Width > maxTextRowWidth Then
                        maxTextRowWidth = rowSize.Width
                    End If
                Next
            End Using

            'Set dialog start size
            fmbForm.Size = New Size(maxTextRowWidth + fmbForm.Width - fmbForm.richTextBoxMessage.Width, Convert.ToInt32(maxTextRowHeight * stringRows.Length) + fmbForm.Height - fmbForm.richTextBoxMessage.Height)
        End Sub

        ''' <summary>
        ''' Set the dialogs icon. 
        ''' When no icon is used: Correct placement and width of rich text box.
        ''' </summary>
        ''' <param name="fmbForm">The FlexibleMessageBox dialog.</param>
        ''' <param name="icon">The MessageBoxIcon.</param>
        Private Shared Sub SetDialogIcon(ByVal fmbForm As FlexibleMessageBoxForm, ByVal icon As MessageBoxIcon)
            Select Case icon
                Case MessageBoxIcon.Information
                    fmbForm.pictureBoxForIcon.Image = SystemIcons.Information.ToBitmap()
                    Exit Select
                Case MessageBoxIcon.Warning
                    fmbForm.pictureBoxForIcon.Image = SystemIcons.Warning.ToBitmap()
                    Exit Select
                Case MessageBoxIcon.[Error]
                    fmbForm.pictureBoxForIcon.Image = SystemIcons.[Error].ToBitmap()
                    Exit Select
                Case MessageBoxIcon.Question
                    fmbForm.pictureBoxForIcon.Image = SystemIcons.Question.ToBitmap()
                    Exit Select
                Case Else
                    'When no icon is used: Correct placement and width of rich text box.
                    fmbForm.pictureBoxForIcon.Visible = False
                    fmbForm.richTextBoxMessage.Left -= fmbForm.pictureBoxForIcon.Width
                    fmbForm.richTextBoxMessage.Width += fmbForm.pictureBoxForIcon.Width
                    Exit Select
            End Select
        End Sub

        ''' <summary>
        ''' Set dialog buttons visibilities and texts. 
        ''' Also set a default button.
        ''' </summary>
        ''' <param name="fmbForm">The FlexibleMessageBox dialog.</param>
        ''' <param name="buttons">The buttons.</param>
        ''' <param name="defaultButton">The default button.</param>
        Private Shared Sub SetDialogButtons(ByVal fmbForm As FlexibleMessageBoxForm, ByVal buttons As MessageBoxButtons, ByVal defaultButton As MessageBoxDefaultButton)
            'Set the buttons visibilities and texts
            Select Case buttons
                Case MessageBoxButtons.AbortRetryIgnore
                    fmbForm._visibleButtonsCount = 3

                    fmbForm.button1.Visible = True
                    fmbForm.button1.Text = fmbForm.GetButtonText(BUTTON_TEXT.ABORT)
                    fmbForm.button1.DialogResult = DialogResult.Abort

                    fmbForm.button2.Visible = True
                    fmbForm.button2.Text = fmbForm.GetButtonText(BUTTON_TEXT.RETRY)
                    fmbForm.button2.DialogResult = DialogResult.Retry

                    fmbForm.button3.Visible = True
                    fmbForm.button3.Text = fmbForm.GetButtonText(BUTTON_TEXT.IGNORE)
                    fmbForm.button3.DialogResult = DialogResult.Ignore

                    fmbForm.ControlBox = False
                    Exit Select

                Case MessageBoxButtons.OKCancel
                    fmbForm._visibleButtonsCount = 2

                    fmbForm.button2.Visible = True
                    fmbForm.button2.Text = fmbForm.GetButtonText(BUTTON_TEXT.OK)
                    fmbForm.button2.DialogResult = DialogResult.OK

                    fmbForm.button3.Visible = True
                    fmbForm.button3.Text = fmbForm.GetButtonText(BUTTON_TEXT.CANCEL)
                    fmbForm.button3.DialogResult = DialogResult.Cancel

                    fmbForm.CancelButton = fmbForm.button3
                    Exit Select

                Case MessageBoxButtons.RetryCancel
                    fmbForm._visibleButtonsCount = 2

                    fmbForm.button2.Visible = True
                    fmbForm.button2.Text = fmbForm.GetButtonText(BUTTON_TEXT.RETRY)
                    fmbForm.button2.DialogResult = DialogResult.Retry

                    fmbForm.button3.Visible = True
                    fmbForm.button3.Text = fmbForm.GetButtonText(BUTTON_TEXT.CANCEL)
                    fmbForm.button3.DialogResult = DialogResult.Cancel

                    fmbForm.CancelButton = fmbForm.button3
                    Exit Select

                Case MessageBoxButtons.YesNo
                    fmbForm._visibleButtonsCount = 2

                    fmbForm.button2.Visible = True
                    fmbForm.button2.Text = fmbForm.GetButtonText(BUTTON_TEXT.YES)
                    fmbForm.button2.DialogResult = DialogResult.Yes

                    fmbForm.button3.Visible = True
                    fmbForm.button3.Text = fmbForm.GetButtonText(BUTTON_TEXT.NO)
                    fmbForm.button3.DialogResult = DialogResult.No

                    fmbForm.ControlBox = False
                    Exit Select

                Case MessageBoxButtons.YesNoCancel
                    fmbForm._visibleButtonsCount = 3

                    fmbForm.button1.Visible = True
                    fmbForm.button1.Text = fmbForm.GetButtonText(BUTTON_TEXT.YES)
                    fmbForm.button1.DialogResult = DialogResult.Yes

                    fmbForm.button2.Visible = True
                    fmbForm.button2.Text = fmbForm.GetButtonText(BUTTON_TEXT.NO)
                    fmbForm.button2.DialogResult = DialogResult.No

                    fmbForm.button3.Visible = True
                    fmbForm.button3.Text = fmbForm.GetButtonText(BUTTON_TEXT.CANCEL)
                    fmbForm.button3.DialogResult = DialogResult.Cancel

                    fmbForm.CancelButton = fmbForm.button3
                    Exit Select

                Case Else 'MessageBoxButtons.OK
                    fmbForm._visibleButtonsCount = 1
                    fmbForm.button3.Visible = True
                    fmbForm.button3.Text = fmbForm.GetButtonText(BUTTON_TEXT.OK)
                    fmbForm.button3.DialogResult = DialogResult.OK

                    fmbForm.CancelButton = fmbForm.button3
                    Exit Select
            End Select

            'Set default button (used in FlexibleMessageBoxForm_Shown)
            fmbForm._defaultButton = defaultButton
        End Sub

#End Region

#Region "Private event handlers"

        ''' <summary>
        ''' Handles the Shown event of the FlexibleMessageBoxForm control.
        ''' </summary>
        ''' <param name="sender">The source of the event.</param>
        ''' <param name="e">The <see cref="System.EventArgs"/> instance containing the event data.</param>
        Private Sub FlexibleMessageBoxForm_Shown(ByVal sender As Object, ByVal e As EventArgs)
            Dim buttonIndexToFocus As Integer = 1
            Dim buttonToFocus As Button

            'Set the default button...
            Select Case Me._defaultButton
                Case MessageBoxDefaultButton.Button2
                    buttonIndexToFocus = 2
                    Exit Select
                Case MessageBoxDefaultButton.Button3
                    buttonIndexToFocus = 3
                    Exit Select
                Case Else 'MessageBoxDefaultButton.Button1
                    buttonIndexToFocus = 1
                    Exit Select
            End Select

            If buttonIndexToFocus > Me._visibleButtonsCount Then
                buttonIndexToFocus = Me._visibleButtonsCount
            End If

            If buttonIndexToFocus = 3 Then
                buttonToFocus = Me.button3
            ElseIf buttonIndexToFocus = 2 Then
                buttonToFocus = Me.button2
            Else
                buttonToFocus = Me.button1
            End If

            buttonToFocus.Focus()
        End Sub

        ''' <summary>
        ''' Handles the LinkClicked event of the richTextBoxMessage control.
        ''' </summary>
        ''' <param name="sender">The source of the event.</param>
        ''' <param name="e">The <see cref="System.Windows.Forms.LinkClickedEventArgs"/> instance containing the event data.</param>
        Private Sub richTextBoxMessage_LinkClicked(ByVal sender As Object, ByVal e As LinkClickedEventArgs)
            Try
                Cursor.Current = Cursors.WaitCursor
                Process.Start(e.LinkText)
            Catch generatedExceptionName As Exception
                'Let the caller of FlexibleMessageBoxForm decide what to do with this exception...
                Throw
            Finally
                Cursor.Current = Cursors.[Default]
            End Try

        End Sub

#End Region

#Region "Properties (only used for binding)"

        ''' <summary>
        ''' The text that is been used for the heading.
        ''' </summary>
        Public Property CaptionText() As String
            Get
                Return m_CaptionText
            End Get
            Set(ByVal value As String)
                m_CaptionText = value
            End Set
        End Property
        Private m_CaptionText As String

        ''' <summary>
        ''' The text that is been used in the FlexibleMessageBoxForm.
        ''' </summary>
        Public Property MessageText() As String
            Get
                Return m_MessageText
            End Get
            Set(ByVal value As String)
                m_MessageText = value
            End Set
        End Property
        Private m_MessageText As String

#End Region

#Region "Public show function"

        ''' <summary>
        ''' Shows the specified message box.
        ''' </summary>
        ''' <param name="owner">The owner.</param>
        ''' <param name="text">The text.</param>
        ''' <param name="caption">The caption.</param>
        ''' <param name="buttons">The buttons.</param>
        ''' <param name="icon">The icon.</param>
        ''' <param name="defaultButton">The default button.</param>
        ''' <returns>The dialog result.</returns>
        Public Overloads Shared Function Show(ByVal owner As IWin32Window, ByVal text As String, ByVal caption As String, ByVal buttons As MessageBoxButtons, ByVal icon As MessageBoxIcon, ByVal defaultButton As MessageBoxDefaultButton) As DialogResult
            'Create a new instance of the FlexibleMessageBox form
            Dim fmbForm As FlexibleMessageBoxForm = New FlexibleMessageBoxForm()
            fmbForm.DoubleBuffered = True
            fmbForm.ShowInTaskbar = False

            'Bind the caption and the message text
            fmbForm.CaptionText = caption
            fmbForm.MessageText = text
            fmbForm.FlexibleMessageBoxFormBindingSource.DataSource = fmbForm

            'Set the buttons visibilities and texts. Also set a default button.
            SetDialogButtons(fmbForm, buttons, defaultButton)

            'Set the dialogs icon. When no icon is used: Correct placement and width of rich text box.
            SetDialogIcon(fmbForm, icon)

            'Set the font for all controls
            fmbForm.Font = mbFont
            fmbForm.richTextBoxMessage.Font = mbFont

            'Calculate the dialogs start size (Try to auto-size width to show longest text row). Also set the maximum dialog size. 
            SetDialogSizes(fmbForm, text)

            'Set the dialogs start position when given. Otherwise center the dialog on the current screen.
            'SetDialogStartPosition(fmbForm, owner)

            'Show the dialog
            Return fmbForm.ShowDialog(owner)
        End Function

#End Region
    End Class
    'class FlexibleMessageBoxForm
#End Region
End Class