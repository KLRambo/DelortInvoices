
#Region " Options "

Option Explicit On
Option Strict On

#End Region

#Region " Imports "

Imports System.Configuration.ConfigurationSettings
Imports System.Configuration

#End Region

Public Class frmAbout
    Inherits frmBaseWindow

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
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents grpMain As System.Windows.Forms.GroupBox
    Friend WithEvents lblLicTo As System.Windows.Forms.Label
    Friend WithEvents lblIntro As System.Windows.Forms.Label
    Friend WithEvents picAbout As System.Windows.Forms.PictureBox
    Friend WithEvents grpFooter As System.Windows.Forms.GroupBox
    Friend WithEvents lblCopyright As System.Windows.Forms.Label
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents lblAuthor As System.Windows.Forms.LinkLabel
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmAbout))
        Me.grpMain = New System.Windows.Forms.GroupBox()
        Me.lblLicTo = New System.Windows.Forms.Label()
        Me.lblIntro = New System.Windows.Forms.Label()
        Me.picAbout = New System.Windows.Forms.PictureBox()
        Me.grpFooter = New System.Windows.Forms.GroupBox()
        Me.lblCopyright = New System.Windows.Forms.Label()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.lblAuthor = New System.Windows.Forms.LinkLabel()
        Me.grpMain.SuspendLayout()
        CType(Me.picAbout, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.grpFooter.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'grpMain
        '
        Me.grpMain.BackColor = System.Drawing.Color.Transparent
        Me.grpMain.Controls.Add(Me.lblLicTo)
        Me.grpMain.Location = New System.Drawing.Point(183, 10)
        Me.grpMain.Name = "grpMain"
        Me.grpMain.Size = New System.Drawing.Size(193, 60)
        Me.grpMain.TabIndex = 0
        Me.grpMain.TabStop = False
        '
        'lblLicTo
        '
        Me.lblLicTo.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblLicTo.Location = New System.Drawing.Point(9, 15)
        Me.lblLicTo.Name = "lblLicTo"
        Me.lblLicTo.Size = New System.Drawing.Size(175, 41)
        Me.lblLicTo.TabIndex = 0
        '
        'lblIntro
        '
        Me.lblIntro.BackColor = System.Drawing.Color.Transparent
        Me.lblIntro.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblIntro.Location = New System.Drawing.Point(183, 0)
        Me.lblIntro.Name = "lblIntro"
        Me.lblIntro.Size = New System.Drawing.Size(184, 16)
        Me.lblIntro.TabIndex = 1
        Me.lblIntro.Text = "This Product is licensed to:"
        '
        'picAbout
        '
        Me.picAbout.Image = CType(resources.GetObject("picAbout.Image"), System.Drawing.Image)
        Me.picAbout.Location = New System.Drawing.Point(0, 6)
        Me.picAbout.Name = "picAbout"
        Me.picAbout.Size = New System.Drawing.Size(176, 106)
        Me.picAbout.TabIndex = 2
        Me.picAbout.TabStop = False
        '
        'grpFooter
        '
        Me.grpFooter.BackColor = System.Drawing.Color.Transparent
        Me.grpFooter.Controls.Add(Me.lblCopyright)
        Me.grpFooter.Location = New System.Drawing.Point(183, 71)
        Me.grpFooter.Name = "grpFooter"
        Me.grpFooter.Size = New System.Drawing.Size(193, 81)
        Me.grpFooter.TabIndex = 3
        Me.grpFooter.TabStop = False
        '
        'lblCopyright
        '
        Me.lblCopyright.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCopyright.Location = New System.Drawing.Point(7, 15)
        Me.lblCopyright.Name = "lblCopyright"
        Me.lblCopyright.Size = New System.Drawing.Size(177, 57)
        Me.lblCopyright.TabIndex = 0
        '
        'GroupBox1
        '
        Me.GroupBox1.BackColor = System.Drawing.Color.Transparent
        Me.GroupBox1.Controls.Add(Me.lblAuthor)
        Me.GroupBox1.Location = New System.Drawing.Point(8, 114)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(168, 38)
        Me.GroupBox1.TabIndex = 4
        Me.GroupBox1.TabStop = False
        '
        'lblAuthor
        '
        Me.lblAuthor.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblAuthor.Location = New System.Drawing.Point(8, 16)
        Me.lblAuthor.Name = "lblAuthor"
        Me.lblAuthor.Size = New System.Drawing.Size(148, 14)
        Me.lblAuthor.TabIndex = 0
        Me.lblAuthor.TabStop = True
        Me.lblAuthor.Text = "Written By Kevin Rambo"
        '
        'frmAbout
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(7, 15)
        Me.ClientSize = New System.Drawing.Size(382, 156)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.grpFooter)
        Me.Controls.Add(Me.picAbout)
        Me.Controls.Add(Me.lblIntro)
        Me.Controls.Add(Me.grpMain)
        Me.Font = New System.Drawing.Font("Verdana", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmAbout"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "About Delort Invoices"
        Me.grpMain.ResumeLayout(False)
        CType(Me.picAbout, System.ComponentModel.ISupportInitialize).EndInit()
        Me.grpFooter.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

#Region " Control event handlers "

    Private Sub frmAbout_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Me.lblCopyright.Text = "This program is protected by copyright law. Unauthorized reproduction or distribution is prohibited."
        Me.lblLicTo.Text = "John Delort" & vbCrLf & "Delort and Sons Homeowner Services Version " & My.Application.Info.Version.ToString


    End Sub

    Private Sub lblAuthor_LinkClicked(ByVal sender As Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles lblAuthor.LinkClicked
        Try
            Dim EmailLink As String
            EmailLink = "mailto:" & _
                      "KevinRambo@Hotmail.com"

        System.Diagnostics.Process.Start(EmailLink)

        Catch ex As Exception
            'do nothing
        End Try

    End Sub

#End Region

End Class
