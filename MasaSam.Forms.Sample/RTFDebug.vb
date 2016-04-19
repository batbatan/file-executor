Public Class cRTFDebug
    Inherits System.Windows.Forms.Form

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
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents rtfOut As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.Label2 = New System.Windows.Forms.Label
        Me.rtfOut = New System.Windows.Forms.TextBox
        Me.SuspendLayout()
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(12, 4)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(256, 16)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "RTF the control rendered:"
        '
        'rtfOut
        '
        Me.rtfOut.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.rtfOut.AutoSize = False
        Me.rtfOut.Font = New System.Drawing.Font("Courier New", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rtfOut.Location = New System.Drawing.Point(8, 24)
        Me.rtfOut.Multiline = True
        Me.rtfOut.Name = "rtfOut"
        Me.rtfOut.ReadOnly = True
        Me.rtfOut.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.rtfOut.Size = New System.Drawing.Size(392, 200)
        Me.rtfOut.TabIndex = 3
        Me.rtfOut.Text = ""
        '
        'cRTFDebug
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(408, 238)
        Me.ControlBox = False
        Me.Controls.Add(Me.rtfOut)
        Me.Controls.Add(Me.Label2)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.SizableToolWindow
        Me.Name = "cRTFDebug"
        Me.ShowInTaskbar = False
        Me.Text = "RTF View"
        Me.ResumeLayout(False)

    End Sub

#End Region

End Class
