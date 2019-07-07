Option Strict Off
Option Explicit On
Friend Class Form3
	Inherits System.Windows.Forms.Form
#Region "Windows フォーム デザイナによって生成されたコード"
	Public Sub New()
		MyBase.New()
		If m_vb6FormDefInstance Is Nothing Then
			If m_InitializingDefInstance Then
				m_vb6FormDefInstance = Me
			Else
				Try 
					'スタートアップ フォームについては、最初に作成されたインスタンスが既定インスタンスになります。
					If System.Reflection.Assembly.GetExecutingAssembly.EntryPoint.DeclaringType Is Me.GetType Then
						m_vb6FormDefInstance = Me
					End If
				Catch
				End Try
			End If
		End If
		'この呼び出しは、Windows フォーム デザイナで必要です。
		InitializeComponent()
	End Sub
	'Form は、コンポーネント一覧に後処理を実行するために dispose をオーバーライドします。
	Protected Overloads Overrides Sub Dispose(ByVal Disposing As Boolean)
		If Disposing Then
			If Not components Is Nothing Then
				components.Dispose()
			End If
		End If
		MyBase.Dispose(Disposing)
	End Sub
	'Windows フォーム デザイナで必要です。
	Private components As System.ComponentModel.IContainer
	Public ToolTip1 As System.Windows.Forms.ToolTip
	Public WithEvents _SpinButton1_2 As AxMSForms.AxSpinButton
	Public WithEvents _Label1_2 As System.Windows.Forms.Label
	Public WithEvents _SpinButton1_1 As AxMSForms.AxSpinButton
	Public WithEvents _Label1_1 As System.Windows.Forms.Label
	Public WithEvents _SpinButton1_0 As AxMSForms.AxSpinButton
	Public WithEvents _Label1_0 As System.Windows.Forms.Label
	Public WithEvents _Shape1_2 As System.Windows.Forms.Label
	Public WithEvents _Shape1_1 As System.Windows.Forms.Label
	Public WithEvents _Shape1_0 As System.Windows.Forms.Label
	Public WithEvents Label1 As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	Public WithEvents Shape1 As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	Public WithEvents SpinButton1 As AxSpinButtonArray.AxSpinButtonArray
	'メモ : 以下のプロシージャは Windows フォーム デザイナで必要です。
	'Windows フォーム デザイナを使って変更できます。
	'コード エディタを使って修正しないでください。
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(Form3))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.ToolTip1.Active = True
		Me._SpinButton1_2 = New AxMSForms.AxSpinButton
		Me._Label1_2 = New System.Windows.Forms.Label
		Me._SpinButton1_1 = New AxMSForms.AxSpinButton
		Me._Label1_1 = New System.Windows.Forms.Label
		Me._SpinButton1_0 = New AxMSForms.AxSpinButton
		Me._Label1_0 = New System.Windows.Forms.Label
		Me._Shape1_2 = New System.Windows.Forms.Label
		Me._Shape1_1 = New System.Windows.Forms.Label
		Me._Shape1_0 = New System.Windows.Forms.Label
		Me.Label1 = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(components)
		Me.Shape1 = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(components)
		Me.SpinButton1 = New AxSpinButtonArray.AxSpinButtonArray(components)
		CType(Me._SpinButton1_2, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me._SpinButton1_1, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me._SpinButton1_0, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.Label1, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.Shape1, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.SpinButton1, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
		Me.BackColor = System.Drawing.Color.Black
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
		Me.Text = "見る角度"
		Me.ClientSize = New System.Drawing.Size(294, 294)
		Me.Location = New System.Drawing.Point(34, 182)
		Me.ForeColor = System.Drawing.Color.White
		Me.MaximizeBox = False
		Me.MinimizeBox = False
		Me.ShowInTaskbar = False
		Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
		Me.ControlBox = True
		Me.Enabled = True
		Me.KeyPreview = False
		Me.Cursor = System.Windows.Forms.Cursors.Default
		Me.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.HelpButton = False
		Me.WindowState = System.Windows.Forms.FormWindowState.Normal
		Me.Name = "Form3"
		_SpinButton1_2.OcxState = CType(resources.GetObject("_SpinButton1_2.OcxState"), System.Windows.Forms.AxHost.State)
		Me._SpinButton1_2.Size = New System.Drawing.Size(33, 49)
		Me._SpinButton1_2.Location = New System.Drawing.Point(56, 64)
		Me._SpinButton1_2.TabIndex = 5
		Me._SpinButton1_2.Name = "_SpinButton1_2"
		Me._Label1_2.BackColor = System.Drawing.Color.Black
		Me._Label1_2.Text = "Z"
		Me._Label1_2.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 18!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
		Me._Label1_2.ForeColor = System.Drawing.Color.White
		Me._Label1_2.Size = New System.Drawing.Size(17, 25)
		Me._Label1_2.Location = New System.Drawing.Point(64, 40)
		Me._Label1_2.TabIndex = 4
		Me._Label1_2.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._Label1_2.Enabled = True
		Me._Label1_2.Cursor = System.Windows.Forms.Cursors.Default
		Me._Label1_2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Label1_2.UseMnemonic = True
		Me._Label1_2.Visible = True
		Me._Label1_2.AutoSize = False
		Me._Label1_2.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._Label1_2.Name = "_Label1_2"
		_SpinButton1_1.OcxState = CType(resources.GetObject("_SpinButton1_1.OcxState"), System.Windows.Forms.AxHost.State)
		Me._SpinButton1_1.Size = New System.Drawing.Size(33, 49)
		Me._SpinButton1_1.Location = New System.Drawing.Point(200, 64)
		Me._SpinButton1_1.TabIndex = 3
		Me._SpinButton1_1.Name = "_SpinButton1_1"
		Me._Label1_1.BackColor = System.Drawing.Color.Black
		Me._Label1_1.Text = "Y"
		Me._Label1_1.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 18!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
		Me._Label1_1.ForeColor = System.Drawing.Color.White
		Me._Label1_1.Size = New System.Drawing.Size(17, 25)
		Me._Label1_1.Location = New System.Drawing.Point(208, 40)
		Me._Label1_1.TabIndex = 2
		Me._Label1_1.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._Label1_1.Enabled = True
		Me._Label1_1.Cursor = System.Windows.Forms.Cursors.Default
		Me._Label1_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Label1_1.UseMnemonic = True
		Me._Label1_1.Visible = True
		Me._Label1_1.AutoSize = False
		Me._Label1_1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._Label1_1.Name = "_Label1_1"
		_SpinButton1_0.OcxState = CType(resources.GetObject("_SpinButton1_0.OcxState"), System.Windows.Forms.AxHost.State)
		Me._SpinButton1_0.Size = New System.Drawing.Size(33, 49)
		Me._SpinButton1_0.Location = New System.Drawing.Point(56, 208)
		Me._SpinButton1_0.TabIndex = 0
		Me._SpinButton1_0.Name = "_SpinButton1_0"
		Me._Label1_0.BackColor = System.Drawing.Color.Black
		Me._Label1_0.Text = "X"
		Me._Label1_0.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 18!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
		Me._Label1_0.ForeColor = System.Drawing.Color.White
		Me._Label1_0.Size = New System.Drawing.Size(17, 25)
		Me._Label1_0.Location = New System.Drawing.Point(64, 184)
		Me._Label1_0.TabIndex = 1
		Me._Label1_0.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._Label1_0.Enabled = True
		Me._Label1_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._Label1_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Label1_0.UseMnemonic = True
		Me._Label1_0.Visible = True
		Me._Label1_0.AutoSize = False
		Me._Label1_0.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._Label1_0.Name = "_Label1_0"
		Me._Shape1_2.Size = New System.Drawing.Size(145, 145)
		Me._Shape1_2.Location = New System.Drawing.Point(0, 0)
		Me._Shape1_2.BackColor = System.Drawing.Color.Red
		Me._Shape1_2.Visible = True
		Me._Shape1_2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me._Shape1_2.ForeColor = System.Drawing.Color.Black
		Me._Shape1_2.Text = "_Shape1_2"
		Me._Shape1_2.Name = "_Shape1_2"
		Me._Shape1_1.Size = New System.Drawing.Size(145, 145)
		Me._Shape1_1.Location = New System.Drawing.Point(144, 0)
		Me._Shape1_1.BackColor = System.Drawing.Color.Red
		Me._Shape1_1.Visible = True
		Me._Shape1_1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me._Shape1_1.ForeColor = System.Drawing.Color.Black
		Me._Shape1_1.Text = "_Shape1_1"
		Me._Shape1_1.Name = "_Shape1_1"
		Me._Shape1_0.Size = New System.Drawing.Size(145, 145)
		Me._Shape1_0.Location = New System.Drawing.Point(0, 144)
		Me._Shape1_0.BackColor = System.Drawing.Color.Red
		Me._Shape1_0.Visible = True
		Me._Shape1_0.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me._Shape1_0.ForeColor = System.Drawing.Color.Black
		Me._Shape1_0.Text = "_Shape1_0"
		Me._Shape1_0.Name = "_Shape1_0"
		Me.Controls.Add(_SpinButton1_2)
		Me.Controls.Add(_Label1_2)
		Me.Controls.Add(_SpinButton1_1)
		Me.Controls.Add(_Label1_1)
		Me.Controls.Add(_SpinButton1_0)
		Me.Controls.Add(_Label1_0)
		Me.Controls.Add(_Shape1_2)
		Me.Controls.Add(_Shape1_1)
		Me.Controls.Add(_Shape1_0)
		Me.Label1.SetIndex(_Label1_2, CType(2, Short))
		Me.Label1.SetIndex(_Label1_1, CType(1, Short))
		Me.Label1.SetIndex(_Label1_0, CType(0, Short))
		Me.Shape1.SetIndex(_Shape1_2, CType(2, Short))
		Me.Shape1.SetIndex(_Shape1_1, CType(1, Short))
		Me.Shape1.SetIndex(_Shape1_0, CType(0, Short))
		Me.SpinButton1.SetIndex(_SpinButton1_2, CType(2, Short))
		Me.SpinButton1.SetIndex(_SpinButton1_1, CType(1, Short))
		Me.SpinButton1.SetIndex(_SpinButton1_0, CType(0, Short))
		CType(Me.SpinButton1, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.Shape1, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.Label1, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me._SpinButton1_0, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me._SpinButton1_1, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me._SpinButton1_2, System.ComponentModel.ISupportInitialize).EndInit()
	End Sub
#End Region 
#Region "アップグレード ウィザードのサポート コード"
	Private Shared m_vb6FormDefInstance As Form3
	Private Shared m_InitializingDefInstance As Boolean
	Public Shared Property DefInstance() As Form3
		Get
			If m_vb6FormDefInstance Is Nothing OrElse m_vb6FormDefInstance.IsDisposed Then
				m_InitializingDefInstance = True
				m_vb6FormDefInstance = New Form3()
				m_InitializingDefInstance = False
			End If
			DefInstance = m_vb6FormDefInstance
		End Get
		Set
			m_vb6FormDefInstance = Value
		End Set
	End Property
#End Region 
	Dim xs(2) As Short
	Dim ys(2) As Short
	Dim xe(2) As Short
	Dim ye(2) As Short
    Dim thrd(2, 2) As Double
	
	Private Sub Form3_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        Dim r As Double
        Dim Index As Double
		For Index = 0 To 2
            r = Int(VB6.PixelsToTwipsX(Shape1(Index).Width) / 2)
            xs(Index) = r + VB6.PixelsToTwipsX(Shape1(Index).Left)
            ys(Index) = r + VB6.PixelsToTwipsY(Shape1(Index).Top)
		Next Index
	End Sub
	
	Private Sub SpinButton1_Change(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles SpinButton1.Change
		Dim Index As Short = SpinButton1.GetIndex(eventSender)
        Dim a As Integer
		If SpinButton1(Index).Value = 1001 Then SpinButton1(Index).Value = 1
		If SpinButton1(Index).Value = 0 Then SpinButton1(Index).Value = 1000
		xe(Index) = System.Math.Sin(SpinButton1(Index).Value / 500 * 3.1415926535) * xs(Index)
		ye(Index) = System.Math.Cos(SpinButton1(Index).Value / 500 * 3.1415926535) * ys(Index)
        Cls()
		For a = 0 To 2
            'UPGRADE_ISSUE: Form メソッド Form3.Line はアップグレードされませんでした。 詳細については次のリンクをクリックしてください : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
			Me.Line (xs(a), ys(a)) - (xe(a), ye(a))
		Next a
        Dim th(2) As Double
        Dim cosine(2) As Double
        Dim sine(2) As Double
		For a = 0 To 2
            th(a) = SpinButton1(a).Value / 1000 * 3.1415
            cosine(a) = System.Math.Cos(th(a))
            sine(a) = System.Math.Sin(th(a))
		Next a
		For a = 0 To 2
            thrd(a, 1) = thrd(a, 1) * cosine(2) - thrd(a, 2) * sine(2)
            thrd(a, 2) = thrd(a, 1) * sine(2) + thrd(a, 2) * cosine(2)
            thrd(a, 0) = thrd(a, 0) * cosine(1) - thrd(a, 2) * sine(1)
            thrd(a, 2) = thrd(a, 0) * sine(1) + thrd(a, 2) * cosine(1)
            thrd(a, 0) = thrd(a, 0) * cosine(0) - thrd(a, 1) * sine(0)
            thrd(a, 1) = thrd(a, 0) * sine(0) + thrd(a, 1) * cosine(0)
		Next a
		Call Form1.DefInstance.移動thrd(thrd(0, 0), thrd(0, 1), thrd(0, 2), thrd(1, 0), thrd(1, 1), thrd(1, 2), thrd(2, 0), thrd(2, 1), thrd(2, 2))
		Call Form1.DefInstance.表示()
	End Sub
End Class