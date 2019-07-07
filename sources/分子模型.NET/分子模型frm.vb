Option Strict Off
Option Explicit On
Friend Class Form1
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
	Public WithEvents ProgressBar1 As AxMSComctlLib.AxProgressBar
	Public WithEvents 入力 As System.Windows.Forms.MenuItem
	Public MainMenu1 As System.Windows.Forms.MainMenu
	'メモ : 以下のプロシージャは Windows フォーム デザイナで必要です。
	'Windows フォーム デザイナを使って変更できます。
	'コード エディタを使って修正しないでください。
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(Form1))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.ToolTip1.Active = True
		Me.ProgressBar1 = New AxMSComctlLib.AxProgressBar
		Me.MainMenu1 = New System.Windows.Forms.MainMenu
		Me.入力 = New System.Windows.Forms.MenuItem
		CType(Me.ProgressBar1, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
		Me.BackColor = System.Drawing.Color.Black
		Me.Text = "分子模型"
		Me.ClientSize = New System.Drawing.Size(611, 414)
		Me.Location = New System.Drawing.Point(376, 202)
		Me.AutoScaleBaseSize = New System.Drawing.Size(5, 12)
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Sizable
		Me.ControlBox = True
		Me.Enabled = True
		Me.KeyPreview = False
		Me.MaximizeBox = True
		Me.MinimizeBox = True
		Me.Cursor = System.Windows.Forms.Cursors.Default
		Me.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.ShowInTaskbar = True
		Me.HelpButton = False
		Me.WindowState = System.Windows.Forms.FormWindowState.Normal
		Me.Name = "Form1"
		ProgressBar1.OcxState = CType(resources.GetObject("ProgressBar1.OcxState"), System.Windows.Forms.AxHost.State)
		Me.ProgressBar1.Size = New System.Drawing.Size(611, 9)
		Me.ProgressBar1.Location = New System.Drawing.Point(0, 400)
		Me.ProgressBar1.TabIndex = 0
		Me.ProgressBar1.Name = "ProgressBar1"
		Me.入力.Text = "入力..."
		Me.入力.Checked = False
		Me.入力.Enabled = True
		Me.入力.Visible = True
		Me.入力.MDIList = False
		Me.Controls.Add(ProgressBar1)
		CType(Me.ProgressBar1, System.ComponentModel.ISupportInitialize).EndInit()
		Me.入力.Index = 0
		MainMenu1.MenuItems.AddRange(New System.Windows.Forms.MenuItem(){Me.入力})
		Me.Menu = MainMenu1
	End Sub
#End Region 
#Region "アップグレード ウィザードのサポート コード"
	Private Shared m_vb6FormDefInstance As Form1
	Private Shared m_InitializingDefInstance As Boolean
	Public Shared Property DefInstance() As Form1
		Get
			If m_vb6FormDefInstance Is Nothing OrElse m_vb6FormDefInstance.IsDisposed Then
				m_InitializingDefInstance = True
				m_vb6FormDefInstance = New Form1()
				m_InitializingDefInstance = False
			End If
			DefInstance = m_vb6FormDefInstance
		End Get
		Set
			m_vb6FormDefInstance = Value
		End Set
	End Property
#End Region 
	Dim atodat(99, 6) As Double '原子の番号・三次元中心座標/色光三原色/半径,
    Public atonum As Double '原子の数
    Dim two(2, 1) As Double
    Dim thrd(2, 2) As Double
    Dim win(500, 500, 2) As Double
    Dim ysee, xsee, zsee As Double '視点
	Dim pointcount As Single
	Dim pc(0) As Single
	
	Private Sub Form1_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        Dim b As Double
        Dim a As Double
		Form3.DefInstance.Show()
        xsee = -1000
        ysee = -1000
        zsee = -1000
        two(0, 0) = -2 * System.Math.Sqrt(5) / 5
        two(0, 1) = -System.Math.Sqrt(5) / 5
        two(1, 0) = 2 * System.Math.Sqrt(5) / 5
        two(1, 1) = 2 * System.Math.Sqrt(5) / 5
        two(2, 0) = 0
        two(2, 1) = 1
		atodat(0, 0) = 100
		atodat(0, 4) = 255
		atodat(0, 5) = 20
		atodat(0, 6) = 20
		For a = 0 To 500
			For b = 0 To 500
                win(a, b, 1) = 10000
			Next b
		Next a
	End Sub
	
	'UPGRADE_WARNING: Form イベント Form1.Unload には新しい動作が含まれます。 詳細については次のリンクをクリックしてください : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2065"'
	Private Sub Form1_Closed(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Closed
		End
	End Sub
	
	Public Sub 入力_Popup(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles 入力.Popup
		入力_Click(eventSender, eventArgs)
	End Sub
	Public Sub 入力_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles 入力.Click
		Form2.DefInstance.Show()
	End Sub
	
	Public Sub 表示()
        Dim zo As Double
        Dim yo As Double
        Dim xo As Double
        Dim blu As Double
        Dim gre As Double
        Dim red As Double
        Dim yy As Double
        Dim xx As Double
        Dim bb As Double
        Dim b As Double
        Dim xy As Double
        Dim zz As Double
        Dim aa As Double
        Dim a As Double
        Dim stbl As Double
        Dim stgr As Double
        Dim strd As Double
        Dim stz As Double
        Dim sty As Double
        Dim stx As Double
        Dim r As Double
        Dim c As Double
        Dim old(400, 2) As Double
        Dim old2(2) As Double
        If atonum > -1 Then
            For c = 0 To atonum
                r = atodat(c, 0)
                stx = atodat(c, 1)
                sty = atodat(c, 2)
                stz = atodat(c, 3)
                strd = atodat(c, 4)
                stgr = atodat(c, 5)
                stbl = atodat(c, 6)
                For a = 0 To 2 Step 0.05
                    aa = a * 3.1415926535
                    zz = System.Math.Sin(aa) * r
                    xy = System.Math.Cos(aa) * r
                    For b = 0 To 1 Step 0.05
                        bb = b * 3.1415926535
                        xx = System.Math.Cos(bb) * xy
                        yy = System.Math.Sin(bb) * xy
                        If a > 0 And b > 0 Then
                            red = ((75 / 256 - a * 150 / 255) + strd / 256) * 256
                            If red > 255 Then red = 255
                            If red < 0 Then red = 0
                            gre = ((75 / 256 - a * 150 / 255) + stgr / 256) * 256
                            If gre > 255 Then gre = 255
                            If gre < 0 Then gre = 0
                            blu = ((75 / 256 - a * 150 / 255) + stbl / 256) * 256
                            If blu > 255 Then blu = 255
                            If blu < 0 Then blu = 0
                            Call draw3dline(xx + stx, yy + sty, zz + stz, xo + stx, yo + sty, zo + stz, RGB(red, gre, blu))
                            Call draw3dline(xx + stx, yy + sty, zz + stz, old(b * 200 - 1, 0) + stx, old(b * 200 - 1, 0) + sty, old(b * 200 - 1, 0) + stz, RGB(red, gre, blu))
                        End If
                        old(b * 200, 0) = xx
                        old(b * 200, 1) = yy
                        old(b * 200, 2) = zz
                        xo = xx
                        yo = xx
                        zo = xx
                    Next b
                    ProgressBar1.Value = a
                Next a
            Next c
        End If
        Call showing()
	End Sub
	
    Public Sub draw3dline(ByRef bx As Double, ByRef by As Double, ByRef bz As Double, ByRef ex As Double, ByRef ey As Double, ByRef ez As Double, ByRef co As Double)
        Dim dy As Double
        Dim dx As Double
        Dim n As Double
        Dim leng As Double
        Dim fy As Double
        Dim fx As Double
        Dim sy As Double
        Dim sx As Double
        Dim dist As Double
        Dim zdis As Double
        Dim ydis As Double
        Dim xdis As Double
        Dim ez2 As Double
        Dim ey2 As Double
        Dim ex2 As Double
        Dim bz2 As Double
        Dim by2 As Double
        Dim bx2 As Double
        bx2 = thrd(0, 0) * bx + thrd(1, 0) * by + thrd(2, 0) * bz
        by2 = thrd(0, 1) * bx + thrd(1, 1) * by + thrd(2, 1) * bz
        bz2 = thrd(0, 2) * bx + thrd(1, 2) * by + thrd(2, 2) * bz
        ex2 = thrd(0, 0) * ex + thrd(1, 0) * ey + thrd(2, 0) * ez
        ey2 = thrd(0, 1) * ex + thrd(1, 1) * ey + thrd(2, 1) * ez
        ez2 = thrd(0, 2) * ex + thrd(1, 2) * ey + thrd(2, 2) * ez
        xdis = xsee - (bx2 + ex2) / 2
        ydis = ysee - (by2 + ey2) / 2
        zdis = zsee - (bz2 + ez2) / 2
        dist = System.Math.Sqrt(xdis ^ 2 + ydis ^ 2 + zdis ^ 2)
        sx = two(0, 0) * bx + two(1, 0) * by + two(2, 0) * by
        sy = two(0, 1) * bx + two(1, 1) * by + two(2, 1) * by
        fx = two(0, 0) * ex + two(1, 0) * ey + two(2, 0) * ey
        fy = two(0, 1) * ex + two(1, 1) * ey + two(2, 1) * ey
        leng = System.Math.Abs(sx - sy)
        If leng < System.Math.Abs(sy - fy) Then leng = System.Math.Abs(sy - fy)
        If leng = 0 Then leng = 1
        For n = 0 To leng
            dx = (sx - fx) / leng * n + fx + 250
            dy = (sy - fy) / leng * n + fy + 250
            If dx >= 0 And dx <= 500 And dy >= 0 And dy <= 500 Then
                If win(dx, dy, 1) >= dist Then
                    win(500, 500, 1) = dist
                    If win(dx, dy, 0) <> co Then
                        win(dx, dy, 0) = co
                        win(dx, dy, 2) = 1
                        pointcount = pointcount + 1
                        pc(0) = (pc(0) + co) / 2
                    End If
                End If
            End If
        Next n
    End Sub

    Public Sub showing()
        Dim b As Double
        Dim a As Double
        For a = 0 To 500
            For b = 0 To 500
                If win(a, b, 2) = 1 Then
                    win(a, b, 2) = 0
                    'UPGRADE_ISSUE: Form メソッド Form1.PSet はアップグレードされませんでした。 詳細については次のリンクをクリックしてください : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
					Form1.DefInstance.PSet (a * 15 + 300, b * 15 + 300), win(a, b, 0)
                End If
            Next b
        Next a
        MsgBox("表示")
        MsgBox(pointcount & "," & pc(0))
    End Sub

    Public Sub 移動thrd(ByRef a As Double, ByRef b As Double, ByRef c As Double, ByRef d As Double, ByRef e As Double, ByRef f As Double, ByRef g As Double, ByRef h As Double, ByRef i As Double)
        thrd(0, 0) = a
        thrd(0, 1) = b
        thrd(0, 2) = c
        thrd(1, 0) = d
        thrd(1, 1) = e
        thrd(1, 2) = f
        thrd(2, 0) = g
        thrd(2, 1) = h
        thrd(2, 2) = i
    End Sub
End Class