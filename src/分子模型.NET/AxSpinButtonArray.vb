'UPGRADE_WARNING: ActiveX コントロール配列を含むフォームを表示するには、プロジェクト全体をコンパイルする必要があります。

Imports System.ComponentModel

<ProvideProperty("Index",GetType(AxMSForms.AxSpinButton))> Public Class AxSpinButtonArray
	Inherits Microsoft.VisualBasic.Compatibility.VB6.BaseOcxArray
	Implements IExtenderProvider

	Public Sub New()
		MyBase.New()
	End Sub

	Public Sub New(ByVal Container As IContainer)
		MyBase.New(Container)
	End Sub

	Public Shadows Event [BeforeDragOver] (ByVal sender As System.Object, ByVal e As AxMSForms.SpinbuttonEvents_BeforeDragOverEvent)
	Public Shadows Event [BeforeDropOrPaste] (ByVal sender As System.Object, ByVal e As AxMSForms.SpinbuttonEvents_BeforeDropOrPasteEvent)
	Public Shadows Event [Change] (ByVal sender As System.Object, ByVal e As System.EventArgs)
	Public Shadows Event [Error] (ByVal sender As System.Object, ByVal e As AxMSForms.SpinbuttonEvents_ErrorEvent)
	Public Shadows Event [KeyDownEvent] (ByVal sender As System.Object, ByVal e As AxMSForms.SpinbuttonEvents_KeyDownEvent)
	Public Shadows Event [KeyPressEvent] (ByVal sender As System.Object, ByVal e As AxMSForms.SpinbuttonEvents_KeyPressEvent)
	Public Shadows Event [KeyUpEvent] (ByVal sender As System.Object, ByVal e As AxMSForms.SpinbuttonEvents_KeyUpEvent)
	Public Shadows Event [SpinUp] (ByVal sender As System.Object, ByVal e As System.EventArgs)
	Public Shadows Event [SpinDown] (ByVal sender As System.Object, ByVal e As System.EventArgs)

	<System.ComponentModel.EditorBrowsableAttribute(System.ComponentModel.EditorBrowsableState.Never)> Public Function CanExtend(ByVal target As Object) As Boolean Implements IExtenderProvider.CanExtend
		If TypeOf target Is AxMSForms.AxSpinButton Then
			Return BaseCanExtend(target)
		End If
	End Function

	Public Function GetIndex(ByVal o As AxMSForms.AxSpinButton) As Short
		Return BaseGetIndex(o)
	End Function

	<System.ComponentModel.EditorBrowsableAttribute(System.ComponentModel.EditorBrowsableState.Never)> Public Sub SetIndex(ByVal o As AxMSForms.AxSpinButton, ByVal Index As Short)
		BaseSetIndex(o, Index)
	End Sub

	<System.ComponentModel.EditorBrowsableAttribute(System.ComponentModel.EditorBrowsableState.Never)> Public Function ShouldSerializeIndex(ByVal o As AxMSForms.AxSpinButton) As Boolean
		Return BaseShouldSerializeIndex(o)
	End Function

	<System.ComponentModel.EditorBrowsableAttribute(System.ComponentModel.EditorBrowsableState.Never)> Public Sub ResetIndex(ByVal o As AxMSForms.AxSpinButton)
		BaseResetIndex(o)
	End Sub

	Default Public ReadOnly Property Item(ByVal Index As Short) As AxMSForms.AxSpinButton
		Get
			Item = CType(BaseGetItem(Index), AxMSForms.AxSpinButton)
		End Get
	End Property

	Protected Overrides Function GetControlInstanceType() As System.Type
		Return GetType(AxMSForms.AxSpinButton)
	End Function

	Protected Overrides Sub HookUpControlEvents(ByVal o As Object)
		Dim ctl As AxMSForms.AxSpinButton = CType(o, AxMSForms.AxSpinButton)
		MyBase.HookUpControlEvents(o)
		If Not BeforeDragOverEvent Is Nothing Then
			AddHandler ctl.BeforeDragOver, New AxMSForms.SpinbuttonEvents_BeforeDragOverEventHandler(AddressOf HandleBeforeDragOver)
		End If
		If Not BeforeDropOrPasteEvent Is Nothing Then
			AddHandler ctl.BeforeDropOrPaste, New AxMSForms.SpinbuttonEvents_BeforeDropOrPasteEventHandler(AddressOf HandleBeforeDropOrPaste)
		End If
		If Not ChangeEvent Is Nothing Then
			AddHandler ctl.Change, New System.EventHandler(AddressOf HandleChange)
		End If
		If Not ErrorEvent Is Nothing Then
			AddHandler ctl.Error, New AxMSForms.SpinbuttonEvents_ErrorEventHandler(AddressOf HandleError)
		End If
		If Not KeyDownEventEvent Is Nothing Then
			AddHandler ctl.KeyDownEvent, New AxMSForms.SpinbuttonEvents_KeyDownEventHandler(AddressOf HandleKeyDownEvent)
		End If
		If Not KeyPressEventEvent Is Nothing Then
			AddHandler ctl.KeyPressEvent, New AxMSForms.SpinbuttonEvents_KeyPressEventHandler(AddressOf HandleKeyPressEvent)
		End If
		If Not KeyUpEventEvent Is Nothing Then
			AddHandler ctl.KeyUpEvent, New AxMSForms.SpinbuttonEvents_KeyUpEventHandler(AddressOf HandleKeyUpEvent)
		End If
		If Not SpinUpEvent Is Nothing Then
			AddHandler ctl.SpinUp, New System.EventHandler(AddressOf HandleSpinUp)
		End If
		If Not SpinDownEvent Is Nothing Then
			AddHandler ctl.SpinDown, New System.EventHandler(AddressOf HandleSpinDown)
		End If
	End Sub

	Private Sub HandleBeforeDragOver (ByVal sender As System.Object, ByVal e As AxMSForms.SpinbuttonEvents_BeforeDragOverEvent) 
		RaiseEvent [BeforeDragOver] (sender, e)
	End Sub

	Private Sub HandleBeforeDropOrPaste (ByVal sender As System.Object, ByVal e As AxMSForms.SpinbuttonEvents_BeforeDropOrPasteEvent) 
		RaiseEvent [BeforeDropOrPaste] (sender, e)
	End Sub

	Private Sub HandleChange (ByVal sender As System.Object, ByVal e As System.EventArgs) 
		RaiseEvent [Change] (sender, e)
	End Sub

	Private Sub HandleError (ByVal sender As System.Object, ByVal e As AxMSForms.SpinbuttonEvents_ErrorEvent) 
		RaiseEvent [Error] (sender, e)
	End Sub

	Private Sub HandleKeyDownEvent (ByVal sender As System.Object, ByVal e As AxMSForms.SpinbuttonEvents_KeyDownEvent) 
		RaiseEvent [KeyDownEvent] (sender, e)
	End Sub

	Private Sub HandleKeyPressEvent (ByVal sender As System.Object, ByVal e As AxMSForms.SpinbuttonEvents_KeyPressEvent) 
		RaiseEvent [KeyPressEvent] (sender, e)
	End Sub

	Private Sub HandleKeyUpEvent (ByVal sender As System.Object, ByVal e As AxMSForms.SpinbuttonEvents_KeyUpEvent) 
		RaiseEvent [KeyUpEvent] (sender, e)
	End Sub

	Private Sub HandleSpinUp (ByVal sender As System.Object, ByVal e As System.EventArgs) 
		RaiseEvent [SpinUp] (sender, e)
	End Sub

	Private Sub HandleSpinDown (ByVal sender As System.Object, ByVal e As System.EventArgs) 
		RaiseEvent [SpinDown] (sender, e)
	End Sub

End Class

