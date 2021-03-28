VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PadFlowLauncher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'初期値用定数
Private Const INIT_WIN_TITLE = "Power Automate Desktop (プレビュー)"
Private Const INIT_BTN_NAME = "実行"
Private Const PAD_EXE_PATH = "C:\Program Files (x86)\Power Automate Desktop\PAD.Console.Host.exe"

Private Const KEY_RIGHT = "{Right}"
Private Const MAX_RETRY = 2

'プロパティ値
Private pWinTitle   As String
Private pBtnName    As String
Private pPADexePath As String

Private pMe         As PadFlowLauncher
Private pObjUIA     As UIAutomationClient.CUIAutomation
Private pObjRoot    As UIAutomationClient.IUIAutomationElement
Private pErr        As tmpErr

Private Type tmpErr
    Description As String
    Source      As String
    Number      As String
End Type

Private Sub Class_Initialize()
    
    On Error GoTo CATCH
    
    pErr.Number = 0
    pWinTitle = INIT_WIN_TITLE
    pBtnName = INIT_BTN_NAME
    pPADexePath = PAD_EXE_PATH
    Set pObjUIA = New UIAutomationClient.CUIAutomation
    Set pObjRoot = pObjUIA.GetRootElement()
    
    Err.Clear: GoTo FINAL
    
CATCH:
    With pErr
        .Number = Err.Number
        .Description = Err.Description
        .Source = Err.Source
    End With
    
FINAL:
    If pErr.Number <> 0 Then
        MsgBox pErr.Description, vbCritical, "Initialize Error"
        End
    End If
    
End Sub

Private Sub Class_Terminate()
    
    Set pObjRoot = Nothing
    Set pObjUIA = Nothing
    
End Sub

'Power Automation Desktopコンソールウィンドウのタイトル
Public Property Let windowTitle(ByVal strWindowTitle As String)
    pWinTitle = strWindowTitle
End Property

'Power Automation Desktopコンソールウィンドウの実行ボタン名
Public Property Let buttonName(ByVal strpBtnName As String)
    pBtnName = strpBtnName
End Property

'Power Automation Desktopコンソールのexeファイルパス
Public Property Let PADexePath(ByVal strPath As String)
    pPADexePath = strPath
End Property

'Elementを取得するメソッド
' objPrnt   :   検索の基準となる親Element
' strNameId :   取得したいElementのNamePropertyId
' lngIndex  :   一致したElementリストのn番目
' 戻り値    :   取得したElement,取得できなかった場合はNothingを返す
Private Function getElement(ByVal objPrnt As UIAutomationClient.IUIAutomationElement, _
                            ByVal strNameId As String, _
                            ByVal lngIndex As Long) As UIAutomationClient.IUIAutomationElement
    
    Dim objCon As UIAutomationClient.IUIAutomationCondition
    Dim objArr As UIAutomationClient.IUIAutomationElementArray
    Dim objRet As UIAutomationClient.IUIAutomationElement
    
    Set objRet = Nothing
    
    On Error GoTo CATCH
    
    Set objCon = pObjUIA.CreatePropertyCondition(UIA_NamePropertyId, strNameId)
    Set objArr = objPrnt.FindAll(TreeScope_Element Or TreeScope_Descendants, objCon)
    
    On Error Resume Next
        Set objRet = objArr.getElement(lngIndex)
    Err.Clear: GoTo FINAL
    
CATCH:
    With pErr
        .Number = Err.Number
        .Description = Err.Description
        .Source = Err.Source
    End With
    
FINAL:
    On Error Resume Next
        Set getElement = objRet
        
        Set objArr = Nothing
        Set objCon = Nothing
    On Error GoTo 0
    
    If pErr.Number <> 0 Then
        Err.Raise pErr.Number, pErr.Source, pErr.Description
    End If
    
End Function

'getElementメソッドをRetryしながら処理するメソッド
' objPrnt   :   検索の基準となる親Element
' strNameId :   取得したいElementのNamePropertyId
' lngIndex  :   一致したElementリストのn番目
' lngTimes  :   最大Retry回数
' 戻り値    :   取得したElement,取得できなかった場合はNothingを返す
Private Function getElementLapper(ByVal objPrnt As UIAutomationClient.IUIAutomationElement, _
                                  ByVal strNameId As String, _
                                  ByVal lngIndex As Long, _
                                  ByVal lngTimes As Long) As UIAutomationClient.IUIAutomationElement
    
    Dim objEle As UIAutomationClient.IUIAutomationElement
    Dim lngCnt As Long: lngCnt = 0
    
    Do
        lngCnt = lngCnt + 1
        
        On Error Resume Next
            Set objEle = Nothing
            Set objEle = getElement(objPrnt, strNameId, lngIndex)
        On Error GoTo 0
        
        If Not objEle Is Nothing Then
            Exit Do
        Else
            If lngCnt > (MAX_RETRY * lngTimes) Then
                Err.Raise 1, "getElementLapper", "Cannot get a UIAElement:" & strNameId
            End If
        End If
        
        Application.Wait Now() + TimeValue("00:00:01")
    Loop
    
    Set getElementLapper = objEle
    
End Function

'Power Automate Desktopのフローを実行するメソッド
' strFlowName   :   実行するフロー名
Public Sub runFlow(ByVal strFlowName As String)
    
    Dim objFrm As UIAutomationClient.IUIAutomationElement
    Dim objRow As UIAutomationClient.IUIAutomationElement
    Dim objBtn As UIAutomationClient.IUIAutomationElement
    Dim objIvk As UIAutomationClient.IUIAutomationInvokePattern
    Dim lngCnt As Long: lngCnt = 0
    
    On Error GoTo CATCH
    
    'ウィンドウが最小化や通知領域にある場合でもexeを実行すれば
    'ウィンドウ表示されるはずなのでexeを実行するだけにしておく
    With New IWshRuntimeLibrary.WshShell
        .Run """" & pPADexePath & """", 0, False
    End With
    
    'ウィンドウ取得
    Set objFrm = getElementLapper(pObjRoot, pWinTitle, 0, 10)
    
    '対象行取得
    Set objRow = getElementLapper(objFrm, strFlowName, 1, 2)
    objRow.SetFocus
    Application.Wait Now() + TimeValue("00:00:01")
    SendKeys KEY_RIGHT
    
    'ボタン取得＆クリック
    Set objBtn = getElementLapper(objRow, pBtnName, 0, 1)
    Set objIvk = objBtn.GetCurrentPattern(UIA_InvokePatternId)
    objIvk.Invoke
    
    Err.Clear: GoTo FINAL
    
CATCH:
    With pErr
        .Number = Err.Number
        .Description = Err.Description
        .Source = Err.Source
    End With
    
FINAL:
    On Error Resume Next
        Set objIvk = Nothing
        Set objBtn = Nothing
        Set objRow = Nothing
        Set objFrm = Nothing
    On Error GoTo 0
    
    If pErr.Number <> 0 Then
        MsgBox pErr.Description, vbCritical, "runFlow Error"
        End
    End If
    
End Sub
