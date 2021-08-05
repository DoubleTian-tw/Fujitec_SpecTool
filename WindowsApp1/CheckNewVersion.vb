Public Class CheckNewVersion
    '字串常數
    Private Const DefaultInfoCode As String = "UTF-8" '儲存預設編碼為utf-8
    Private Const DefaultInfoFeature As String = "小工具更新" '儲存預設更新版本資訊的特徵
    Private Const DefaultInfoSeparator As String = "|" '儲存預設更新版本資訊的分隔線
    Private Const DefaultNotCheckReturn As String = "Not yet" '儲存還沒檢查更新時的回傳訊息
    Private Const DefaultCheckErrorReturn As String = "Error" '儲存檢查更新失敗時回傳的訊息

    '更新資訊組態
    Private sInfoURL As String '儲存更新版本資訊的網址
    Private sInfoCode As String = DefaultInfoCode '編碼
    Private sInfoFeature_main As String = DefaultInfoFeature '儲存版本更新特徵，預設為"小工具更新"
    Private sInfoFeature_update As String = DefaultInfoFeature '儲存版本更新特徵，預設為"小工具更新"


    '檢查結果
    Private shInfoConsequence As Short = -1 '0為沒有新版本;1為有新版本;2為檢查更新失敗;-1為尚未檢查
    Private shInfoConsequence_up As Short = -1 '0為沒有新版本;1為有新版本;2為檢查更新失敗;-1為尚未檢查
    Private sInfoNumber_main As String '儲存取得的版本號碼
    Private sInfoNumber_update As String '儲存取得的版本號碼

    Public Sub New(ByVal SetInfoURL As String) 'CheckNew建構子(多載1),引述為SetInfoURL,載入更新版本的網址
        sInfoURL = SetInfoURL '設定更新版本資訊的網址
    End Sub

    Public Sub New(ByVal SetInfoURL As String, ByVal SetInfoFeature As String) 'CheckNew建構子(多載2),載入網址與判斷特徵
        sInfoURL = SetInfoURL '設定更新版本資訊的網址
        sInfoFeature_main = SetInfoFeature ''設定更新版本資訊的特徵
        sInfoFeature_update = SetInfoFeature ''設定更新版本資訊的特徵_up
    End Sub

    Public WriteOnly Property SetInfoCode() As String '設定更新版本資訊的編碼
        Set(ByVal Code As String)
            sInfoCode = Code
        End Set
    End Property

    Public WriteOnly Property SetInfoURL() As String '設定網址
        Set(ByVal URL As String)
            sInfoURL = URL
        End Set
    End Property

    Public WriteOnly Property SetInfoFeature() As String '設定特徵
        Set(ByVal Feature As String)
            sInfoFeature_main = Feature
        End Set
    End Property

    Public ReadOnly Property GetMyVersion() As String '取得程式的版本號碼(格式:x.x.x)
        Get
            Return My.Application.Info.Version.Major & "." & My.Application.Info.Version.Minor & "." & My.Application.Info.Version.Build '取得版本號碼
        End Get
    End Property

    Public ReadOnly Property GetMyVersion_up() As String '取得程式的版本號碼(格式:x.x.x)
        Get
            Dim fileName As String = $"{ProgramAllName.fileName_updateProgram}.exe" '"Update_magicTool.exe"
            Dim filePath As String
            Dim fileVer_Maj, fileVer_Min, fileVer_Build, fileVersion As String

            Try
                filePath = System.IO.Path.GetFullPath(fileName)

                fileVer_Build = FileVersionInfo.GetVersionInfo(filePath).FileBuildPart
                fileVer_Maj = FileVersionInfo.GetVersionInfo(filePath).FileMajorPart
                fileVer_Min = FileVersionInfo.GetVersionInfo(filePath).FileMinorPart
                fileVersion = fileVer_Maj & "." & fileVer_Min & "." & fileVer_Build

                Return fileVersion
            Catch
            End Try
        End Get
    End Property

    Public ReadOnly Property GetCheckConsequence() As String '取得檢查更新的結果
        Get
            Return shInfoConsequence
        End Get
    End Property

    Public ReadOnly Property GetCheckConsequence_up() As String '取得檢查更新的結果
        Get
            Return shInfoConsequence_up
        End Get
    End Property

    Public ReadOnly Property GetCheckConsequenceNumber() As String '取得最新版本號碼
        Get
            If shInfoConsequence = -1 Then '如果還沒檢查
                Return DefaultNotCheckReturn
            ElseIf shInfoConsequence = 2 Then '如果檢查失敗
                Return DefaultCheckErrorReturn
            Else '如果檢查成功
                Return sInfoNumber_main
            End If
        End Get
    End Property

    Public ReadOnly Property GetCheckConsequenceNumber_up() As String '取得最新版本號碼
        Get
            If shInfoConsequence_up = -1 Then '如果還沒檢查
                Return DefaultNotCheckReturn
            ElseIf shInfoConsequence_up = 2 Then '如果檢查失敗
                Return DefaultCheckErrorReturn
            Else '如果檢查成功
                Return sInfoNumber_update
            End If
        End Get
    End Property

    Public Sub CheckNewVersion() '檢查更新
        On Error Resume Next
        Dim CheckNewClient As New Net.WebClient '宣告checkNewClient為webClinet物件
        Dim Encode As System.Text.Encoding = System.Text.Encoding.GetEncoding(sInfoCode) '設定編碼
        Dim Info As IO.StreamReader = New IO.StreamReader(CheckNewClient.OpenRead(sInfoURL), Encode) '取得更新資訊
        Dim sInfoData As String = Info.ReadLine '儲存更新第一行資訊

        Debug.Print("InfoData = """ & sInfoData & """") '除錯用，顯示出從網路上下載的字串

        If Strings.InStr(1, sInfoData, sInfoFeature_main) > 0 Then '若能找到識別關鍵字(InfoFeature)，則更新資訊取得成功
            Dim InfoArr() As String = Split(sInfoData, DefaultInfoSeparator) '分解字串為版本號碼(0)和識別關鍵字(1)
            sInfoNumber_main = InfoArr(0) '儲存最新版本號碼
            If InfoArr(0) <> GetMyVersion() Then '比對版本號碼，若不相同，代表有更新
                shInfoConsequence = 1
            Else
                shInfoConsequence = 0
            End If
        Else '若未能找到識別關鍵字(專案名稱)，則更新資訊取得失敗
            shInfoConsequence = 2
        End If
    End Sub

    Public Sub CheckNewVersion_Up() '檢查更新
        On Error Resume Next
        Dim CheckNewClient As New Net.WebClient '宣告checkNewClient為webClinet物件
        Dim Encode As System.Text.Encoding = System.Text.Encoding.GetEncoding(sInfoCode) '設定編碼
        Dim Info As IO.StreamReader = New IO.StreamReader(CheckNewClient.OpenRead(sInfoURL), Encode) '取得更新資訊
        Dim sInfoData As String = Info.ReadLine '儲存更新資訊
        Dim sInfoData2 As String = Info.ReadLine '儲存更新第二行資訊

        Debug.Print("InfoData = """ & sInfoData2 & """") '除錯用，顯示出從網路上下載的字串

        If Strings.InStr(1, sInfoData2, sInfoFeature_update) > 0 Then '若能找到識別關鍵字(InfoFeature)，則更新資訊取得成功
            Dim InfoArr() As String = Split(sInfoData2, DefaultInfoSeparator) '分解字串為版本號碼(0)和識別關鍵字(1)
            sInfoNumber_update = InfoArr(0) '儲存最新版本號碼
            If InfoArr(0) <> GetMyVersion_up() Then '比對版本號碼，若不相同，代表有更新
                shInfoConsequence_up = 1
            Else
                shInfoConsequence_up = 0
            End If
        Else '若未能找到識別關鍵字(專案名稱)，則更新資訊取得失敗
            shInfoConsequence_up = 2
        End If
    End Sub
End Class
