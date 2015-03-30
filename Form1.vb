'ウィンドウサイズ:1024, 768
'デザインのツールはPictureBox1, Shutdown(Button1), Label1, TableLayoutPanel1の4つ
'TabelLayoutPanel1はフォントをHGｺﾞｼｯｸM, 24pt, style=Bold, アンカーを全方向
'PictureBox1はSizeModeをZoom
'Shutdownはテキストをソフト終了, 位置はそれなりのところに
'Label1については、場所は適当でも初期化の時に固定される

Public Class Form1
    
    Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" ( _
        ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Integer, ByVal hwndCallback As Integer) As Integer

    Declare Function timeGetTime Lib "winmm.dll" Alias "timeGetTime" () As Long

    Dim img(75) As System.Drawing.Image
    Dim list(75) As Integer
    Dim toppage As System.Drawing.Image

    Dim range As Integer = 74

    Dim tRate As Double = 1.0

    Dim num As Integer

    Dim insertRow As Integer = 0
    Dim insertColumn As Integer = 0

    Dim finish As Boolean = False

    Dim pushS As Boolean = False
    Dim pushF As Boolean = False

    Dim cmd(3) As String

    Private thr As System.Threading.Thread

    Private Sub Wait(ByVal waittime As Long)
        Dim starttime As Long

        starttime = timeGetTime()

        Do While timeGetTime() - starttime < waittime
            My.Application.DoEvents()
        Loop
    End Sub

    'シャッフル時にストップするまで画像を切り替えるスレッド
    Sub run()

        Dim r As New System.Random()

        Do
            '乱数生成
            num = r.Next(0, range)

            '画像変更
            PictureBox1.Image = img(num)

            'ストップしたらスレッドを終了
            If finish = True Then
                thr.Abort()
            End If

            '待ち時間は75msの倍数
            System.Threading.Thread.Sleep(75 * tRate)
        Loop

    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim imgpath As String = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) & "\img\"
        Dim mp3path As String = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) & "\se\"

        '画像をすべて読み込む
        For i = 0 To 74
            img(i) = System.Drawing.Image.FromFile(imgpath & (i + 1).ToString & ".png")
            list(i) = i
        Next

        'シャッフル前の画像を読み込み、表示する
        toppage = System.Drawing.Image.FromFile(imgpath & "top.png")
        PictureBox1.Image = toppage

        'ラベルを変更する
        Label1.Location = New Point(32, 622)
        Label1.Text = "スタート"

        '効果音3種のopenコマンドを生成する
        cmd(0) = "open """ + mp3path + "drumroll.mp3" + """ alias " + "drumroll"
        cmd(1) = "open """ + mp3path + "select.mp3" + """ alias " + "select"
        cmd(2) = "open """ + mp3path + "fanfare.mp3" + """ alias " + "fanfare"

        '効果音を開く
        mciSendString(cmd(0), "", 0, 0)
        mciSendString(cmd(1), "", 0, 0)
        mciSendString(cmd(2), "", 0, 0)

    End Sub

    Private Sub Form1_KeyDown(ByVal sender As Object, _
        ByVal e As System.Windows.Forms.KeyEventArgs) _
        Handles MyBase.KeyDown

        '動作キーは「S」
        If e.KeyCode = Keys.S Then
            If pushS = False Then   'スタート
                'シャッフル用スレッドを開始
                thr = New System.Threading.Thread(AddressOf Me.run)
                thr.Start()

                'ドラムロールの開始
                mciSendString("seek " + "drumroll" + " to start", "", 0, 0)
                mciSendString("play " + "drumroll", "", 0, 0)

                '1度押されたフラグを立てる
                pushS = True

                'ラベルを変更
                Label1.Location = New Point(7, 622)
                Label1.Text = "シャッフル中"
            Else                    'ストップ
                Dim r As New System.Random()    '乱数
                Dim newLabel As New Label()     '出た数字格納のためのラベル

                'ラベルを変更
                Label1.Location = New Point(32, 622)
                Label1.Text = "ストップ"

                'スレッドを停止し、押されたフラグをクリアする
                finish = True
                pushS = False

                '1度目のルーレット
                For i As Integer = 0 To 4
                    tRate = 2
                    mciSendString("seek " + "select" + " to start", "", 0, 0)
                    mciSendString("Play " + "select", "", 0, 0)
                    Wait(75 * tRate)
                    num = r.Next(0, range)
                    PictureBox1.Image = img(num)
                Next

                '2度目のルーレット
                For i As Integer = 0 To 4
                    tRate = 3
                    mciSendString("seek " + "select" + " to start", "", 0, 0)
                    mciSendString("Play " + "select", "", 0, 0)
                    Wait(75 * tRate)
                    num = r.Next(0, range)
                    PictureBox1.Image = img(num)
                Next

                '最終結果の生成
                num = r.Next(0, range)
                PictureBox1.Image = img(num)
                tRate = 1

                '他の音声を停止し、ファンファーレを鳴らす
                mciSendString("stop " + "drumroll", "", 0, 0)
                mciSendString("stop " + "select", "", 0, 0)
                mciSendString("seek " + "fanfare" + " to start", "", 0, 0)
                mciSendString("play " + "fanfare", "", 0, 0)

                '出た数字を表示する
                newLabel.Text = num.ToString
                newLabel.TextAlign = ContentAlignment.MiddleCenter
                newLabel.Anchor = AnchorStyles.Bottom Or AnchorStyles.Left Or AnchorStyles.Right Or AnchorStyles.Top
                TableLayoutPanel1.Controls.Add(newLabel, insertColumn, insertRow)
                '表示する場所の決定
                If insertColumn = 14 Then
                    insertColumn = 0
                    insertRow += 1
                Else
                    insertColumn += 1
                End If

                For j As Integer = num To range - 1
                    img(j) = img(j + 1)
                    list(j) = list(j + 1)
                Next

                '選ばれた数字を選択候補から外す
                range -= 1

                'スレッドを再度立ち上げ可能にする
                finish = False
            End If

        ElseIf e.KeyCode = Keys.F Then
            If pushF = False Then
                Me.WindowState = FormWindowState.Normal
                Me.FormBorderStyle = FormBorderStyle.None
                Me.WindowState = FormWindowState.Maximized
                pushF = True
            Else
                Me.FormBorderStyle = FormBorderStyle.Sizable
                Me.WindowState = FormWindowState.Normal
                pushF = False
            End If
        End If

    End Sub

    Private Sub Shutdown_Click(sender As Object, e As EventArgs) Handles Shutdown.Click
        'スレッドを開始している場合は、終了させる
        If pushS = True Then
            thr.Abort()
        End If
        
        Me.Close()
    End Sub
End Class
