'ウィンドウサイズ:1024, 768
'デザインのツールはPictureBox1, Shutdown(Button1), Label1, TableLayoutPanel1の4つ
'TabelLayoutPanel1はフォントをHGｺﾞｼｯｸM, 24pt, style=Bold, アンカーを全方向
'PictureBox1はSizeModeをZoom
'Shutdownはテキストをソフト終了, 位置はそれなりのところに
'Label1については、場所は適当でも初期化の時に固定される

Public Class Form1
    <System.Runtime.InteropServices.DllImport("winmm.dll", CharSet:=System.Runtime.InteropServices.CharSet.Auto)> _
    Private Shared Function mciSendString(ByVal command As String, ByVal buffer As System.Text.StringBuilder, _
        ByVal bufferSize As Integer, ByVal hwndCallback As IntPtr) As Integer
    End Function

    Dim img(75) As System.Drawing.Image
    Dim list(75) As Integer
    Dim toppage As System.Drawing.Image

    Dim range As Integer = 74
    Dim cnt As Integer = 0
    Dim first As Boolean = True

    Dim tRate As Double = 1.0

    Dim num As Integer

    Dim insertRow As Integer = 0
    Dim insertColumn As Integer = 0

    Dim finish As Boolean = False

    Dim pushS As Boolean = False
    Dim pushF As Boolean = False

    Dim cmd(3) As String

    Private thr As System.Threading.Thread
    Delegate Sub SetBingoNumCallBack(ByVal mNum As String)
    Delegate Sub SetBingoTextCallBack(ByVal mNum As String)

    Sub run()

        Dim r As New System.Random()

        Do
            num = r.Next(0, range)

            PictureBox1.Image = img(num)

            SetBingoText("")

            If finish = True Then
                If cnt < 5 Then
                    tRate = 2
                    cnt += 1
                    mciSendString("Play " + "select", Nothing, 0, IntPtr.Zero)
                ElseIf (5 <= cnt) And (cnt < 10) Then
                    tRate = 3
                    cnt += 1
                    mciSendString("Play " + "select", Nothing, 0, IntPtr.Zero)
                ElseIf 10 <= cnt Then
                    Dim showNum As Integer = list(num) + 1

                    tRate = 1
                    cnt = 0

                    mciSendString("Stop " + "drumroll", Nothing, 0, IntPtr.Zero)
                    mciSendString("Stop " + "select", Nothing, 0, IntPtr.Zero)
                    mciSendString("Play " + "fanfare", Nothing, 0, IntPtr.Zero)

                    SetBingoNum(showNum.ToString)

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

                    range -= 1

                    finish = False
                    thr.Abort()
                End If
            End If

            System.Threading.Thread.Sleep(75 * tRate)
        Loop

    End Sub

    'おそらくこの中で画像の更新を行う(リストに番号を表示するのは停止ボタンイベント内で行う)
    Private Sub SetBingoNum(ByVal mNum As String)
        If Me.Label1.InvokeRequired Then
            Dim d As New SetBingoNumCallBack(AddressOf SetBingoNum)
            Me.Invoke(d, New Object() {mNum})
        Else
            Dim newLabel As New Label()

            newLabel.Text = mNum
            newLabel.TextAlign = ContentAlignment.MiddleCenter
            newLabel.Anchor = AnchorStyles.Bottom Or AnchorStyles.Left Or AnchorStyles.Right Or AnchorStyles.Top
            TableLayoutPanel1.Controls.Add(newLabel, insertColumn, insertRow)
        End If
    End Sub

    Private Sub SetBingoText(ByVal mNum As String)
        If Me.Label1.InvokeRequired Then
            Dim d As New SetBingoTextCallBack(AddressOf SetBingoText)
            Me.Invoke(d, New Object() {mNum})
        Else
            If pushS = True Then
                Label1.Location = New Point(7, 622)
                Label1.Text = "シャッフル中"
            Else
                Label1.Location = New Point(32, 622)
                Label1.Text = "ストップ"
            End If
        End If
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim imgpath As String = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) & "\img\"
        Dim mp3path As String = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) & "\se\"

        For i = 0 To 74
            img(i) = System.Drawing.Image.FromFile(imgpath & (i + 1).ToString & ".png")
            list(i) = i
        Next

        toppage = System.Drawing.Image.FromFile(imgpath & "top.png")

        PictureBox1.Image = toppage

        Label1.Location = New Point(32, 622)
        Label1.Text = "スタート"

        'http://homepage1.nifty.com/rucio/main/kikou/kikou3_mciSendString.html
        cmd(0) = "open """ + mp3path + "drumroll.mp3" + """ alias " + "dorumroll"
        cmd(1) = "open """ + mp3path + "select.mp3" + """ alias " + "select"
        cmd(2) = "open """ + mp3path + "fanfare.mp3" + """ alias " + "fanfare"

        mciSendString(cmd(0), Nothing, 0, IntPtr.Zero)
        mciSendString(cmd(1), Nothing, 0, IntPtr.Zero)
        mciSendString(cmd(2), Nothing, 0, IntPtr.Zero)
    End Sub

    Private Sub Form1_KeyDown(ByVal sender As Object, _
        ByVal e As System.Windows.Forms.KeyEventArgs) _
        Handles MyBase.KeyDown

        '動作キーは「スペース」
        If e.KeyCode = Keys.S Then
            If pushS = False Then   'スタート
                'サブルーチンをスレッドとして作成
                thr = New System.Threading.Thread(AddressOf Me.run)
                'スレッドの開始
                thr.Start()
                mciSendString("Play " + "drumroll", Nothing, 0, IntPtr.Zero)
                pushS = True
            Else                    'ストップ
                finish = True
                pushS = False
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
        If pushS = True Then
            thr.Abort()
        End If
        mciSendString("Close " + "drumroll", Nothing, 0, IntPtr.Zero)
        mciSendString("Close " + "select", Nothing, 0, IntPtr.Zero)
        mciSendString("Close " + "fanfare", Nothing, 0, IntPtr.Zero)
        Me.Close()
    End Sub
End Class
