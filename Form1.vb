Public Class Form1

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

    Dim push As Boolean = False

    Private thr As System.Threading.Thread
    Delegate Sub SetBingoNumCallBack(ByVal mNum As String)

    Sub run()

        Dim r As New System.Random()

        Do
            num = r.Next(0, range)

            If finish = True Then
                If cnt < 5 Then
                    tRate = 2
                    cnt += 1
                ElseIf (5 <= cnt) And (cnt < 10) Then
                    tRate = 3
                    cnt += 1
                ElseIf 10 <= cnt Then
                    'スレッドの強制終了
                    tRate = 1
                    cnt = 0

                    SetBingoNum((list(num) + 1).ToString)

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

            newLabel.Text = num.ToString
            newLabel.TextAlign = ContentAlignment.MiddleCenter
            newLabel.Anchor = AnchorStyles.Bottom Or AnchorStyles.Left Or AnchorStyles.Right Or AnchorStyles.Top

            'newLabel.Font = New Font("HGゴシックM", 20)
            TableLayoutPanel1.Controls.Add(newLabel, insertColumn, insertRow)
        End If
    End Sub

    '画像シャッフルスレッドを実行する
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles StartButton.Click
        'サブルーチンをスレッドとして作成
        thr = New System.Threading.Thread(AddressOf Me.run)
        'スレッドの開始
        thr.Start()
    End Sub

    '画像シャッフルスレッドを停止し、選ばれた番号をリストに表示、候補から削除する
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        finish = True

    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim path As String = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) & "\img\"

        For i = 0 To 74
            img(i) = System.Drawing.Image.FromFile(path & (i + 1).ToString & ".png")
            list(i) = i
        Next

        toppage = System.Drawing.Image.FromFile(path & "top.png")

        PictureBox1.Image = toppage
    End Sub

End Class
