Imports System.Linq.Expressions
Imports System


Public Class Form1
    ' Variabili per memorizzare l'espressione corrente
    Dim espressione As String = ""

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Inizializza la TextBox principale
        txtDisplay.Multiline = False
        txtDisplay.ReadOnly = True
        txtDisplay.TextAlign = HorizontalAlignment.Right
        ' txtDisplay.ScrollBars = ScrollBars.Horizontal

        ' Inizializza la TextBox dei risultati
        txtRisultati.Multiline = False

        txtRisultati.TextAlign = HorizontalAlignment.Right
        ' txtRisultati.ScrollBars = ScrollBars.Horizontal
    End Sub

    ' Gestore per i pulsanti numerici
    Private Sub Numero_Click(sender As Object, e As EventArgs) Handles btn0.Click, btn1.Click, btn2.Click, btn3.Click, btn4.Click, btn5.Click, btn6.Click, btn7.Click, btn8.Click, btn9.Click
        Dim pulsante As Button = CType(sender, Button)
        espressione &= pulsante.Text
        AggiornaDisplay()
    End Sub

    ' Gestore per i pulsanti operativi (+, -, *, /)
    Private Sub Operazione_Click(sender As Object, e As EventArgs) Handles btnAdd.Click, btnSubtract.Click, btnMultiply.Click, btnDivide.Click
        Dim pulsante As Button = CType(sender, Button)
        If espressione <> "" Then
            espressione &= pulsante.Text
            AggiornaDisplay()
        End If
    End Sub

    ' Gestore per il pulsante parentesi aperta
    Private Sub btnOpenParentheses_Click(sender As Object, e As EventArgs) Handles btnOpenParentheses.Click
        espressione &= "("
        AggiornaDisplay()
    End Sub

    ' Gestore per il pulsante parentesi chiusa
    Private Sub btnCloseParentheses_Click(sender As Object, e As EventArgs) Handles btnCloseParentheses.Click
        espressione &= ")"
        AggiornaDisplay()
    End Sub

    ' Gestore per il pulsante di percentuale
    Private Sub btnPercentage_Click(sender As Object, e As EventArgs) Handles btnPercentage.Click
        If espressione <> "" Then
            espressione &= " / 100"
            AggiornaDisplay()
        End If
    End Sub

    ' Gestore per il pulsante di uguale
    Private Sub btnEquals_Click(sender As Object, e As EventArgs) Handles btnEquals.Click
        If espressione <> "" Then

            ' Controllo del bilanciamento delle parentesi
            Dim aperte As Integer = espressione.Count(Function(c) c = "(")
            Dim chiuse As Integer = espressione.Count(Function(c) c = ")")

            If aperte < chiuse Then
                ' Ci sono più parentesi chiuse rispetto a quelle aperte, elimina quelle di troppo
                Dim diff As Integer = chiuse - aperte
                espressione = espressione.Substring(0, espressione.Length - diff)
            End If

            Dim risultato As Object = EvaluateExpression(espressione)
            If risultato IsNot Nothing Then
                ' Converte il risultato in stringa
                Dim risultatoString As String = CType(risultato, String)
                ' Verifica se il risultato inizia con il segno negativo
                If risultatoString.StartsWith("-") Then
                    ' Imposta il colore del testo in rosso
                    txtRisultati.ReadOnly = False
                    txtRisultati.ForeColor = Color.Red
                    txtRisultati.ReadOnly = True
                End If
                txtRisultati.AppendText(risultato.ToString())

                'espressione = risultato.ToString()
                ' txtDisplay.Text = espressione
                txtDisplay.SelectionStart = txtDisplay.Text.Length
                txtDisplay.ScrollToCaret()
                txtRisultati.SelectionStart = txtRisultati.Text.Length
                txtRisultati.ScrollToCaret()

            Else
                txtRisultati.AppendText(espressione & " = Errore di espressione: ")
                txtRisultati.SelectionStart = txtRisultati.Text.Length
                txtRisultati.ScrollToCaret()
            End If

        End If
        btnEquals.Enabled = False
    End Sub

    ' Gestore per il pulsante di cancellazione
    Private Sub btnClear_Click(sender As Object, e As EventArgs) Handles btnClear.Click
        espressione = ""
        AggiornaDisplay()
        txtRisultati.Text = ""
        btnEquals.Enabled = True
    End Sub

    ' Funzione per valutare l'espressione matematica
    Private Function EvaluateExpression(expression As String) As Object
        Dim dt As New DataTable()
        Dim risultato As Object
        Try
            ' Sostituisci il simbolo % con / 100
            expression = expression.Replace("%", "/100")
            risultato = dt.Compute(expression, "")
        Catch ex As Exception
            ' Gestione degli errori durante la valutazione dell'espressione
            risultato = ex.Message
        End Try
        Return risultato
    End Function

    ' Metodo per aggiornare il testo nella TextBox principale e farlo scorrere
    Private Sub AggiornaDisplay()
        txtDisplay.Text = espressione
        txtDisplay.SelectionStart = txtDisplay.Text.Length
        txtDisplay.ScrollToCaret()
    End Sub

    Private Sub ButtonCancella_Click(sender As Object, e As EventArgs) Handles ButtonCancella.Click
        ' Controlla se il TextBox ha del testo
        If txtDisplay.Text.Length > 0 Then
            ' Rimuovi l'ultimo carattere dal testo nel TextBox
            espressione = txtDisplay.Text.Substring(0, txtDisplay.Text.Length - 1)
            txtDisplay.Text = espressione
            txtDisplay.SelectionStart = txtDisplay.Text.Length
            txtDisplay.ScrollToCaret()
        End If
    End Sub

End Class
