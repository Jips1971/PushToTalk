Public Class PortSelect

    Private Sub PortSelect_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ComboBoxPorts.Items.Clear()
        ComboBoxPorts.Items.AddRange(IO.Ports.SerialPort.GetPortNames())

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If ComboBoxPorts.SelectedItem Is Nothing Then
            MessageBox.Show("Please select a COM port from the list.")
            Return
        End If

        SelectedPortValue = ComboBoxPorts.SelectedItem.ToString()
        Me.DialogResult = DialogResult.OK


    End Sub
End Class