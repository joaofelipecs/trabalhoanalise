' =============================================================================
' MÓDULO VBA — MacrosAssetName
' Como importar: Alt+F11 → clique direito no projeto → Importar arquivo
'               → escolha este arquivo MacrosAssetName.bas
'
' Pré-requisitos:
'   1. Python instalado com xlwings:  pip install xlwings pandas openpyxl
'   2. xlwings add-in instalado:      xlwings addin install   (no terminal)
'   3. Este arquivo .bas e asset_name.py na MESMA pasta que o .xlsm
' =============================================================================

Option Explicit

' -----------------------------------------------------------------------------
' SUB PRINCIPAL — vincule esta sub a um botão na planilha
' -----------------------------------------------------------------------------
Sub BotaoAtualizarAssetName()

    ' Mostra mensagem de progresso na barra de status do Excel
    Application.StatusBar = "Calculando Asset Name... aguarde."
    Application.ScreenUpdating = False

    On Error GoTo TratarErro

    ' Chama a função Python  →  arquivo: asset_name.py  |  função: atualizar_asset_name
    RunPython "import asset_name; asset_name.atualizar_asset_name()"

    ' Sucesso
    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox "Asset Name atualizado com sucesso!", vbInformation, "Concluído"
    Exit Sub

TratarErro:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox "Erro ao executar o Python:" & vbNewLine & Err.Description, _
           vbCritical, "Erro"

End Sub


' -----------------------------------------------------------------------------
' SUB AUXILIAR — use para testar se o Python/xlwings está funcionando
' -----------------------------------------------------------------------------
Sub TestarConexaoPython()
    RunPython "print('Python conectado com sucesso!')"
    MsgBox "Python respondeu! Verifique o console do xlwings.", vbInformation, "Teste OK"
End Sub
