import subprocess

def pedir_archivo():
    script = '''
    Add-Type -AssemblyName System.Windows.Forms
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Filter = "Archivos Excel (*.xlsx;*.xls,;*.xlsm)|*.xlsx;*.xls;*.xlsm"
    $dialog.Title = "Selecciona un archivo Excel"
    if ($dialog.ShowDialog() -eq 'OK') { $dialog.FileName }
    '''
    
    resultado = subprocess.run(
        ["powershell", "-Command", script],
        capture_output=True,
        text=True,
        creationflags=subprocess.CREATE_NO_WINDOW
    )
    
    archivo = resultado.stdout.strip()
    if not archivo:
        return None
    
    if not archivo:
        print("No se seleccionó ningún archivo.")
        return None
    
    print(f"Archivo seleccionado: {archivo}")
    return archivo
