Set WShell = CreateObject("WScript.Shell")
WShell.Run """C:\Users\Operador\AppData\Local\Programs\Python\Python312\Scripts\streamlit.exe"" run ""C:\Users\Operador\OneDrive\Desktop\claude\proyectos claude\informe_stock.py"" --server.port 8501 --server.address 0.0.0.0 --server.headless true --browser.gatherUsageStats false", 0, False
