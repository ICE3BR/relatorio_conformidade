# TESTE INICAL

teste texto



```python
pyinstaller --onefile  --icon "icon.ico" --noconsole --clean --name=RelatorioConformidade --add-data="MODELO RELATORIO.docx;." --add-data="Conformidade  - RPV.docx;." --hidden-import=pandas --hidden-import=openpyxl --hidden-import=docx --hidden-import=dateutil.relativedelta --hidden-import=tkinter v3.py
```
