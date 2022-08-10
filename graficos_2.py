import openpyxl
from openpyxl.chart import BarChart, Reference
# Formatar para uma função

planilha = openpyxl.load_workbook('teste.xlsx')
# Definir uma rotina para colocar o gŕafico na aba certa
ws = planilha.worksheets[9]

# categoria = Reference(ws, min_col=1, min_row=2, max_row=9)
categoria = Reference(ws, min_row=1, min_col=2, max_col=6)
# valores = Reference(ws, min_col=2, max_col=6, min_row=1, max_row=9)
valores = Reference(ws, min_col=1, max_col=6, min_row=1, max_row=9)

graphic = BarChart()

graphic.add_data(valores, titles_from_data=True, from_rows=True)
graphic.set_categories(categoria)

ws.add_chart(graphic, 'B11')

print(valores)
planilha.save('teste.xlsx')
