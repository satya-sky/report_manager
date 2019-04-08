COLUMN_CONFIG = {
    'CLS': {
        'Column1': {
            'alignment': {'horizontal': 'left', 'vertical': 'center'},
            'border': {'top': 'thick', 'bottom': 'thick', 'left': 'thin', 'right': 'thin'},
            'pattern': {'color': '00B5E2FF', 'type': 'solid', 'fontSize': 13},
        }
    }
    'NIZ': {
        // TODO
    }
}

def getConfig(name):
    return COLUMN_CONFIG.get(name)

def formatSheet(name='CLS', ws):
    config = getConfig(name)
    rows = df.shape[0] + 1
    columns = df.shape[1] + 1
    for col, colConfig in config.iteritems():
        if colConfig.get('pattern'):
            pattern_color = openpyxl.styles.colors.Color(rgb=colConfig['pattern']['color'])
            pattern_fill = openpyxl.styles.fills.PatternFill(patternType=colConfig['pattern']['type'], fgColor=pattern_color)
            pattern_font = Font(size = colConfig['pattern']['fontSize'])
        if colConfig.get('border'):
            border_style = Border(left=Side(style=colConfig['border']['left']),
                                right=Side(style=colConfig['border']['right']),
                                top=Side(style=colConfig['border']['top']),
                                bottom=Side(style=colConfig['border']['bottom']))
        for i in range(2,rows):
            for j in range(i,rows):
                 if ws.cell(row = j, column = 1).value == ws.cell(row = j+1, column = 1).value:
                     j = j+1
                 else:
                     ws.merge_cells(start_row = i, start_column = 1, end_row = j, end_column = 1)
                     cell = ws.cell(row=i, column=1)
                     cell.border = border_style
                     cell.alignment = Alignment(horizontal=colConfig['alignment']['horizontal'], vertical=colConfig['alignment']['vertical'])
                     cell.fill = pattern_fill
                     cell.font = pattern_font
                     i = j
                     break


if __name__ == "__main__":
    formatSheet(CLS_CONFIG)
