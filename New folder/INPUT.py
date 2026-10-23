import subprocess
import os
import xlwings as xw
import re
import numpy as np

def get_fwd_data(input_file):
    extracted_data = []
    
    # Open the text export file with utf-8 encoding
    with open(input_file, 'r', encoding='utf-16', errors='ignore') as f:
        for line in f:
            line = line.strip()
            if not line:
                continue
            
            # Skip metadata and header rows
            if any(line.startswith(prefix) for prefix in [
                'Date Created', 'Project Name', 'Project Location', 'Performer Name', 
                'Client Name', 'Location', 'Disttrict', 'Test Start', 'Test End', 
                'Deflection Sensors', 'Date\tTime']):
                continue
            
            # Split tab-separated values
            parts = line.split('\t')
            
            # Ensure we have enough columns for data rows
            if len(parts) >= 22:
                try:
                    # Convert chainage format like '0+100' to meters (e.g., 100.0)
                    ch_str = parts[4]
                    if '+' in ch_str:
                        ch_parts = ch_str.split('+')
                        chainage_val = int(ch_parts[0]) * 1000 + float(ch_parts[1])
                    else:
                        chainage_val = float(ch_str)

                    row = [
                        chainage_val,      # 0: Chainage in meters
                        int(parts[5]),     # 1: Drop Number
                        float(parts[6]),   # 2: Load (kN)
                        float(parts[7]),   # 3: D1 (Center)
                        float(parts[8]),   # 4: D2
                        float(parts[9]),   # 5: D3
                        float(parts[10]),  # 6: D4
                        float(parts[11]),  # 7: D5
                        float(parts[12]),  # 8: D6
                        float(parts[13]),  # 9: D7
                        float(parts[14]),  # 10: D8
                        float(parts[15]),  # 11: D9
                        float(parts[16]),  # 12: Air Temp (°C)
                        float(parts[17]),  # 13: Surface Temp (°C)
                        parts[18],         # 14: Comment
                        float(parts[19]),  # 15: Bituminous Thickness (mm)
                        float(parts[20]),  # 16: Granular Base Thickness (mm)
                        float(parts[21]),  # 17: Subgrade Thickness (mm)
                    ]
                    extracted_data.append(row)
                except ValueError:
                    # Skip rows where conversion fails
                    continue

    return extracted_data

def merger(s1,s2,sheet):
    app.screen_updating = False      # ← FIX #1: Disable screen refresh
    app.display_alerts = False       # ← FIX #2: Suppress dialogs
    app.enable_events = False
    sheet.range(f'{s1}:{s2}').api.UnMerge()        # ← FIX #3: Disable events
    sheet.range(f'{s1}:{s2}').clear_contents()
    try:
        sheet.range(f'{s1}:{s2}').api.Merge()
    finally:
        app.screen_updating = True
        app.display_alerts = True
        app.enable_events = True

with xw.App(visible=False) as app:
    input_file="NHAI remuli lha3000 200 500.txt"
    file_path = 'Modulus_Calculation.xlsx'
    if os.path.exists(file_path):
        wb2 = app.books.open(file_path)
    else:
        wb2 = app.books.add()  # Create new workbook
        wb2.save(file_path)
    #wb2=app.books.open('Modulus_Calculation.xlsx')
    try:
        ws2 = wb2.sheets['L1']
    except:
        ws2 = wb2.sheets.add('L2', before=wb2.sheets[0]) 
    result=get_fwd_data(input_file)
    #print(result)
    # Putting the Header Columns
    merger('A1','A2',ws2)
    ws2.range('A1').value="Chainage"
    merger('B1','B2',ws2)
    ws2.range('B1').value="Load(N)"
    merger('C1','C2',ws2)
    ws2.range('C1').value="Stress(MPa)"
    merger('D1','L1',ws2)
    ws2.range('D1').value="Measured Deflections at various sensors"
    ws2.range('D2').value="D0 (mm)\n0 mm"
    ws2.range('E2').value="D1 (mm)\n200 mm"
    ws2.range('F2').value="D2 (mm)\n300 mm"
    ws2.range('G2').value="D3 (mm)\n450 mm"
    ws2.range('H2').value="D4 (mm)\n600 mm"
    ws2.range('I2').value="D5 (mm)\n900 mm"
    ws2.range('J2').value="D6 (mm)\n1200 mm"
    ws2.range('K2').value="D7 (mm)\n1500 mm"
    ws2.range('L2').value="D8 (mm)\n1800 mm"
    merger('M1','N1',ws2)
    ws2.range('M1').value="Thickness of each layers"
    ws2.range('M2').value="Bituminous Layer Thickness"
    ws2.range('N2').value="Base Layer Thickness"
    merger('O1','T1',ws2)
    ws2.range('O1').value="Range of E values"
    ws2.range('O2').value="Bit. E lower Range"
    ws2.range('P2').value="Bit. E upper Range"
    ws2.range('Q2').value="Base E lower Range"
    ws2.range('R2').value="Base E upper Range"
    ws2.range('S2').value="Sub-Grade E lower Range"
    ws2.range('T2').value="Sub-Grade E upper Range"
    merger('U1','W1',ws2)
    ws2.range('U1').value="Predicted E values"
    ws2.range('U2').value="Bituminous (E1)"
    ws2.range('V2').value="Base (E2)"
    ws2.range('W2').value="Sub-Grade (E3)"
    merger('X1','X2',ws2)
    ws2.range('X1').value="RMSE Using IITPAVE(%)"
    merger('Y1','Y2',ws2)
    ws2.range('Y1').value="Asphalt Temperature"
    merger('Z1','Z2',ws2)
    ws2.range('Z1').value="Lambda"
    merger('AA1','AA2',ws2)
    ws2.range('AA1').value="Season"
    merger('AB1','AD1',ws2)
    ws2.range('AB1').value="Field Modulus After Temp & Seasonal Corrections"
    ws2.range('AB2').value="Bituminous (E1)"
    ws2.range('AC2').value="Base (E2)"
    ws2.range('AD2').value="Sub-Grade (E3)"
    merger('AE1','AE2',ws2)
    ws2.range('AE1').value="Sl No."
    merger('AF1','AG1',ws2)
    ws2.range('AF1').value="Chainage"
    ws2.range('AF2').value="From"
    ws2.range('AG2').value="To"
    merger('AH1','AH2',ws2)
    ws2.range('AH1').value="Length (Km)"
    merger('AI1','AK1',ws2)
    ws2.range('AI1').value="15th Percentile E Value "
    ws2.range('AI2').value="Bituminous Layer"
    ws2.range('AJ2').value="Base Layer"
    ws2.range('AK2').value="Sub-Grade Layer"
    merger('AL1','AL2',ws2)
    ws2.range('AL1').value="Fatigue Strain"
    merger('AM1','AM2',ws2)
    ws2.range('AM1').value="Rutting Strain"
    merger('AN1','AN2',ws2)
    ws2.range('AN1').value="Fatigue Residual Life(MSA)"
    merger('AO1','AO2',ws2)
    ws2.range('AO1').value="Rutting Residual Life(MSA)"
    merger('AP1','AP2',ws2)
    ws2.range('AP1').value="Remaining life(MSA) as per 90% Reliability"
    merger('AQ1','AQ2',ws2)
    ws2.range('AQ1').value="Overlay Required (Yes/No)"
    merger('AR1','AR2',ws2)
    ws2.range('AR1').value="Overlay Thickness Required (mm)"
    merger('AS1','AS2',ws2)
    ws2.range('AS1').value="E moduli of Overlay"
    merger('AT1','AT2',ws2)
    ws2.range('AT1').value="Poisson's Ratio of Overlay"
    merger('AU1','AU2',ws2)
    ws2.range('AU1').value="Fatigue Strain Post Overlay"
    merger('AV1','AV2',ws2)
    ws2.range('AV1').value="Rutting Strain Post Overlay"
    merger('AW1','AW2',ws2)
    ws2.range('AW1').value="Fatigue Residual Life(MSA) Post Overlay"
    merger('AX1','AX2',ws2)
    ws2.range('AX1').value="Rutting Residual Life(MSA) Post Overlay"
    merger('AY1','AY2',ws2)
    ws2.range('AY1').value="Remaining life(MSA) as per 90% Reliability Post Overlay"
    ws2.range('A1:AY2').api.WrapText = True
    ws2.range('A:AY').api.HorizontalAlignment = -4108
    ws2.range('A:AY').api.VerticalAlignment = -4108
    ws2.range('A1:AY2').api.Font.Bold = True
    ws2.range('A1:AY2').color = (192, 192, 192)
    ws2.range('A:AY').api.Borders.LineStyle = 1
    ws2.range('U:U').number_format='0.00'
    ws2.range('V:V').number_format='0.00'
    ws2.range('W:W').number_format='0.00'
    ws2.range('X:X').number_format='0.00'
    ws2.range('Z:Z').number_format='0.0000'
    ws2.range('AB:AB').number_format='0.00'
    ws2.range('AC:AC').number_format='0.00'
    ws2.range('AD:AD').number_format='0.00'
    ws2.range('AI:AI').number_format='0.00'
    ws2.range('AJ:AJ').number_format='0.00'
    ws2.range('AK:AK').number_format='0.00'
    ws2.range('AN:AN').number_format='0.00'
    ws2.range('AO:AO').number_format='0.00'
    ws2.range('AP:AP').number_format='0.00'
    ws2.range('AW:AW').number_format='0.00'
    ws2.range('AX:AX').number_format='0.00'
    ws2.range('AY:AY').number_format='0.00'
    ii=3
    j=0
    CBR=8
    while j<len(result):
        number=1
        d=[0,0,0,0,0,0,0,0,0]
        d[0]=(result[j][3]/result[j][2])*40
        d[1]=(result[j][4]/result[j][2])*40
        d[2]=(result[j][5]/result[j][2])*40
        d[3]=(result[j][6]/result[j][2])*40
        d[4]=(result[j][7]/result[j][2])*40
        d[5]=(result[j][8]/result[j][2])*40
        d[6]=(result[j][9]/result[j][2])*40
        d[7]=(result[j][10]/result[j][2])*40
        d[8]=(result[j][11]/result[j][2])*40
        for i in range(1,3):
            if result[j][0]==result[j+i][0]:
                d[0]+=(result[j+i][3]/result[j+i][2])*40
                d[1]+=(result[j+i][4]/result[j+i][2])*40
                d[2]+=(result[j+i][5]/result[j+i][2])*40
                d[3]+=(result[j+i][6]/result[j+i][2])*40
                d[4]+=(result[j+i][7]/result[j+i][2])*40
                d[5]+=(result[j+i][8]/result[j+i][2])*40
                d[6]+=(result[j+i][9]/result[j+i][2])*40
                d[7]+=(result[j+i][10]/result[j+i][2])*40
                d[8]+=(result[j+i][11]/result[j+i][2])*40
                number+=1
        thickness=[115,450]
        d=np.array(d)
        d=d/(number*1000)
        ranges=[400,3000,100,500,20,200]
        ws2.range(f'A{ii}').value=result[j][0]
        ws2.range(f'B{ii}').value=40000
        ws2.range(f'C{ii}').value=0.56
        ws2.range(f'D{ii}').value=d[0]
        ws2.range(f'E{ii}').value=d[1]
        ws2.range(f'F{ii}').value=d[2]
        ws2.range(f'G{ii}').value=d[3]
        ws2.range(f'H{ii}').value=d[4]
        ws2.range(f'I{ii}').value=d[5]
        ws2.range(f'J{ii}').value=d[6]
        ws2.range(f'K{ii}').value=d[7]
        ws2.range(f'L{ii}').value=d[8]
    # Thickness of bituminous and base layer   
        ws2.range(f'M{ii}').value=thickness[0]
        ws2.range(f'N{ii}').value=thickness[1]
    
        ws2.range(f'O{ii}').value=ranges[0]
        ws2.range(f'P{ii}').value=ranges[1]
        ws2.range(f'Q{ii}').value=ranges[2]
        ws2.range(f'R{ii}').value=ranges[3]
        ws2.range(f'S{ii}').value=ranges[4]
        ws2.range(f'T{ii}').value=ranges[5]
        ws2.range(f'Y{ii}').value=result[j][13]
        ws2.range(f'AA{ii}').value="Monsoon"
        print(f"Point {ii-2} done")
        ii=ii+1
        j+=number
    wb2.save()
    wb2.close()