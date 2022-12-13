import PySimpleGUI as sg # gui 구현모듈
from evtx import PyEvtxParser # .evtx 파일 분석 모듈
import json
import webbrowser # html, css 생성 모듈
import sys
import os
import win32com.shell.shell as shell

# 이벤트 로그 파일을 첨부하려면 권한이 필요하기 때문에 권한 상승 코드를 처음에 실행하여 권한 상승한 뒤 프로그램 시작
if sys.argv[-1] != 'asadmin':
    script = os.path.abspath(sys.argv[0])
    params = ' '.join([script] + sys.argv[1:] + ['asadmin'])
    shell.ShellExecuteEx(lpVerb='runas', lpFile=sys.executable, lpParameters=params)
    sys.exit(0)

# ---functions definition
def volumeSNParser(rec, vbrNo, sig):
    if sig == '4D53444F5335':  # FAT34
        s = rec['Event']['EventData'][vbrNo][134:142]
        sn = "".join(reversed([s[i:i + 2] for i in range(0, len(s), 2)]))
        return sn
    elif sig == '455846415420':  # ExFAT volume
        s = rec['Event']['EventData'][vbrNo][200:208]
        sn = "".join(reversed([s[i:i + 2] for i in range(0, len(s), 2)]))
        return sn
    elif sig == '4E5446532020':  # NTFS volume
        s = rec['Event']['EventData'][vbrNo][144:152]
        sn = "".join(reversed([s[i:i + 2] for i in range(0, len(s), 2)]))
        return sn
    else:
        s = 'unknown device'
        return s


def fullparse():
    filename = values['-IN-']   # 파일명 (evtx)
    FullParseRecordsDict = {}   # 모든 연결 매체의 로그 기록 목록
    AllPluggedInSerials = []    # 연결된 매체의 시리얼넘버
    EachPluggedDeviceDict = {}  # 각각의 연결 매체 목록
    IsPartitionDiagnosticEVTXFullParse = True # evtx에 대한 분석여부 확인
    FullParseHTMLWritten = True               # html파일 작성 성공여부 확인


    try:                        #
        parser = PyEvtxParser(filename)
        for record in parser.records_json():
            data = json.loads(record['data'])
            FullParseRecordsDict[(data['Event']['System']['EventRecordID'])] = data
            if data['Event']['System']['EventID'] != 1006:
                IsPartitionDiagnosticEVTXFullParse = False
        FullParseLogStartTime = FullParseRecordsDict[1]['Event']['System']['TimeCreated']['#attributes']['SystemTime'].replace('T', ' ').replace('Z', ' UTC')
        FullParseLogEndTime = \
        FullParseRecordsDict[len(FullParseRecordsDict)]['Event']['System']['TimeCreated']['#attributes']['SystemTime'].replace('T', ' ').replace('Z', ' UTC')

        if IsPartitionDiagnosticEVTXFullParse:
            for i in sorted(FullParseRecordsDict.keys()):
                if FullParseRecordsDict[i]['Event']['EventData']['SerialNumber'] not in AllPluggedInSerials:
                    AllPluggedInSerials.append(FullParseRecordsDict[i]['Event']['EventData']['SerialNumber'])
                    if FullParseRecordsDict[i]['Event']['EventData']['PartitionStyle'] == 1:
                        EachPluggedDeviceDict[FullParseRecordsDict[i]['Event']['EventData']['SerialNumber']] = [FullParseRecordsDict[i]['Event']['EventData']['Manufacturer'],
                            FullParseRecordsDict[i]['Event']['EventData']['Model'],
                            FullParseRecordsDict[i]['Event']['System']['TimeCreated']['#attributes']['SystemTime'].replace('T', ' ').replace('Z', ' UTC'),
                            FullParseRecordsDict[i]['Event']['System']['TimeCreated']['#attributes']['SystemTime'].replace('T', ' ').replace('Z', ' UTC'), ['GPT - NO VSN INFO']]
                    else:
                        SN0 = volumeSNParser(FullParseRecordsDict[i], 'Vbr0',
                                             FullParseRecordsDict[i]['Event']['EventData']['Vbr0'][6:18])
                        if FullParseRecordsDict[i]['Event']['EventData']['Vbr1'] != '':
                            SN1 = volumeSNParser(FullParseRecordsDict[i], 'Vbr1',
                                                 FullParseRecordsDict[i]['Event']['EventData']['Vbr1'][6:18])
                        else:
                            SN1 = '-'
                        if FullParseRecordsDict[i]['Event']['EventData']['Vbr2'] != '':
                            SN2 = volumeSNParser(FullParseRecordsDict[i], 'Vbr2',
                                                 FullParseRecordsDict[i]['Event']['EventData']['Vbr2'][6:18])
                        else:
                            SN2 = '-'
                        if FullParseRecordsDict[i]['Event']['EventData']['Vbr3'] != '':
                            SN3 = volumeSNParser(FullParseRecordsDict[i], 'Vbr3',
                                                 FullParseRecordsDict[i]['Event']['EventData']['Vbr3'][6:18])
                        else:
                            SN3 = '-'
                        EachPluggedDeviceDict[FullParseRecordsDict[i]['Event']['EventData']['SerialNumber']] = [
                            FullParseRecordsDict[i]['Event']['EventData']['Manufacturer'],
                            FullParseRecordsDict[i]['Event']['EventData']['Model'],
                            FullParseRecordsDict[i]['Event']['System']['TimeCreated']['#attributes']['SystemTime'].replace('T', ' ').replace('Z', ' UTC'),
                            FullParseRecordsDict[i]['Event']['System']['TimeCreated']['#attributes']['SystemTime'].replace('T', ' ').replace('Z', ' UTC'), [SN0, SN1, SN2, SN3]]
                else:
                    if FullParseRecordsDict[i]['Event']['EventData']['PartitionStyle'] == 1:
                        EachPluggedDeviceDict[FullParseRecordsDict[i]['Event']['EventData']['SerialNumber']][3] = \
                        FullParseRecordsDict[i]['Event']['System']['TimeCreated']['#attributes']['SystemTime'].replace('T', ' ').replace('Z', ' UTC')
                    else:
                        if FullParseRecordsDict[i]['Event']['EventData']['Vbr0'] != '':
                            SN0 = volumeSNParser(FullParseRecordsDict[i], 'Vbr0',
                                                 FullParseRecordsDict[i]['Event']['EventData']['Vbr0'][6:18])
                            if FullParseRecordsDict[i]['Event']['EventData']['Vbr1'] != '':
                                SN1 = volumeSNParser(FullParseRecordsDict[i], 'Vbr1',
                                                     FullParseRecordsDict[i]['Event']['EventData']['Vbr1'][6:18])
                            else:
                                SN1 = '-'
                            if FullParseRecordsDict[i]['Event']['EventData']['Vbr2'] != '':
                                SN2 = volumeSNParser(FullParseRecordsDict[i], 'Vbr2',
                                                     FullParseRecordsDict[i]['Event']['EventData']['Vbr2'][6:18])
                            else:
                                SN2 = '-'
                            if FullParseRecordsDict[i]['Event']['EventData']['Vbr3'] != '':
                                SN3 = volumeSNParser(FullParseRecordsDict[i], 'Vbr3',
                                                     FullParseRecordsDict[i]['Event']['EventData']['Vbr3'][6:18])
                            else:
                                SN3 = '-'
                            EachPluggedDeviceDict[FullParseRecordsDict[i]['Event']['EventData']['SerialNumber']][3] = \
                            FullParseRecordsDict[i]['Event']['System']['TimeCreated']['#attributes']['SystemTime'].replace('T', ' ').replace('Z', ' UTC')
                            if SN3 not in EachPluggedDeviceDict[FullParseRecordsDict[i]['Event']['EventData']['SerialNumber']][4]:
                                EachPluggedDeviceDict[FullParseRecordsDict[i]['Event']['EventData']['SerialNumber']][4].append(SN3)
                            if SN2 not in EachPluggedDeviceDict[
                                FullParseRecordsDict[i]['Event']['EventData']['SerialNumber']][4]:
                                EachPluggedDeviceDict[FullParseRecordsDict[i]['Event']['EventData']['SerialNumber']][4].append(SN2)
                            if SN1 not in EachPluggedDeviceDict[
                                FullParseRecordsDict[i]['Event']['EventData']['SerialNumber']][4]:
                                EachPluggedDeviceDict[FullParseRecordsDict[i]['Event']['EventData']['SerialNumber']][4].append(SN1)
                            if SN0 not in EachPluggedDeviceDict[
                                FullParseRecordsDict[i]['Event']['EventData']['SerialNumber']][4]:
                                EachPluggedDeviceDict[FullParseRecordsDict[i]['Event']['EventData']['SerialNumber']][4].append(SN0)
                        else:
                            continue

                # 분석한 보고서를 작성할 html파일에 적힐 코드

            FullParsehtml_code = '''<!DOCTYPE html> 
			<html>
			<head>
			<meta charset="utf-8" />
			<title> 분석 보고서 </title>
			<meta name="viewport" content="width=device-width, initial-scale=1">
			<link rel="stylesheet" type="text/css" media="screen" href="style.css" />
			</head>
			<body>
			<div class="wrapper">
			<div class="header">                        
			<H1> 분석 보고서 </H1>
			</div>'''

            FullParsehtml_code += f'\n<table align="center"> \n <caption><h2><b>연결된 모든 저장매체에 대한 분석 보고서</b></h2></caption>'
            FullParsehtml_code += '\n<tr style="background-color:DarkGrey"> \n <th>매체 시리얼 넘버</th> \n <th>제조사</th> \n <th>모델명</th> \n <th>최초 연결 시각 (UTC)</th> \n <th>마지막 연결 시각 (UTC)</th> \n <th>로그에서 추출된 vsn</th>'

            for serial in AllPluggedInSerials:
                FullParsehtml_code += f'\n<tr> \n <td>{serial}</td> \n<td>{EachPluggedDeviceDict[serial][0]}</td> \n <td>{EachPluggedDeviceDict[serial][1]}</td> \n<td>{EachPluggedDeviceDict[serial][2]}</td> \n<td>{EachPluggedDeviceDict[serial][3]} </td>\n <td>'
                for i in range(len(EachPluggedDeviceDict[serial][4])):
                    if EachPluggedDeviceDict[serial][4][i] == '-' or EachPluggedDeviceDict[serial][4][i] == 'Unknown Volume Type':
                        continue
                    FullParsehtml_code += f'{EachPluggedDeviceDict[serial][4][i]} '
                FullParsehtml_code += '</td>'

            FullParsehtml_code += f'\n</table> \n<br>\n<p align="center" style="color:white">\n <u style="font-size:20px"> 이벤트로그에 저장되어있는 최초와 마지막 타임라인\n</u> <br>\n 최초 시각: {FullParseLogStartTime} <br>\n마지막 시각: {FullParseLogEndTime} <br>\n</p> \n<br> <br> <br>\n<div class="push"></div> \n</div> \n <div class="footer">저장매체 로그 분석기 프로토타입</div>\n</body>\n</html>'

            # 분석 보고서의 css(style)파일에 적힐 코드

            FullParsecss_code = '''table{border-collapse:collapse;}
			th{text-align:center;background-color:#000000;color=white;}
			table,th,td{border:1px solid #000;}
			tr{text-align:center;background-color:#000000; color:white;}
			html, body {
			height: 100%;
			margin: 0;
			}
			.wrapper {
			min-height: 100%;
			background-color: #808080;
			margin-bottom: -50px;
			font-family: "Courier New", sans-serif;
			color=white;
			}
			.header{
			background-color: dark grey;
			color=white;
			}
			.header h1 {
			text-align: center;
			font-family: "Courier New", sans-serif;
			color=red;
			}
			.push {
			height: 50px;
			background-color: #808080;
			}
			.footer {
			height: 50px;
			background-color: #808080;
			color=white;
			text-align: center;
			}		'''

            # html, css파일을 생성하고, 위의 코드를 작성한 뒤 저장한다. (기본경로일 때와, 사용자 경로일때 각각 구분)
            if values['-INSAVE-'] == defaultOutputPathText:
                try:
                    with open('저장매체분석보고서.html', 'w', encoding='utf8') as fout:
                        fout.write(FullParsehtml_code)
                    with open('../style.css', 'w', encoding='utf8') as cssout:
                        cssout.write(FullParsecss_code)
                except:
                    FullParseHTMLWritten = False
            else:
                try:
                    with open(f"{values['-INSAVE-']}/저장매체분석보고서.html", 'w', encoding='utf8') as fout:
                        fout.write(FullParsehtml_code)
                    with open(f"{values['-INSAVE-']}/style.css", 'w', encoding='utf8') as cssout:
                        cssout.write(FullParsecss_code)
                except:
                    FullParseHTMLWritten = False
                # gui의 progress window에 띄울 메세지
            if FullParseHTMLWritten:
                if values['-DISKSN-'] == '':
                    print('파싱을 시작합니다.')
                    print('.................')
                    print('파싱 완료.')
                    print('.................')
                    print('분석을 시작합니다.')
                    print('.................')
                    print('분석 완료.')
                    print('.................')
                    print('---------------------------')
                    print('분석 결과')
                    print('---------------------------')
                    print()
                    print()
                    print(f'저장매체 이벤트로그의 타임라인입니다.')
                    print(f'최초 연결 시각: {FullParseLogStartTime}')
                    print(f'마지막 연결 시각: {FullParseLogEndTime}')
                    print()
                print('연결된 모든 장치에 대한 분석이 완료되었습니다.')
                sg.PopupOK('분석 보고서 작성이 성공적으로 완료되었습니다.', title='작성완료!', background_color='#000000')
            else:
                print('보고서 작성에 오류가 발생했습니다.\n저장하려는 폴더의 권한을 확인하세요.')
        else:
            sg.PopupOK('PartitionDiagnostic 로그 파일이 아닌 다른 로그 파일입니다.', title='!!!', background_color='#000000')
            window['-IN-'].update('')
            window['-DISKSN-'].update('')
    except Exception as e:
        print(e)
        sg.PopupOK('로그를 분석하는데 오류가 발생했습니다.', title='Error', background_color='#000000')
        window['-IN-'].update('')
        window['-DISKSN-'].update('')

            # gui - 메뉴바 레이아웃
menu_def = [['파일', ['종료']],
            ['도움', ['자세히']], ]
            # gui - file input window와 버튼 레이아웃
DiskSNFrameLayout = [[sg.Text('시리얼 넘버 입력란', background_color='#000000')],
                     [sg.In(key='-DISKSN-')]]

InputFrameLayout = [[sg.Text('분석 대상 usb 이벤트 로그', background_color='#000000')],
                    [sg.In(key='-IN-', readonly=True, background_color='#808080'),
                     sg.FileBrowse(file_types=(('Windows Event Log', '*.evtx'),))]]
            # 초기값으로 기본 경로를 불러오되, 사용자가 원할시 파일탐색기 실행.
defaultOutputPathText = '기본 경로는 exe파일이 존재하는 디렉토리입니다.'
OutputSaveFrameLayout = [[sg.Text('보고서를 저장할 경로를 선택하세요.', background_color='#000000')],
                         [sg.In(key='-INSAVE-', readonly=True, background_color='#334147',
                                default_text=defaultOutputPathText, text_color='grey'),
                          sg.FolderBrowse(key='-SAVEBTN-', disabled=True, enable_events=True)]]

col_layout = [[sg.Frame('장치의 시리얼 넘버를 입력하세요.', DiskSNFrameLayout, background_color='#000000')],
              [sg.Frame('이벤트 로그파일을 첨부하세요.', InputFrameLayout, background_color='#000000', pad=((0, 0), (0, 65)))],
              [sg.Frame('분석 보고서를 저장할 경로를 입력하세요.', OutputSaveFrameLayout, background_color='#000000')],
              [sg.Checkbox('HTML형식 저장', background_color='#000000', enable_events=True, key='-HTMLCHK-'),
               sg.Checkbox('액셀 형식으로 저장합니다.', background_color='#000000', enable_events=True, key='-CSVCHK-')],
              [sg.Checkbox('연결된 모든 저장매체에 대한 보고서를 작성합니다.\n(시리얼 넘버를 입력한 파일을 포함합니다.)', background_color='#000000',
                           enable_events=True, key='-FULLCHK-')],
              [sg.Button('종료', size=(7, 1)), sg.Button('분석', size=(7, 1))]]

            # gui - progress window 디자인
layout = [[sg.Menu(menu_def, key='-MENUBAR-')],
          [sg.Column(col_layout, element_justification='c', background_color='#000000'),
           sg.Frame('분석 진행 현황',[[sg.Output(size=(70, 25),background_color='#13053b',text_color='#46f024')]],background_color='#000000')],
          [sg.Text('저장매체 로그 분석기 프로토타입', background_color='#000000', text_color='#b2c2bf')]]

window = sg.Window('저장매체 로그 분석기 프로토타입', layout, background_color='#000000')

            # 프로그램 실행
while True:
    event, values = window.read()
            # 종료 누를시 창 종료
    if event in (sg.WIN_CLOSED, '종료'):
        break
            # 메뉴바 기능 정의
    if event == '자세히':
        sg.PopupOK('저장매체 로그 분석기 프로토타입 \n\n 눈물 젖은 시스템보안 기말과제', title='정보',background_color='#000000')
            # gui - checkbox 이벤트
    if event == '-HTMLCHK-' or event == '-CSVCHK-' or event == '-FULLCHK-':
        if values['-HTMLCHK-'] == False and values['-CSVCHK-'] == False and values['-FULLCHK-'] == False:
            window['-SAVEBTN-'].update(disabled=True)
            window['-INSAVE-'].update(value=defaultOutputPathText, text_color='grey')
        else:
            window['-SAVEBTN-'].update(disabled=False)
            window['-INSAVE-'].update(text_color='#000000')
            # 분석 button event
    if event == "분석":
        if values['-IN-'] == '': # 로그파일란 빈칸일 시 오류반환
            sg.PopupOK('분석하려는 이벤트 로그파일을 첨부하세요!', title='오류!', background_color='#000000')
            window['-DISKSN-'].update('')
        else:
            if values['-DISKSN-'] == '':
                if values['-FULLCHK-']: # 모든 연결 매체에 대한 보고서 작성
                    fullparse() # def fullparse 실행
                else:               # 시리얼넘버란 빈칸일 시 오류반환
                    sg.PopupOK('분석하려는 장치의 시리얼 넘버를 입력하세요!', title='오류!', background_color='#000000')
                    window['-IN-'].update('')
            else:
                try:            # evtxparser 모듈로 로그 분석
                    filename = values['-IN-']
                    # onlyname = fullpath[fullpath.rfind('/')+1:]
                    parser = PyEvtxParser(filename)
                    records_dict = {}
                    serial = values['-DISKSN-']
                    isDiskPlugedin = False
                    isDiskMBR = False
                    DataListsList = []
                    AllSNs = []
                    CSVTicked = False
                    HTMLTicked = False
                    IsPartitionDiagnosticEVTX = True
                    # html파일 작성
                    html_code = '''<!DOCTYPE html>
					<html>
					<head>
					<meta charset="utf-8" />
					<title> 분석 보고서 </title>
					<meta name="viewport" content="width=device-width, initial-scale=1">
					<link rel="stylesheet" type="text/css" media="screen" href="style.css" />
					</head>
					<body>
					<div class="wrapper">
					<div class="header">                        
					<H1> 분석 보고서 </H1>
					</div>'''

                    html_code += f'\n<table align="center"> \n <caption><h2><b>Partition%4Diagnostic.evtx 분석 보고서<br> 장치 시리얼 넘버: {serial}</b></h2></caption>'
                    html_code += '\n<tr style="background-color:DarkGrey"> \n <th>이벤트 로그 ID</th> \n <th>장치 연결 시각 (UTC)</th> \n <th>제조사</th> \n <th>모델명</th> \n <th>볼륨1 SN</th> \n <th>|볼륨2 SN</th> \n <th>|볼륨3 SN</th> \n <th>|볼륨4 SN</th> \n <th>|플래그</th>'

                    for record in parser.records_json():
                        data = json.loads(record['data'])
                        records_dict[(data['Event']['System']['EventRecordID'])] = data
                        if data['Event']['System']['EventID'] != 1006:
                            IsPartitionDiagnosticEVTX = False
                    if IsPartitionDiagnosticEVTX:
                        print('파싱 시작')
                        print('.................')
                        LogStartTime = records_dict[1]['Event']['System']['TimeCreated']['#attributes'][
                            'SystemTime'].replace('T', ' ').replace('Z', ' UTC')
                        LogEndTime = records_dict[len(records_dict)]['Event']['System']['TimeCreated']['#attributes'][
                            'SystemTime'].replace('T', ' ').replace('Z', ' UTC')
                        print('파싱 완료')
                        print('.................')
                        print('분석 시작')
                        print('.................')

                        for i in sorted(records_dict.keys()):
                            if serial == records_dict[i]['Event']['EventData']['SerialNumber']:
                                isDiskPlugedin = True
                                if records_dict[i]['Event']['EventData']['Vbr0'] != '':
                                    isDiskMBR = True
                                    SN0ToCheck = volumeSNParser(records_dict[i], 'Vbr0',
                                                                records_dict[i]['Event']['EventData']['Vbr0'][6:18])
                                    if records_dict[i]['Event']['EventData']['Vbr1'] != '':
                                        SN1ToCheck = volumeSNParser(records_dict[i], 'Vbr1',
                                                                    records_dict[i]['Event']['EventData']['Vbr1'][6:18])
                                    else:
                                        SN1ToCheck = '-'
                                    if records_dict[i]['Event']['EventData']['Vbr2'] != '':
                                        SN2ToCheck = volumeSNParser(records_dict[i], 'Vbr2',
                                                                    records_dict[i]['Event']['EventData']['Vbr2'][6:18])
                                    else:
                                        SN2ToCheck = '-'
                                    if records_dict[i]['Event']['EventData']['Vbr3'] != '':
                                        SN3ToCheck = volumeSNParser(records_dict[i], 'Vbr3',
                                                                    records_dict[i]['Event']['EventData']['Vbr3'][6:18])
                                    else:
                                        SN3ToCheck = '-'
                                    break
                        print('분석 완료')
                        print('.................')
                        print('---------------------------')
                        print('분석 결과')
                        print('---------------------------')
                        print()
                        print(f'이벤트로그의 타임라인')
                        print(f'최초 연결 시각: {LogStartTime}')
                        print(f'마지막 연결 시각: {LogEndTime}')
                        print()
                        if isDiskPlugedin == False:
                            print(f'다음의 시리얼 넘버를 가진 장치를 찾지 못했습니다. 시리얼 넘버 : {serial} ')
                            if values['-HTMLCHK-'] or values['-CSVCHK-']:
                                print('보고서를 작성하는데 오류가 발생했습니다.')
                            if values['-FULLCHK-']:
                                fullparse()
                        elif isDiskMBR == False:
                            print(f'다음의 시리얼 넘버를 가진 장치를 찾지 못했습니다. 시리얼 넘버 : {serial} ')
                            if values['-HTMLCHK-'] or values['-CSVCHK-']:
                                print('보고서를 작성하는데 오류가 발생했습니다.')
                            if values['-FULLCHK-']:
                                fullparse()
                        else:
                            for i in sorted(records_dict.keys()):
                                if serial == records_dict[i]['Event']['EventData']['SerialNumber']:
                                    if records_dict[i]['Event']['EventData']['PartitionStyle'] == 1:
                                        DataList = []
                                        DataList.append(records_dict[i]['Event']['System']['EventRecordID'])
                                        DataList.append(records_dict[i]['Event']['System']['TimeCreated']['#attributes']['SystemTime'].replace('T', ' ').replace('Z', ' UTC'))
                                        DataList.append(records_dict[i]['Event']['EventData']['Manufacturer'])
                                        DataList.append(records_dict[i]['Event']['EventData']['Model'])
                                        DataList.append('GPT partitioning scheme - No VSN info')
                                        DataList.append('GPT partitioning scheme - No VSN info')
                                        DataList.append('GPT partitioning scheme - No VSN info')
                                        DataList.append('GPT partitioning scheme - No VSN info')
                                        DataListsList.append(DataList)
                                    else:
                                        if records_dict[i]['Event']['EventData']['Vbr0'] == '' and \
                                                records_dict[i]['Event']['EventData']['Vbr1'] == '' and \
                                                records_dict[i]['Event']['EventData']['Vbr2'] == '' and \
                                                records_dict[i]['Event']['EventData']['Vbr3'] == '':
                                            continue
                                        else:
                                            DataList = []
                                            DataList.append(records_dict[i]['Event']['System']['EventRecordID'])
                                            DataList.append(records_dict[i]['Event']['System']['TimeCreated']['#attributes']['SystemTime'].replace('T', ' ').replace('Z', ' UTC'))
                                            DataList.append(records_dict[i]['Event']['EventData']['Manufacturer'])
                                            DataList.append(records_dict[i]['Event']['EventData']['Model'])
                                            SN0 = volumeSNParser(records_dict[i], 'Vbr0', records_dict[i]['Event']['EventData']['Vbr0'][6:18])
                                            DataList.append(SN0)
                                            if SN0 not in AllSNs: AllSNs.append(SN0)
                                            if records_dict[i]['Event']['EventData']['Vbr1'] == '':
                                                DataList.append('-')
                                            else:
                                                SN1 = volumeSNParser(records_dict[i], 'Vbr1',
                                                                     records_dict[i]['Event']['EventData']['Vbr1'][6:18])
                                                DataList.append(SN1)
                                                if SN1 not in AllSNs: AllSNs.append(SN1)
                                            if records_dict[i]['Event']['EventData']['Vbr2'] == '':
                                                DataList.append('-')
                                            else:
                                                SN2 = volumeSNParser(records_dict[i], 'Vbr2',
                                                                     records_dict[i]['Event']['EventData']['Vbr2'][6:18])
                                                DataList.append(SN2)
                                                if SN2 not in AllSNs: AllSNs.append(SN2)
                                            if records_dict[i]['Event']['EventData']['Vbr3'] == '':
                                                DataList.append('-')
                                            else:
                                                SN3 = volumeSNParser(records_dict[i], 'Vbr3',
                                                                     records_dict[i]['Event']['EventData']['Vbr3'][6:18])
                                                DataList.append(SN3)
                                                if SN3 not in AllSNs: AllSNs.append(SN3)
                                        DataListsList.append(DataList)

                            for DList in DataListsList:
                                if DList[4] == 'GPT partitioning scheme - No VSN info':
                                    html_code += f'\n<tr> \n <td>{DList[0]}</td> \n <td>{DList[1]}</td> \n<td>{DList[2]}</td> \n<td>{DList[3]} </td> \n <td>{DList[4]}</td> \n<td>{DList[5]}</td>\n <td>{DList[6]}</td> \n<td>{DList[7]}</td> \n<td> GPT media </td>'
                                elif DList[4] == SN0ToCheck and DList[5] == SN1ToCheck and DList[6] == SN2ToCheck and \
                                        DList[7] == SN3ToCheck:
                                    html_code += f'\n<tr> \n <td>{DList[0]}</td> \n <td>{DList[1]}</td> \n<td>{DList[2]}</td> \n<td>{DList[3]} </td> \n <td>{DList[4]}</td> \n<td>{DList[5]}</td>\n <td>{DList[6]}</td> \n<td>{DList[7]}</td> \n<td> - </td>'
                                else:
                                    html_code += f'\n<tr style="background-color:#000000"> \n <td>{DList[0]}</td> \n <td>{DList[1]}</td> \n<td>{DList[2]}</td> \n<td>{DList[3]} </td> \n <td>{DList[4]}</td> \n<td>{DList[5]}</td>\n <td>{DList[6]}</td> \n<td>{DList[7]}</td> \n<td> <b>VSN Change - Possible Format Action</b> </td>'
                                    SN0ToCheck = DList[4]
                                    SN1ToCheck = DList[5]
                                    SN2ToCheck = DList[6]
                                    SN3ToCheck = DList[7]
                                # html형식 저장 체크했을 시, 웹 페이지의 html,css 생성
                            if values['-HTMLCHK-']:
                                HTMLTicked = True
                                HTMLWritten = True
                                html_code += f'\n</table> \n<br>\n<p align="center" style="color:white">\n <u style="font-size:20px"> 이벤트로그 파일의 전체 분석 시각</u> <br>\n 최초 연결 시각: {LogStartTime} <br>\n마지막 연결 시각: {LogEndTime} <br>\n</p> \n<br> <br> <br> \n<div class="push"></div> \n</div> \n <div class="footer">저장매체 로그 분석기 프로토타입</div>\n</body>\n</html>'
                                css_code = '''table{border-collapse:collapse;}
								th{text-align:center;background-color:#ffffff;color=white;}
								table,th,td{border:1px solid #000;}
								tr{text-align:center;background-color:#000000; color:white;}
								html, body {
								height: 100%;
								margin: 0;
								}
								.wrapper {
								min-height: 100%;
								background-color: #808080;
								margin-bottom: -50px;
								font-family: "Courier New", sans-serif;
								color=white;
								}
								.header{
								background-color: dark grey;
								color=white;
								}
								.header h1 {
								text-align: center;
								font-family: "Courier New", sans-serif;
								color=red;
								}
								.push {
								height: 50px;
								background-color: #808080;
								}
								.footer {
								height: 50px;
								background-color: #808080;
								color=white;
								text-align: center;
								}		'''
                                # 기본 경로일시 저장
                                if values['-INSAVE-'] == defaultOutputPathText:
                                    try:
                                        with open(f'{serial}_보고서.html', 'w', encoding='utf8') as fout:
                                            fout.write(html_code)
                                        with open('../style.css', 'w', encoding='utf8') as cssout:
                                            cssout.write(css_code)
                                    except:
                                        HTMLWritten = False
                                # 사용자 지정 경로일시 저장
                                else:
                                    try:
                                        with open(f"{values['-INSAVE-']}/{serial}_보고서.html", 'w',
                                                  encoding='utf8') as fout:
                                            fout.write(html_code)
                                        with open(f"{values['-INSAVE-']}/style.css", 'w', encoding='utf8') as cssout:
                                            cssout.write(css_code)
                                    except:
                                        HTMLWritten = False
                                # csv로 저장 체크했을때, csv 저장 (기본 경로)
                            if values['-CSVCHK-']:
                                CSVTicked = True
                                CSVWritten = True
                                if values['-INSAVE-'] == defaultOutputPathText:
                                    try:
                                        with open(f'{serial}_보고서.csv', 'w', encoding='utf8') as fout:
                                            for DList in DataListsList:
                                                fout.write(
                                                    f'{DList[0]},{DList[1]},{DList[2]},{DList[3]},{DList[4]},{DList[5]},{DList[6]},{DList[7]}\n')
                                    except:
                                        CSVWritten = False
                                # csv로 저장 체크했을때, csv 저장 (사용자 지정 경로)
                                else:
                                    try:
                                        with open(f"{values['-INSAVE-']}/{serial}_보고서.csv", 'w',
                                                  encoding='utf8') as fout:
                                            for DList in DataListsList:
                                                fout.write(
                                                    f'{DList[0]},{DList[1]},{DList[2]},{DList[3]},{DList[4]},{DList[5]},{DList[6]},{DList[7]}\n')
                                    except:
                                        CSVWritten = False
                            print(f'{serial}의 시리얼 넘버를 가진 장치가 {len(DataListsList)}회 연결 되었습니다.')
                            print(f'{len(AllSNs)}개의 유일한 vsn을 보유하고 있습니다.')
                            for SN in AllSNs:
                                print(SN)
                            if HTMLTicked:
                                if HTMLWritten:
                                    print('HTML 보고서 작성에 성공했습니다.')
                                else:
                                    print('HTML 보고서 작성은 성공했지만 저장에 실패하였습니다.\n저장 폴더에 대한 권한을 확인해주세요!')
                            if CSVTicked:
                                if CSVWritten:
                                    print('CSV 보고서 작성에 성공했습니다.')
                                else:
                                    print('CSV 보고서 작성은 성공했지만 저장에 실패하였습니다.\n저장 폴더에 대한 권한을 확인해주세요!')
                            if values['-FULLCHK-']:
                                fullparse()
                            sg.PopupOK('분석이 성공적으로 완료되었습니다!\n디렉토리를 확인하세요.', title='분석 성공!', background_color='#000000')
                    else:
                        sg.PopupOK('USB에 관련된 이벤트 로그가 아닙니다!', title='!!!', background_color='#000000')
                        window['-IN-'].update('')
                        window['-DISKSN-'].update('')
                except:
                    sg.PopupOK('분석에 오류가 발생했습니다.', title='앗!', background_color='#000000')
                    window['-IN-'].update('')
                    window['-DISKSN-'].update('')

window.close()