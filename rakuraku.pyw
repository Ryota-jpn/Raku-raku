import PySimpleGUI as sg
import datetime
import jpholiday
import calendar

sg.theme("Default")

categoeyChoices = ["片道", "往復", "--"]
seikyuChoices = ["なし", "あり"]
reasonChoices = ["通勤費(通常勤務地)",
                 "在宅チャージ",
                 "営業関連面談／AMU・顧客",
                 "社内面談／人事評価制度",
                 "社内面談／その他※要備考入力",
                 "現場業務(通常勤務地以外)",
                 "社内業務／ハモデイ・会議等",
                 "社内業務／総会・健康診断",
                 "内勤Unit／採用関係",
                 "内勤Unit／その他※要備考入力",
                 "その他※要備考入力"]
wayChoices = ["--", "電車（国内）", "バス（国内）", "その他※原則使用不可"]

layout = [[sg.Text("  楽々清算用CSVファイル作成アプリ", font=(None,30))],
          [sg.Text()],
          [sg.Text("      作成する年月を指定してください（yyyy/mm）*")],
          [sg.Text("    "), sg.InputText("",key="date")],
          [sg.Text("      出発地を入力してください")],
          [sg.Text("    "), sg.InputText("",key="start")],
          [sg.Text("      到着地を入力してください")],
          [sg.Text("    "), sg.InputText("",key="end")],
          [sg.Text("      往復/片道を入力してください")],
          [sg.Text("    "), sg.OptionMenu(categoeyChoices, key="category", size=(10), default_value="片道")],
          [sg.Text("      金額（片道）を入力してください")],
          [sg.Text("    "), sg.InputText("0",key="amount", justification="right")],
          [sg.Text("      客先請求の有無を入力してください *")],
          [sg.Text("    "), sg.OptionMenu(seikyuChoices, key="seikyu", size=(10))],
          [sg.Text("      申請理由を入力してください *")],
          [sg.Text("    "), sg.OptionMenu(reasonChoices, key="reason", size=(30))],
          [sg.Text("      交通機関を入力してください")],
          [sg.Text("    "), sg.OptionMenu(wayChoices, key="way", size=(20), default_value="--")],
          [sg.Text("      備考を入力してください")],
          [sg.Text("    "), sg.InputText("",key="memo")],
          [sg.Text("      保存するファイルを選択してください（csv形式のみ）")],
          [sg.Text("    "), sg.InputText(), sg.FileSaveAs("ファイルを選択", key="file")],
          [sg.Text("",key="result")],
          [sg.Text("")],
          [sg.Text("                 "), sg.Button("CSVファイル作成", k="btn", size=(40)), sg.Text("     "), sg.Button("終了", key="endBtn")]]

win = sg.Window("楽々清算CSV作成アプリ", layout, font=(None, 14), size=(700,850))

def isBizDay(date):
    date = datetime.date(int(date[0:4]), int(date[4:6]), int(date[6:8]))
    if date.weekday() >= 5 or jpholiday.is_holiday(date):
        return 0
    else:
        return 1

def createCsv():
    date = value["date"]
    year = date[:4]
    month = date[5:]
    start = value["start"]
    end = value["end"]
    category = value["category"]
    amount = value["amount"]
    seikyu = value["seikyu"]
    reason = value["reason"]
    way = value["way"]
    memo = value["memo"]
    filePath = value["file"]
    header = "日付,出発,到着,往復,金額/Km,客先請求,申請理由,交通機関,備考\n"
    firstLoop = True

    if checkValue(date, seikyu, reason, amount) or checkFile(filePath):
        return

    filePath = fileFormat(filePath)
    dateList = []
    dateYM = calendar.monthcalendar(int(year), int(month))

    for d1 in dateYM:
        for d2 in d1:
            if d2 != 0:
                if len(str(d2)) < 2:
                    d2 = "0" + str(d2)
                if len(str(month)) < 2:
                    month = "0" + str(month)
                    dateList.append(f"{year}{month}{d2}")
                else:
                    dateList.append(f"{year}{month}{d2}")

    file = open(filePath, "w", encoding="UTF-8")

    print(dateList)

    for days in dateList:
        if firstLoop:
            file.write(header)
            firstLoop = False
        if isBizDay(days):
            exp = f"{days[:4]}/{days[4:6]}/{days[6:]},{start},{end},{category},{amount},{seikyu},{reason},{way},{memo}\n"
            file.write(exp)

    file.close()

    resultMsg()

def checkValue(date, seikyu, reason, amount):
    errCount = 0
    if not date:
        sg.PopupTimed("作成する年月を入力してください。")
        errCount += 1
    if not seikyu:
        sg.PopupTimed("客先請求の有無を入力してください。")
        errCount += 1
    if not reason:
        sg.PopupTimed("申請理由を入力してください。")
        errCount += 1
    if reason == "在宅チャージ":
        if amount != "250":
            sg.PopupTimed("申請理由が「在宅チャージ」の場合は金額は250円のみです。")
            errCount += 1
    return errCount

def checkFile(filePath):
    errCount = 0
    if not filePath:
        sg.PopupTimed("ファイル名を入力してください。")
        errCount += 1
    return errCount

def fileFormat(filePath):
    if filePath.find(".") == -1:
        filePath += ".csv"
        return filePath
    return filePath

def resultMsg():
    msg = "      csvが作成されました。"
    win["result"].update(msg)

while True:
    event, value = win.read()
    if event == "btn":
        createCsv()
    if event == "endBtn":
        break
    if event == sg.WIN_CLOSED:
        break

win.close()