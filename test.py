import win32com.client
from pykrx import stock
from datetime import datetime, timedelta

# Excel 엑세스
excel = win32com.client.Dispatch("Excel.Application")  # 엑셀 프로그램 실행
excel.Visible = True  # 엑셀을 보이게 설정

# "종목분석" 시트 선택
ws = excel.ActiveWorkbook.Worksheets("종목분석")
ws.Activate()

# "종목코드" 위치 찾기
cell = ws.Cells.Find(What="종목코드")
start_row = cell.Row + 1  # "종목코드" 아래 행부터 시작
col_index = cell.Column

# 어제 날짜 계산
yesterday = (datetime.now() - timedelta(1)).strftime('%Y%m%d')

# 어제 날짜의 주식 시장에서 거래된 종목 코드와 종목명 가져오기
tickers = stock.get_market_ticker_list(date=yesterday, market="ALL")
total_tickers = len(tickers)

# 엑셀 입력 전 메시지 출력
print("엑셀 입력 중...")

# 각 종목의 정보 입력
for i, ticker in enumerate(tickers):
    try:
        name = stock.get_market_ticker_name(ticker)
        closing_price = int(stock.get_market_ohlcv_by_date(fromdate=yesterday, todate=yesterday, ticker=ticker).iloc[0]['종가'])
        fundamental = stock.get_market_fundamental_by_ticker(yesterday, market='ALL').loc[ticker]

        row = start_row + i
        ws.Cells(row, col_index).Value = ticker
        ws.Cells(row, col_index + 1).Value = name
        ws.Cells(row, col_index + 2).Value = closing_price
        ws.Cells(row, col_index + 3).Value = int(fundamental['BPS'])
        ws.Cells(row, col_index + 4).Value = fundamental['PER']
        ws.Cells(row, col_index + 5).Value = fundamental['PBR']
        ws.Cells(row, col_index + 6).Value = int(fundamental['EPS'])
        ws.Cells(row, col_index + 7).Value = fundamental['DIV']
        ws.Cells(row, col_index + 8).Value = int(fundamental['DPS'])
    except KeyError:
        print(f"데이터가 없는 종목코드: {ticker}")
        continue

    # 완료율 출력
    completion_percentage = ((i + 1) / total_tickers) * 100
    print(f"진행 상태: {completion_percentage:.2f}% 완료")

# 변경 사항 저장하고 Excel 닫기
excel.ActiveWorkbook.Save()
# excel.Quit()

print("엑셀 입력 완료.")
