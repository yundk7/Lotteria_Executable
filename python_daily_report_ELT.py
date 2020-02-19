import pandas as pd
import datetime
from datetime import timedelta
import glob

if __name__ == "__main__" :

#     today = pd.to_datetime("today")
#     today = (datetime.datetime.today()-datetime.timedelta(days=1)).strftime("%Y-%m-%d")
#     ## datetime.timedelta(days=1)##
#     today
    df = pd.read_excel("시간대별 매출 조회.xls",nrows=2)
    today = df.iloc[-1][0][9:19]

    #바로고 배달완료 건수

#     delivery = pd.read_csv("접수현황.csv",encoding = "euc_kr")

    for f in glob.glob("접수현황*.csv"):
        delivery = pd.read_csv(f, encoding = "euc_kr")
    
    delivery = delivery[delivery["진행상황"]=="완료"]
    
    delivery["진행시간"] = pd.to_datetime(delivery["진행시간"])

    del_complete = len(delivery[delivery["진행시간"].dt.normalize() == today])

    del_complete

    sales = pd.read_excel("일자별 결제수단 판매형태 매출 조회.xls",skiprows = 7)
    sales = sales[sales["일자"]==today]

    types = list(filter(lambda x: ".1" in x, sales.columns))

    sales_df = sales[types].T
    sales_df.index = sales_df.index.str.replace(".1","")
    sales_df.columns = ["금액"]

    sales_df["객수"] = sales[sales_df.index].T
    sales_df

#     selfcard = open("자체배달카드.txt",'r')
#     selfcard = selfcard.readlines()
#     selfcard = int(selfcard[0])
    selfcard = pd.read_excel("일자별 결제수단 판매형태 매출 조회.xls",nrows=1)
    selfcard = selfcard.iloc[0][0]

    total = sales_df["금액"].T["판매"]
    cash = sales_df["금액"].T["현금"]# - selfcard
    cash

    lottept = sales_df["금액"].T[["롯데포인트","제휴포인트"]].sum()

    coupon = sales_df["금액"].T[["제품교환권회수","자사 상품권","타사 상품권","모바일쿠폰","선불카드"]].sum()
    card = sales_df["금액"].T["신용카드"] + selfcard
    card

    df11 = sales_df.T[["반품","급식","폐기"]].T.reset_index()
    bigo = df11["index"].astype(str) + " : " + df11["금액"].astype(str) + " : " + df11["객수"].astype(str)

    #홈서비스 구분
    homes = pd.read_excel("시간대별 홈서비스 매출 실적 조회.xls",skiprows=5)

    types = list(filter(lambda x: ".1" in x, homes.columns))

    df = homes[types].iloc[25][2:13]

    df.index = df.index.str.replace(".1","")
    homes_df = pd.DataFrame(df)
    homes_df = homes_df[homes_df[25]!=0]
    homes_df.rename(columns = {25:"금액"},inplace = True)

    homes_df["객수"] = homes[homes_df.index].iloc[-1]
    homes_df

    total_del = homes_df["금액"].sum()
    total_del

    homes_df["금액"].T["자체배달"]

    self_del = homes_df["금액"].T["자체배달"]
    self_del

    rest_del = total_del-self_del

    homesn = homes_df["객수"].sum()
    homesn
    self_del_n = homes_df["객수"].T["자체배달"]
    self_del_n

    df10 = homes_df.reset_index()
    bigo = bigo.append(df10["index"].astype(str) + " : " + df10["금액"].astype(str) + " : " + df10["객수"].astype(str))
    bigo.reset_index(drop=True,inplace=True)

    #심야매출
    hourly = pd.read_excel("시간대별 매출 조회.xls",skiprows=8)
    night = hourly["실적"][0:4].append(hourly["실적"][9:10]).append(hourly["실적"][22:24])
    night = night.sum()
    night
    orders = hourly[24:]["객수"].item()

    cash_receipt = pd.read_excel("현금영수증 거래내역.xls",skiprows = 5)
    cash_receipt = cash_receipt[cash_receipt["승인일자"]==today]["승인금액"].sum()
    cash_receipt

    df1 = pd.DataFrame({
        "총매출":[total],
        "현금매출":[cash],
        "카드매출":[card],
        "롯데포인트":[lottept],
        "제품교환권+모바일": [coupon],
        "심야매출":[night],
        "현금영수증": [cash_receipt],
        "자체배달": [self_del],
        "콜센터": [rest_del],
        "자체배달 개수": [self_del_n],
        "콜센터 개수":[homesn-self_del_n],
        "객수": [orders],
        "객단가":[int(total / orders)],
        "배달대행 건수": [del_complete],
        "배달주문 건수": [homesn],
        "배달 차이 건수": [homesn-del_complete]
    })

    df1 = df1.T.reset_index()
    df1 = pd.concat([df1,bigo],axis=1,ignore_index=True).apply(lambda x: pd.Series(x.dropna().values))
    # df1 = pd.concat
    df1.to_csv("영업보고(당일).csv",index=False,encoding="euc_kr")
