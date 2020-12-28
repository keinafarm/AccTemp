from ATWorkBook import ATWorkBook
import datetime as dt
from TemperatureCalculator import TemperatureCalculator

"""
    Z-GIS用のExcelファイル中から、積算温度の管理データを見つけて、積算温度を記入する

    Usage:
        python AccTemp <Excelデータファイル名> <県名> <地区名>

    1) 積算温度の計算対象となるシートは、シート名に「積算温度」を含める
    2) 2行目に、"開始日", の文字列があるカラムは必須
    3) 2行目に、"終了日", "目標積算温度", "現状積算温度", "目標到達度", "予測終了日"の各タイトルが出てくるカラムはオプション
    4) 気象庁の"http://www.data.jma.go.jp/gmd/risk/obsdl/index.php"から、指定期間の気温データを取得し、積算温度他を下記のルールで算出する
        開始日:☓  -> 無視
        開始日:○ 終了日:☓ or 未達 目標積算温度:☓ -> ( 開始日：前日まで）現状積算温度を計算
        開始日:○ 終了日:☓ or 未達 目標積算温度:○ -> (開始日：前日まで）現状積算温度,目標到達度,予測終了日を計算
        開始日:○ 終了日:○ and 過去日 目標積算温度:○ 現状積算温度:☓-> (開始日：終了日まで）現状積算温度,目標到達度,(開始日：前日まで）予測終了日を計算
        開始日:○ 終了日:○ and 過去日 目標積算温度:☓ 現状積算温度:○-> 無視
        開始日:○ 終了日:○ and 過去日 目標積算温度:☓ 現状積算温度:☓-> (開始日：終了日まで）現状積算温度
        開始日:○ 終了日:○ and 過去日 目標積算温度:○ 現状積算温度:○  予測終了日:☓ or 未達-> (開始日：終了日まで）現状積算温度,目標到達度,(開始日：前日まで）予測終了日を計算
        開始日:○ 終了日:○ and 過去日 目標積算温度:○ 現状積算温度:○  予測終了日:○ and 過去日-> 無視

"""


####################
#
#   積算温度を計算する部分
#
####################

def data_operator(field_data, accumulated_temperature_function, estimate_date_function):
    """
        現状積算温度,目標到達度,予測終了日を計算

        開始日:☓  -> 無視
        開始日:○ 終了日:☓ or 未達 目標積算温度:☓ -> ( 開始日：前日まで）現状積算温度を計算
        開始日:○ 終了日:☓ or 未達 目標積算温度:○ -> (開始日：前日まで）現状積算温度,目標到達度,予測終了日を計算
        開始日:○ 終了日:○ and 過去日 目標積算温度:○ 現状積算温度:☓-> (開始日：終了日まで）現状積算温度,目標到達度,(開始日：前日まで）予測終了日を計算
        開始日:○ 終了日:○ and 過去日 目標積算温度:☓ 現状積算温度:○-> 無視
        開始日:○ 終了日:○ and 過去日 目標積算温度:☓ 現状積算温度:☓-> (開始日：終了日まで）現状積算温度
        開始日:○ 終了日:○ and 過去日 目標積算温度:○ 現状積算温度:○  予測終了日:☓ or 未達-> (開始日：終了日まで）現状積算温度,目標到達度,(開始日：前日まで）予測終了日を計算
        開始日:○ 終了日:○ and 過去日 目標積算温度:○ 現状積算温度:○  予測終了日:○ and 過去日-> 無視

    :param field_data:計算元ネタ(各圃場分のデータ)
    :param accumulated_temperature_function:積算温度計算関数
    :param estimate_date_function:到達日予測関数
    :return:
    """
    today = dt.datetime.now()  # 今日の日時
    yesterday = today - dt.timedelta(days=1)  # 昨日の日時

    if field_data.start_date is None:
        return  # 開始日:☓  -> 無視
    if field_data.end_date is None or field_data.end_date >= today:  # 終了日が無いかまだ到達していない場合

        current_temperature = accumulated_temperature_function(field_data.start_date, yesterday)
        field_data.current_temperature = current_temperature
        # field_data.current_temperatureに直接値を入れると、読み出し時Noneになる可能性がある
        if field_data.target_temperature is None:
            return  # 開始日:○ 終了日:☓ or 未達 目標積算温度:☓ -> ( 開始日：前日まで）現状積算温度を計算
        else:
            field_data.rate = current_temperature / field_data.target_temperature
            field_data.estimate_date = estimate_date_function(field_data.start_date, field_data.target_temperature)
            # 開始日:○ 終了日:☓ or 未達 目標積算温度:○ -> (開始日：前日まで）現状積算温度,目標到達度,予測終了日を計算

    elif field_data.target_temperature is None:
        if field_data.current_temperature is not None:
            return  # 開始日:○ 終了日:○ and 過去日 目標積算温度:☓ 現状積算温度:○-> 無視
        else:
            field_data.current_temperature = accumulated_temperature_function(field_data.start_date,
                                                                              field_data.end_date)
            return  # 開始日:○ 終了日:○ and 過去日 目標積算温度:☓ 現状積算温度:☓-> (開始日：終了日まで）現状積算温度

    elif field_data.estimate_date is not None and field_data.estimate_date < today:
        return  # 開始日:○ 終了日:○ and 過去日 目標積算温度:○ 現状積算温度:○  予測終了日:○ and 過去日-> 無視
    else:
        # 開始日:○ 終了日:○ and 過去日 目標積算温度:○ 現状積算温度:○  予測終了日:☓ or 未達-> (開始日：終了日まで）現状積算温度,目標到達度,(開始日：前日まで）予測終了日を計算
        current_temperature = accumulated_temperature_function(field_data.start_date, field_data.end_date)
        field_data.current_temperature = current_temperature
        # field_data.current_temperatureに直接値を入れると、読み出し時Noneになる可能性がある
        field_data.rate = current_temperature / field_data.target_temperature
        field_data.estimate_date = estimate_date_function(field_data.start_date, field_data.target_temperature).date()


####################
#
#   Main class
#
####################

class AccTemp:
    def __init__(self, data_file, prefecture, station):
        """
        メインクラス
        :param data_file: Excelデータファイル名
        :param prefecture: 県名
        :param station: 地区名
        """
        self.data_file = data_file
        self.prefecture = prefecture
        self.station = station

    def run(self):
        # Excelファイルを読み込む
        user_data = ATWorkBook(self.data_file)

        ####################
        #
        #   取得期間は、Excelファイル内の開始日のうち一番古い日付、から、終了日の一番新しい日付の間になる
        #   但し、積算満了日の予測を行う場合、平年温度は、一年前の日付から持ってこなければならないので、
        #   少なくとも1年分のデータを取得する必要がある（正確にやれば、１年はいらないかもだけど、平年温度は１年以上は必要ない）
        #   したがって、Excelファイル内の一番古い開始日が１年前より後であれば、１年前を開始日にする
        #   １年分のデータを取得するので、終了日は、少なくとも開始日の一年後よりもあとにする
        #
        ####################

        # 一年前の日付を得る
        yesterday = dt.datetime.now() - dt.timedelta(days=1)
        one_year_ago = dt.datetime(yesterday.year - 1, yesterday.month, yesterday.day)

        # 気温の取得範囲を決定する
        start_date = one_year_ago  # 開始日の最大値:少なくとも、今日から一年前
        end_date = dt.datetime(1960, 1, 1)  # 終了日の最小値:一番古い日
        for sheet in user_data.sheets_list:  # すべてのシート分
            for data in sheet.data_list:  # 全データ分
                if data.start_date is not None and data.start_date < start_date:
                    start_date = data.start_date  # 開始日は、より小さい値を採用
                if data.end_date is not None and data.end_date > end_date:
                    end_date = data.end_date  # 終了日は、より大きい値を採用

        # 取得開始日の１年後の日付を得る
        one_year_later = dt.datetime(start_date.year + 1, start_date.month, start_date.day)

        # 取得期間が１年に満たなかったら、取得期間を１年間にする
        if end_date < one_year_later:
            end_date = one_year_later

        print("start:" + start_date.strftime("%Y/%m/%d"))
        print("end:" + end_date.strftime("%Y/%m/%d"))

        # 予め、計算期間をloadしておく
        temperature_calculator = TemperatureCalculator(start_date, end_date, self.prefecture, self.station)

        for sheet in user_data.sheets_list:  # すべてのシート分
            for data in sheet.data_list:  # 全データ分
                data_operator(data, temperature_calculator.get_accumulated_temperature,
                              temperature_calculator.forecasted_day)

        user_data.flash()  # Excelファイルに保存


if __name__ == "__main__":
    import sys

    args = sys.argv
    if len(args) != 4:
        print("パタメータは3つです (%d)" % (len(args) - 1))
        print("Usage:")
        print(" AccTemp Excelファイル名 県名 地区名")
        sys.exit(1)

    obj = AccTemp(args[1], args[2], args[3])
    obj.run()
