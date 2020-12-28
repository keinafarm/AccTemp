# -*- coding: utf-8 -*-

from MeteorologicalAgency import MeteorologicalAgency
import pandas as pd
from datetime import date

"""
    気温データを扱う
    ・積算温度を計算する
    ・予測日付を計算する

"""


class TemperatureCalculator:
    calculator_list = {}  # 地域ごとの気象庁データオブジェクトを記憶する

    def __new__(cls, start_date, end_date, prefecture, station):
        """
        気象庁の温度データを管理するクラス
        同じ地域の気象庁データを作成しようとした場合、既に生成したオブジェクトを返す
        :param start_date:取得期間開始日
        :param end_date:取得期間終了日
        :param prefecture:県名
        :param station:地域名
        :return 温度データを管理するオブジェクト
        """
        index = prefecture + '_' + station  # 県名と地域名でindexを作る
        if index in TemperatureCalculator.calculator_list:  # 既に生成されているかどうかをチェック
            return TemperatureCalculator.calculator_list[index]  # 記憶しているオブジェクトを返す

        # 新たにオブジェクトを作る
        calculator = super(TemperatureCalculator, cls).__new__(cls)
        # 生成したオブジェクトを記憶する
        calculator = TemperatureCalculator.calculator_list.setdefault(index, calculator)
        return calculator

    def __init__(self, start_date, end_date, prefecture, station):
        """
        既に生成されたオブジェクトでの呼び出し時、既に読み込んだ温度データの期間をチェックし
        指定された期間のデータを既に読み込んでいれば、何もしない

        :param start_date:
        :param end_date:
        :param prefecture:
        :param station:
        """
        if hasattr(self, 'prefecture'):
            if self.check_period(start_date, end_date):
                return
            # 指定された期間と、既に読みこんでいる期間を比較し、より大きい範囲に補正する
            if start_date > self.start_date:
                start_date = start_date
            if end_date < self.end_date:
                end_date = self.end_date

        self.prefecture = prefecture
        self.station = station
        self.meteorological = MeteorologicalAgency(prefecture, station)  # 気象庁からデータを持ってくるオブジェクト
        self.start_date = start_date  # 指定された期間を覚えておく
        self.end_date = end_date
        temperature_list = self.meteorological.get_temperature_list(start_date, end_date)  # 平均気温と平年気温の時系列リストを得る
        # リストをPandas.DataFrameにする
        self.temperature_data = pd.DataFrame(temperature_list,
                                             columns=['date', 'mean temperature', 'average temperature'])
        self.temperature_data.set_index('date', inplace=True)  # 日付カラムをindexにする

    def get_accumulated_temperature(self, start, end):
        """
        積算温度を得る
        :param start: 開始日
        :param end: 終了日
        :return: 積算温度
        """
        out = self.temperature_data[start:end]['mean temperature'].sum()
        return out

    def get_average_temperature(self, specified_date):
        """
        平年温度を得る
        :param specified_date:取得する日にち
        :return:平年温度
        """
        out = self.temperature_data.at[pd.to_datetime(specified_date), 'average temperature']
        return out

    def check_period(self, start, end):
        """
        現在抱えている値が、指定した期間を含んでいるか
        :param start: 開始日
        :param end: 終了日
        :return: True = 含んでいる False = 含んでいない
        """
        if self.start_date is None:
            return False
        if self.start_date >= start:
            return False
        if self.end_date <= end:
            return False
        return True

    def forecasted_day(self, start, target_temperature):
        """
        積算温度が指定した温度になる日を予測する
        :param start: 積算開始日
        :param target_temperature: 目標積算温度
        :return:到達予測日
        """
        # 取得した日付よりも以前からの積算温度は勘弁してね
        if start < self.start_date:
            raise ValueError("取得した日以前の日付が指定されました。start is %s < self.start_date is %s" % (
                start.strftime("%Y/%m/%d"), self.start_date.strftime("%Y/%m/%d")))

        # 積算開始日から最終温度取得日までの積算温度を計算する
        accumulated_temperature = 0
        current_day = start
        for row in self.temperature_data[start:self.end_date].itertuples():
            current_day = row.Index
            temperature = row[1]  # 平均気温
            accumulated_temperature += temperature
            if accumulated_temperature >= target_temperature:
                break

#           print(current_day, temperature, accumulated_temperature)

        if accumulated_temperature >= target_temperature:
            return current_day.to_pydatetime()  # Pandas.Timestampをpython datetimeに変換して返す

        # 平年温度が取得できているかチェックする
        last_day = date(self.end_date.year - 1, self.end_date.month, self.end_date.day)
        if last_day < self.start_date:
            raise ValueError("平年温度を取得するデータがありません。last_day is %s < self.start_date is %s" % (
                last_day.strftime("%Y/%m/%d"), self.start_date.strftime("%Y/%m/%d")))

        # 積算温度が目標に届かなかった時の処理
        for row in self.temperature_data[last_day:self.end_date].itertuples():
            current_day = row.Index
            temperature = row[2]  # 平年温度
            accumulated_temperature += temperature
            if accumulated_temperature >= target_temperature:
                break

#           print(current_day, temperature, accumulated_temperature)

        return_day = date(current_day.year + 1, current_day.month, current_day.day)
        return return_day


if __name__ == "__main__":
    """
    obj = TemperatureCalculator("高知", "窪川")
    obj.load_temperature_data(date(2020, 7, 28), date(2020, 10, 5))
    temp_sum = obj.get_accumulated_temperature(date(2020, 7, 30), date(2020, 9, 30))
    print(temp_sum)

    temp_sum = obj.get_accumulated_temperature(date(2020, 8, 8), date(2020, 9, 16))
    print(temp_sum)

    temp_sum = obj.get_accumulated_temperature(date(2020, 8, 8), date(2020, 9, 20))
    print(temp_sum)

    temp_sum = obj.get_accumulated_temperature(date(2020, 8, 10), date(2020, 9, 14))
    print(temp_sum)

    temp_sum = obj.get_accumulated_temperature(date(2020, 7, 29), date(2020, 9, 9))
    print(temp_sum)

    temp_sum = obj.get_accumulated_temperature(date(2020, 8, 12), date(2020, 9, 26))
    print(temp_sum)

    temp_sum = obj.get_accumulated_temperature(date(2020, 8, 12), date(2020, 9, 27))
    print(temp_sum)

    temp_sum = obj.get_accumulated_temperature(date(2020, 8, 9), date(2020, 9, 15))
    print(temp_sum)

    temp_sum = obj.get_accumulated_temperature(date(2020, 8, 9), date(2020, 10, 3))
    print(temp_sum)

    temp_sum = obj.get_accumulated_temperature(date(2020, 8, 8), date(2020, 9, 20))
    print(temp_sum)

    temp_sum = obj.get_accumulated_temperature(date(2020, 8, 10), date(2020, 9, 15))
    print(temp_sum)

    temp_sum = obj.get_accumulated_temperature(date(2020, 7, 28), date(2020, 9, 17))
    print(temp_sum)

    temp_sum = obj.get_accumulated_temperature(date(2020, 8, 8), date(2020, 9, 21))
    print(temp_sum)

    temp_sum = obj.get_accumulated_temperature(date(2020, 8, 6), date(2020, 9, 25))
    print(temp_sum)

    print(obj.get_average_temperature(date(2020, 8, 1)))
    print(obj.get_average_temperature(date(2020, 8, 7)))
    print(obj.get_average_temperature(date(2020, 8, 9)))
    print(obj.get_average_temperature(date(2020, 8, 15)))
    """

    obj1 = TemperatureCalculator(date(2020, 1, 26), date(2020, 2, 26), "高知", "窪川")
    print(obj1)
    obj2 = TemperatureCalculator(date(2019, 12, 26), date(2020, 12, 26), "高知", "本山")
    print(obj2)
    obj3 = TemperatureCalculator(date(2020, 1, 27), date(2020, 2, 15), "高知", "窪川")
    print(obj3)
    obj4 = TemperatureCalculator(date(2019, 1, 26), date(2020, 2, 15), "高知", "窪川")
    print(obj4)
    obj5 = TemperatureCalculator(date(2019, 1, 26), date(2020, 12, 26), "高知", "窪川")
    print(obj5)

    day = obj5.forecasted_day(date(2020, 12, 25), 1000)
    print("ret:" + day.strftime("%Y/%m/%d"))

    day = obj5.forecasted_day(date(2020, 8, 1), 1000)
    print("ret:" + day.strftime("%Y/%m/%d"))
