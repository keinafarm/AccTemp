# -*- coding: utf-8 -*-

"""
https://gist.github.com/barusan/3f098cc74b92fad00b9bb4478da35385
を参照した

気象庁から平均気温と平年気温のデータを取得する

"""
from datetime import date
import urllib.request
import lxml.html
from datetime import datetime as dt


def encode_data(data):
    """
    Map型のオブジェクトをパーセントエンコードされた ASCII 文字列に変換する
    :param data: 変換するmap型オブジェクト
    :return: 変換したURL文字列
    """
    return urllib.parse.urlencode(data).encode(encoding='ascii')


def get_phpsessid():
    """
    セッションIDを得る
    気象庁とやりとりする為のPHPSESSIDを取得する
    :return:PHPSESSID
    """

    URL = "http://www.data.jma.go.jp/gmd/risk/obsdl/index.php"  # セッションIDEAを取得する為のURL
    xml = urllib.request.urlopen(URL).read().decode("utf-8")
    tree = lxml.html.fromstring(xml)
    return tree.cssselect("input#sid")[0].value


def get_station(prefecture="高知", station="窪川"):
    """
    地点のstation IDを得る
    :param prefecture:気象庁に通知する県名
    :param station:気象庁に通知する地区名
    :return:
    """

    def get_station_by_id(pd):
        """
        指定したIDのSTATION ID Listを得る
        :param pd: 0:全国 0以外:station id listを得る県のid
        :return: station id list
        """

        def kansoku_items(bits):
            """
            取得する観測データ：気温のみ
            :param bits:
            :return:
            """
            return dict(rain=(bits[0] == "1"),
                        wind=(bits[1] == "1"),
                        temp=(bits[2] == "1"),
                        sun=(bits[3] == "1"),
                        snow=(bits[4] == "1"))

        def parse_station(dom):
            stitle = dom.get("title").replace("：", ":")
            title = dict(filter(lambda y: len(y) == 2,
                                map(lambda x: x.split(":"), stitle.split("\n"))))

            name = title["地点名"]
            stid = dom.cssselect("input[name=stid]")[0].value
            stname = dom.cssselect("input[name=stname]")[0].value
            kansoku = kansoku_items(dom.cssselect("input[name=kansoku]")[0].value)
            assert name == stname
            return stname, dict(id=stid, flags=kansoku)

        def parse_prefs(dom):
            name = dom.text
            prid = int(dom.cssselect("input[name=prid]")[0].value)
            return name, prid

        URL = "http://www.data.jma.go.jp/gmd/risk/obsdl/top/station"
        data = encode_data({"pd": "%02d" % pd})
        xml = urllib.request.urlopen(URL, data=data).read().decode("utf-8")
        tree = lxml.html.fromstring(xml)

        if pd > 0:
            station_ids = dict(map(parse_station, tree.cssselect("div.station")))      # 地区名とstation IDの対応辞書を作成
        else:
            station_ids = dict(map(parse_prefs, tree.cssselect("div.prefecture")))      # 県名とprefecture IDの対応辞書を作成
        return station_ids

    station0 = get_station_by_id(0)         # station IDを0にして全国単位で検索
    station1 = station0[prefecture]         # station1には、指定した県のprefecture IDが入る
    station2 = get_station_by_id(station1)  # station IDをprefecture IDにして、指定した県単位で検索
    station = station2[station]["id"]       # 取得したstation0 IDリストから、該当する地区のstation IDを得る

    return station


def download_temperature_csv(phpsessid, station, element, start_date, end_date):
    """
    気温データをcsv形式でダウンロードする
    :param phpsessid: セッションID
    :param station: 地点 ID
    :param element: 取得データID
    :param start_date: 開始日付
    :param end_date: 終了日付
    :return: 取得したCSVデータ
    """
    params = {              #  サイトに送るForm Data
        "PHPSESSID": phpsessid,
        # 共通フラグ
        "rmkFlag": 1,  # 利用上注意が必要なデータを格納する
        "disconnectFlag": 1,  # 観測環境の変化にかかわらずデータを格納する
        "csvFlag": 1,  # すべて数値で格納する
        "ymdLiteral": 1,  # 日付は日付リテラルで格納する
        "youbiFlag": 0,  # 日付に曜日を表示する
        "kijiFlag": 0,  # 最高・最低（最大・最小）値の発生時刻を表示
        # 日別値データ選択
        "aggrgPeriod": 1,  # 日別値
        "stationNumList": '["%s"]' % station,  # 観測地点IDのリスト
        "elementNumList": '[["%s",""]]' % element,  # 項目IDのリスト
        "ymdList": '["%d", "%d", "%d", "%d", "%d", "%d"]' % (
            start_date.year, end_date.year,
            start_date.month, end_date.month,
            start_date.day, end_date.day),  # 取得する期間
        "jikantaiFlag": 0,  # 特定の時間帯のみ表示する
        "interAnnualFlag": 1,  # 連続した期間で表示する
        'jikantaiList': [1, 24],
        "optionNumList": ' [["op1", 0]]',  # 平年値も得る [["op1",0]]
        "downloadFlag": "true",  # CSV としてダウンロードする？
        "huukouFlag": 0,
    }

    print("load")
    URL = "http://www.data.jma.go.jp/gmd/risk/obsdl/show/table"
    data = encode_data(params)
    csv = urllib.request.urlopen(URL, data=data).read().decode("cp932", "ignore")
    # なぜか、文字コードの変換エラーが出る事があるが、どうせタイトル行は削るので、無視する
    #    csv = urllib.request.urlopen(URL, data=data).read().decode("shift-jis")
    print("complete")
    return csv


class MeteorologicalAgency:
    """
    気象庁から平均気温と平年気温を取得するクラス
    """

    def __init__(self, prefecture="高知", station="窪川"):
        """
        取得する地点を指定する
        :param prefecture:気象庁に通知する県名
        :param station:気象庁に通知する地区名
        """
        self.prefecture = prefecture
        self.station = station
        self.csv = None

    def get_temperature_string(self, start_date, end_data):
        """
        指定した期間の平均気温と平年値のリストをCSV形式で得る
        :param start_date: 期間の開始日
        :param end_data: 期間の終了日
        :return: 気象庁から送られてきた気温データの文字列（CSV)
        """

        station_id = get_station(self.prefecture, self.station)
        phpsess_id = get_phpsessid()
        self.csv = download_temperature_csv(phpsess_id, station_id, 201,
                                            start_date, end_data)
        return self.csv

    def get_temperature_list(self, start_date, end_data):
        """
        指定した期間の平均気温と平年気温のリストを得る
        :param start_date: 期間の開始日
        :param end_data: 期間の終了日
        :return: [[日付(date型),平均気温(float),平年気温(float)]...]のリスト
        """
        self.get_temperature_string(start_date, end_data)
        out_list = []
        for item in self.csv.splitlines(True)[6:]:  # 最初の6行はタイトルなので無視する
            line_data = item.split(',')  # カンマで分割する
            data_pair = [dt.strptime(line_data[0], '%Y/%m/%d'), float(line_data[1]), float(line_data[4])]
            # 日付型と浮動小数に変換する
            out_list.append(data_pair)
        return out_list


if __name__ == "__main__":
    obj = MeteorologicalAgency("高知", "窪川")
    temperature_list = obj.get_temperature_list(date(2019, 12, 26), date(2020, 12, 25))

    print(temperature_list)
