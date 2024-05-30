import requests,re,time
from datetime import datetime,timedelta
from collections import defaultdict
import pandas as pd
import loadDataInExcel, loadGoogleSheet
import cosmo, setsuko, nutra
import os
from dotenv import load_dotenv

load_dotenv()

def parsing():
    # url на аналитику
    url = 'https://seller-analytics-api.wildberries.ru/api/v2/nm-report/detail'
    # url на получение id
    url_1 = 'https://advert-api.wb.ru/adv/v1/promotion/count'
    # url на получение статистики
    url_2 = 'https://advert-api.wb.ru/adv/v2/fullstats'

    API_KEY1 = os.getenv("API_KEY1")
    API_KEY11 = os.getenv("API_KEY11")
    API_KEY2 = os.getenv("API_KEY2")
    API_KEY22 = os.getenv("API_KEY22")
    API_KEY3 = os.getenv("API_KEY3")
    API_KEY33 = os.getenv("API_KEY33")
    NUTRA = os.getenv('SECRET_JSON')

    HeaderApiKey1 = {
        'Authorization': f'{API_KEY1}',
        'Content-Type': 'application/json'
    }

    HeaderApiKey11 = {
        'Authorization': f'{API_KEY11}',
        'Content-Type': 'application/json'
    }

    HeaderApiKey2 = {
        'Authorization': f'{API_KEY2}',
        'Content-Type': 'application/json'
    }

    HeaderApiKey22 = {
        'Authorization': f'{API_KEY22}',
        'Content-Type': 'application/json'
    }

    HeaderApiKey3 = {
        'Authorization': f'{API_KEY3}',
        'Content-Type': 'application/json'
    }

    HeaderApiKey33 = {
        'Authorization': f'{API_KEY33}',
        'Content-Type': 'application/json'
    }

    now = datetime.now()
    now = now - timedelta(days=1)
    start_of_month = datetime(now.year, now.month, 1)
    dates = pd.date_range(start_of_month, now, freq='D')

    newdates = dates
    for date in newdates:
        next_day = False
        Jdata = None
        Jdata1 = None

        Jdata00 = None
        Jdata11 = None

        Jdata000 = None
        Jdata111 = None
        while next_day == False:
            data = {
                "brandNames": [],
                "timezone": "Europe/Moscow",
                "period": {
                    "begin": date.strftime("%Y-%m-%d %H:%M:%S"),
                    "end": (date + timedelta(days=1)).strftime("%Y-%m-%d %H:%M:%S")
                },
                "orderBy": {
                    "field": "ordersSumRub",
                    "mode": "desc"
                },
                "page": 1
            }
            time.sleep(1)
            print(date)
            print()
            response = requests.post(url, json=data, headers=HeaderApiKey1)
            response1 = requests.get(url_1, headers=HeaderApiKey11)
            print("н")

            response00 = requests.post(url, json=data, headers=HeaderApiKey2)
            response11 = requests.get(url_1, headers=HeaderApiKey22)

            response000 = requests.post(url, json=data, headers=HeaderApiKey3)
            response111 = requests.get(url_1, headers=HeaderApiKey33)


            if response.status_code == 200 and response00.status_code == 200 and response000.status_code == 200:
                print('Данные успешно получены за', date.strftime("%Y-%m-%d"))
                print("-----")
                Jdata = response.json()
                Jdata1 = response1.json()

                Jdata00 = response00.json()
                Jdata11 = response11.json()

                Jdata000 = response000.json()
                Jdata111 = response111.json()

                next_day = True
            else:
                if response.status_code != 200 and response00.status_code != 200 and response000.status_code != 200:
                    print(response.status_code)
                    Jdata = Jdata
                    Jdata1 = Jdata1

                    Jdata00 = Jdata00
                    Jdata11 = Jdata11

                    Jdata000 = Jdata000
                    Jdata111 = Jdata111
                    next_day = False
                    continue

            def deleate(res):
                while re.search('previousPeriod.*?stocks', str(res), flags=re.DOTALL):
                    res = re.sub('previousPeriod.*?stocks', '', str(res), flags=re.DOTALL)
                return res

            def Remove(arr):
                cleaned_arr = []
                for word in arr:
                    word = ''.join(e for e in word if e.isalnum() or e.isspace())
                    if word:
                        cleaned_arr.append(word)
                return cleaned_arr

            def SpaceX(cleaned_arr):
                return ''.join(cleaned_arr)

            pop = deleate(Jdata)
            pop1 = deleate(Jdata1)

            pop00 = deleate(Jdata00)
            pop11 = deleate(Jdata11)

            pop000 = deleate(Jdata000)
            pop111 = deleate(Jdata111)

            cleaned_text = Remove(str(pop))
            cleaned_text1 = Remove(str(pop1))

            cleaned_text00 = Remove(str(pop00))
            cleaned_text11 = Remove(str(pop11))

            cleaned_text000 = Remove(str(pop000))
            cleaned_text111 = Remove(str(pop111))

            result = SpaceX(cleaned_text)
            result1 = SpaceX(cleaned_text1)

            result00 = SpaceX(cleaned_text00)
            result11 = SpaceX(cleaned_text11)

            result000 = SpaceX(cleaned_text000)
            result111 = SpaceX(cleaned_text111)

            words = result.split()
            words1 = result1.split()

            words00 = result00.split()
            words11 = result11.split()

            words000 = result000.split()
            words111 = result111.split()

            words += words00 + words000

            indices_id = [i for i, x in enumerate(words1) if x == "advertId"]
            id = [int(words1[i + 1]) if i + 1 < len(words1) else None for i in indices_id]

            indices_id1 = [i for i, x in enumerate(words11) if x == "advertId"]
            id1 = [int(words11[i + 1]) if i + 1 < len(words11) else None for i in indices_id1]

            indices_id11 = [i for i, x in enumerate(words111) if x == "advertId"]
            id11 = [int(words111[i + 1]) if i + 1 < len(words111) else None for i in indices_id11]

            date_from = date.strftime("%Y-%m-%d")
            next_day2 = False

            while next_day2 == False:
                params1 = [{'id': c, 'dates': [date_from]} for c in id]
                params11 = [{'id': c, 'dates': [date_from]} for c in id1]
                params111 = [{'id': c, 'dates': [date_from]} for c in id11]
                response2 = requests.post(url_2, headers=HeaderApiKey1, json=params1)
                response22 = requests.post(url_2, headers=HeaderApiKey2, json=params11)
                response222 = requests.post(url_2, headers=HeaderApiKey3, json=params111)

                if response2.status_code == 200 and response22.status_code == 200 and response222.status_code ==200:

                    next_day2 = True
                else:
                    if response2.status_code != 200 and response22.status_code != 200 and response222.status_code !=200:
                        print('Ожидание отклика сервера...')
                        time.sleep(20)
                        next_day2 = False
                        continue

            Jdata2 = response2.json()

            Jdata22 = response22.json()

            Jdata222 = response222.json()

            camp_data = []
            camp_data2 = []
            camp_data22 = []

            for c in Jdata2:
                for d in c['days']:
                    for a in d['apps']:
                        for nm in a['nm']:
                            nm['appType'] = a['appType']
                            nm['date'] = d['date']
                            nm['advertId'] = c['advertId']
                            camp_data.append(nm)

            for c in Jdata22:
                for d in c['days']:
                    for a in d['apps']:
                        for nm in a['nm']:
                            nm['appType'] = a['appType']
                            nm['date'] = d['date']
                            nm['advertId'] = c['advertId']
                            camp_data2.append(nm)

            if not isinstance(Jdata222, list):
                Jdata222_list = [Jdata222]
            else:
                Jdata222_list = Jdata222

            for c in Jdata222_list:
                for d in c['days']:
                    for a in d['apps']:
                        for nm in a['nm']:
                            nm['appType'] = a['appType']
                            nm['date'] = d['date']
                            nm['advertId'] = c['advertId']
                            camp_data22.append(nm)

            camp_df = pd.DataFrame(camp_data)
            camp_df2 = pd.DataFrame(camp_data2)
            camp_df3 = pd.DataFrame(camp_data22)

            df_filtered = camp_df.groupby('advertId').agg(
                {'nmId': 'first', 'views': 'sum', 'clicks': 'sum'}).reset_index()

            df_filtered = df_filtered.groupby('nmId').agg(
                lambda x: x.sum() if x.name != 'advertId' else x.iloc[
                    0]).reset_index()

            df_filtered.drop(columns=['advertId'], inplace=True)
            df_filtered['CTR'] = (round(df_filtered['clicks'] / df_filtered['views'] * 100, 2))
            camp_data1 = df_filtered.set_index('nmId').to_dict(orient="index")

            for k, v in camp_data1.items():
                camp_data1[k]['Показы'] = v.pop('views')
                camp_data1[k]['Клики'] = v.pop('clicks')
                camp_data1[k]['CTR'] = v.pop('CTR')

            df_filtered1 = camp_df2.groupby('advertId').agg(
                {'nmId': 'first', 'views': 'sum', 'clicks': 'sum'}).reset_index()

            df_filtered1 = df_filtered1.groupby('nmId').agg(
                lambda x: x.sum() if x.name != 'advertId' else x.iloc[
                    0]).reset_index()

            df_filtered1.drop(columns=['advertId'], inplace=True)
            df_filtered1['CTR'] = (round(df_filtered1['clicks'] / df_filtered1['views'] * 100, 2))
            camp_data2 = df_filtered1.set_index('nmId').to_dict(orient="index")

            for k, v in camp_data2.items():
                camp_data2[k]['Показы'] = v.pop('views')
                camp_data2[k]['Клики'] = v.pop('clicks')
                camp_data2[k]['CTR'] = v.pop('CTR')

            df_filtered11 = camp_df3.groupby('advertId').agg(
                {'nmId': 'first', 'views': 'sum', 'clicks': 'sum'}).reset_index()

            df_filtered11 = df_filtered11.groupby('nmId').agg(
                lambda x: x.sum() if x.name != 'advertId' else x.iloc[
                    0]).reset_index()

            df_filtered11.drop(columns=['advertId'], inplace=True)
            df_filtered11['CTR'] = (round(df_filtered11['clicks'] / df_filtered11['views'] * 100, 2))
            camp_data3 = df_filtered11.set_index('nmId').to_dict(orient="index")

            for k, v in camp_data3.items():
                camp_data3[k]['Показы'] = v.pop('views')
                camp_data3[k]['Клики'] = v.pop('clicks')
                camp_data3[k]['CTR'] = v.pop('CTR')

            found_brand = False
            buffer = []
            brands = []

            for word in words:
                if word == "brandName":
                    found_brand = True
                    if buffer:
                        brand_name = ' '.join(buffer)
                        brand_name = brand_name.replace("brandName", "")
                        brands.append(brand_name)
                        buffer = []
                elif word == "object":
                    found_brand = False

                if found_brand:
                    buffer.append(word)

            if buffer:
                brand_name = ' '.join(buffer)
                brand_name = brand_name.replace("brandName", "")
                brands.append(brand_name)

            indices_name = [i for i, x in enumerate(words) if x == "name"]
            name = [words[i + 1] if i + 1 < len(words) else None for i in indices_name]
            indices_nmID = [i for i, x in enumerate(words) if x == "nmID"]
            nmID = [words[i + 1] if i + 1 < len(words) else None for i in indices_nmID]
            indices_ost = [i for i, x in enumerate(words) if x == "stocksWb"]
            stocksWb = [words[i + 1] if i + 1 < len(words) else None for i in indices_ost]
            indices_o = [i for i, x in enumerate(words) if x == "openCardCount"]
            openCardCount = [words[i + 1] if i + 1 < len(words) else None for i in indices_o]
            indices_a = [i for i, x in enumerate(words) if x == "addToCartPercent"]
            addToCartPercent = [words[i + 1] if i + 1 < len(words) else None for i in indices_a]
            indices_c = [i for i, x in enumerate(words) if x == "cartToOrderPercent"]
            cartToOrderPercent = [words[i + 1] if i + 1 < len(words) else None for i in indices_c]
            indices_aa = [i for i, x in enumerate(words) if x == "addToCartCount"]
            addToCartCount = [words[i + 1] if i + 1 < len(words) else None for i in indices_aa]
            combined_list = []

            max_len = max(len(brands), len(openCardCount), len(addToCartPercent), len(cartToOrderPercent),
                          len(addToCartCount), len(stocksWb), len(nmID), len(name))

            for i in range(max_len):
                if i < len(name):
                    combined_list.append(name[i])
                if i < len(brands):
                    combined_list.append("brand: " + brands[i])
                if i < len(openCardCount):
                    combined_list.append("Переходы: " + openCardCount[i])
                if i < len(addToCartPercent):
                    combined_list.append("Конверсии в корзину: " + addToCartPercent[i])
                if i < len(cartToOrderPercent):
                    combined_list.append("Конверсии в заказ: " + cartToOrderPercent[i])
                if i < len(addToCartCount):
                    combined_list.append("Добавление в корзину: " + addToCartCount[i])
                if i < len(stocksWb):
                    combined_list.append("Остатки товаров на складе: " + stocksWb[i])
                if i < len(nmID):
                    combined_list.append("ID: " + nmID[i])

            brand_data = defaultdict(
                lambda: {'Переходы': 0, 'Конверсии в корзину': 0, 'Конверсии в заказ': 0, 'Добавление в корзину': 0,
                         'Остатки товаров на складе': 0, 'Бренд': "", 'Показы': '-', 'Клики': "-", 'CTR': '-'})
            IDD = []

            idol = 0
            while idol < len(combined_list):
                if combined_list[idol].startswith('Возбуждающие'):
                    del combined_list[idol:idol + 8]
                elif combined_list[idol].startswith('Лубриканты'):
                    del combined_list[idol:idol + 8]
                else:
                    idol += 1

            index = 0
            while index < len(combined_list):
                del combined_list[index]
                index += 7

            index_to_remove = []

            for i in range(len(combined_list)):
                if 'brand:  burner fat' in combined_list[i]:
                    for j in range(i, i + 7):
                        index_to_remove.append(j)
                if 'brand:  SUCCUBA' in combined_list[i]:
                    for j in range(i, i + 7):
                        index_to_remove.append(j)
                if 'brand:  Lovelei' in combined_list[i]:
                    for j in range(i, i + 7):
                        index_to_remove.append(j)
                if 'brand:  RECIPE of Love' in combined_list[i]:
                    for j in range(i, i + 7):
                        index_to_remove.append(j)

            filtered_data666 = [combined_list[i] for i in range(len(combined_list)) if i not in index_to_remove]

            for i in range(0, len(filtered_data666), 7):
                brand = filtered_data666[i].split(': ')[1]
                openCardCount = int(filtered_data666[i + 1].split(': ')[1])
                addToCartPercent = int(filtered_data666[i + 2].split(': ')[1])
                cartToOrderPercent = int(filtered_data666[i + 3].split(': ')[1])
                addToCartCount = int(filtered_data666[i + 4].split(': ')[1])
                stocksWb = int(filtered_data666[i + 5].split(': ')[1])
                nmID1 = int(filtered_data666[i + 6].split(': ')[1])

                brand_data[nmID1]['Переходы'] = openCardCount
                brand_data[nmID1]['Конверсии в корзину'] = addToCartPercent
                brand_data[nmID1]['Конверсии в заказ'] = cartToOrderPercent
                brand_data[nmID1]['Добавление в корзину'] = addToCartCount
                brand_data[nmID1]['Остатки товаров на складе'] = stocksWb
                brand_data[nmID1]['Бренд'] = brand
                brand_data[nmID1]['ID'] = nmID1
                IDD.append(nmID1)

            for key in camp_data1.keys():
                if key in brand_data.keys():
                    brand_data[key].update(camp_data1[key])
                    brand_data[key]['Показы'] = camp_data1[key]['Показы']
                    brand_data[key]['Клики'] = camp_data1[key]['Клики']
                    brand_data[key]['CTR'] = camp_data1[key]['CTR']

            for key in camp_data2.keys():
                if key in brand_data.keys():
                    brand_data[key].update(camp_data2[key])
                    brand_data[key]['Показы'] = camp_data2[key]['Показы']
                    brand_data[key]['Клики'] = camp_data2[key]['Клики']
                    brand_data[key]['CTR'] = camp_data2[key]['CTR']

            for key in camp_data3.keys():
                if key in brand_data.keys():
                    brand_data[key].update(camp_data3[key])
                    brand_data[key]['Показы'] = camp_data3[key]['Показы']
                    brand_data[key]['Клики'] = camp_data3[key]['Клики']
                    brand_data[key]['CTR'] = camp_data3[key]['CTR']

            keys_list = dict(sorted(brand_data.items(), key=lambda x: x[1]['Бренд']))

            loadDataInExcel.Data(keys_list, brand_data, dates)
            if response.status_code == 200:
                loadDataInExcel.columnStat += 1

            if date.strftime("%Y-%m-%d") == now.strftime("%Y-%m-%d"):
                loadGoogleSheet.CopyFromExcInGsh()
                cosmo.rooo()
                setsuko.rooo()
                nutra.rooo()

parsing()
