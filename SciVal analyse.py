import pandas as pd
import numpy as np
import os
import re
import pickle
from collections import Counter
from functools import reduce
import sys


def print_count(_list, num=10):
    statistical = Counter(_list)
    for j in statistical.most_common(num):
        print(str(list(j)[0]) + '	' + str(list(j)[1]) + '	' + str(round((list(j)[1] / len(_list)) * 100, 4)) + "%")


def save_count(_list, path, name='name', num=10):
    # 输出列表频次统计结果
    # for i in _list:
    #     print(i)
    statistical = Counter(_list)
    _all = []
    for j in statistical.most_common(num):
        _all.append({
            name: str(list(j)[0]),
            'count': int(list(j)[1]),
            'proportion': str(round((list(j)[1] / len(_list)) * 100, 4)) + "%"
        })
    pd.DataFrame(_all).to_excel(path)


def list_strip(_list):
    if _list[0] != 0 and _list[-1] != 0:
        return _list
    elif _list[0] == 0 and _list[-1] != 0:
        _start = 0
        for idx, num in enumerate(_list):
            if num == 0:
                _start = idx + 1
            else:
                break
        return _list[_start:]
    elif _list[0] != 0 and _list[-1] == 0:
        _end = 0
        for idx, num in enumerate(_list[::-1]):
            if num == 0:
                _end = idx + 1
            else:
                break
        return _list[:-_end]
    else:
        _start = 0
        _end = 0
        for idx, num in enumerate(_list):
            if num == 0:
                _start = idx + 1
            else:
                break
        for idx, num in enumerate(_list[::-1]):
            if num == 0:
                _end = idx + 1
            else:
                break

    return _list[_start:-_end]


def compute_CAGR(_list):
    _list = list_strip(_list)
    if len(_list) <= 2:
        CAGR = '-'
    else:
        CAGR = (_list[-1] / _list[0]) ** (1 / (len(_list)-1)) - 1
        # CAGR = str(round(CAGR * 100, 4)) + "%"

    return CAGR


def journal_export():
    journal_data = pd.read_excel('./journal/全部期刊数据.xlsx')
    journal_map = {}
    for i in range(journal_data.shape[0]):
        IF_text = str(journal_data.at[i, 'Journal Impact Factor'])
        if IF_text != 'Not Available':
            journal_map[str(journal_data.at[i, 'Full Journal Title']).lower()] = float(IF_text)

    with open('./runtime data/journal_IF_map', 'wb') as f:
        pickle.dump(journal_map, f)
        f.close()


def save_data():
    """
    all_data = {
        '2011': [page_data1, page_data2, ....],
        '2012': [page_data1, page_data2, ....],
        ....
        '2020': [page_data1, page_data2, ....]
    }
    """
    all_data = {}
    dir_path = './data/'
    file_list = [file for file in os.listdir(dir_path) if '.xlsx' in file and '~$' not in file]
    data_list = []
    title_set = set()

    for file in file_list:
        data = pd.read_excel(f'{dir_path}{file}')
        data_list.append((file[:-5], data))
        for title in data.columns:
            title_set.add(title)

    title_list = list(title_set)

    for file, data in data_list:
        all_data[file] = []
        for i in range(data.shape[0]):
            cur = {}
            for title in title_list:
                cur[title] = data.at[i, title]
            all_data[file].append(cur)

    with open('./runtime data/all_data', 'wb') as f:
        pickle.dump(all_data, f)
        f.close()


def institutions_papers_export():
    """
    all_data = {
        'institution_1': [page_data1, page_data2, ....],
        'institution_2': [page_data1, page_data2, ....],
        ....
        'institution_n': [page_data1, page_data2, ....]
    }
    """
    all_data = get_data("all_data")
    institutions_text_and_data_list = reduce(
        lambda _all, item_list: _all + [(item['Institutions'], item) for item in item_list],
        all_data.values(), [])

    institutions_and_data_list = reduce(
        lambda _all, cur_data: _all + [(institution.strip(), cur_data[1]) for institution in cur_data[0].split('|')],
        institutions_text_and_data_list, [])

    institutions_papers = {}
    institutions_list = list(set([institution[0] for institution in institutions_and_data_list]))

    for institution in institutions_list:
        institutions_papers[institution] = []

    for institution, paper in institutions_and_data_list:
        institutions_papers[institution].append(paper)

    with open('./runtime data/institutions_papers', 'wb') as f:
        pickle.dump(institutions_papers, f)
        f.close()


def sub_institutions_papers_export():
    data = get_data('institutions_papers')
    use_institutions = [
        'University of Florida', 'Universidade de São Paulo', 'University of California at Davis',
        'Universidade Federal de Viçosa', 'Wageningen University & Research', 'Cornell University',
        'Nanjing Agricultural University', 'Northwest Agriculture and Forestry University',
        'China Agricultural University', 'Zhejiang University', 'South China Agricultural University',
        'Huazhong Agricultural University']
    use_institutions_papers = {}
    for institution in use_institutions:
        if institution == 'University of California at Davis':
            use_institutions_papers[institution] = data[institution] + data['University of California Office of the President']
        else:
            use_institutions_papers[institution] = data[institution]
    with open('./runtime data/sub_institutions_papers', 'wb') as f:
        pickle.dump(use_institutions_papers, f)
        f.close()


def country_papers_export():
    """
    all_data = {
        'country_1': [page_data1, page_data2, ....],
        'country_2': [page_data1, page_data2, ....],
        ....
        'country_n': [page_data1, page_data2, ....]
    }
    """
    all_data = get_data("all_data")
    country_text_and_data_list = reduce(
        lambda _all, item_list: _all + [(item['Country/Region'], item) for item in item_list],
        all_data.values(), [])

    country_text_and_data_list = reduce(
        lambda _all, cur_data: _all + [(country.strip(), cur_data[1]) for country in cur_data[0].split('|')],
        country_text_and_data_list, [])

    country_papers = {}
    country_list = list(set([country[0] for country in country_text_and_data_list]))

    for institution in country_list:
        country_papers[institution] = []

    for institution, paper in country_text_and_data_list:
        country_papers[institution].append(paper)

    with open('./runtime data/country_papers', 'wb') as f:
        pickle.dump(country_papers, f)
        f.close()


def sub_country_papers_export():
    use_country = ['United States', 'China', 'Brazil', 'India', 'United Kingdom', 'Australia', 'Germany', 'Japan',
     'France', 'Spain', 'Canada', 'Italy', 'South Korea', 'Iran', 'Netherlands', 'South Africa', 'Poland', 'Mexico',
     'Pakistan', 'New Zealand', 'Egypt', 'Argentina', 'Switzerland', 'Turkey', 'Russian Federation', 'Belgium',
     'Czech Republic', 'Thailand', 'Sweden']

    use_country_papers = {}

    country_papers = get_data('country_papers')
    for country, paper_list in country_papers.items():
        if country in use_country:
            use_country_papers[country] = paper_list

    with open('./runtime data/sub_country_papers', 'wb') as f:
        pickle.dump(use_country_papers, f)
        f.close()


def topic_cluster_papers_export():
    """
        all_data = {
            topic_cluster_1': [page_data1, page_data2, ....],
            topic_cluster_2': [page_data1, page_data2, ....],
            ....
            topic_cluster_n': [page_data1, page_data2, ....]
        }
        """
    all_data = get_data("all_data")
    topic_cluster_and_data_list = reduce(
        lambda _all, item_list: _all + [(item['Topic Cluster name'], item) for item in item_list],
        all_data.values(), [])

    topic_cluster_papers = {}
    topic_cluster_list = list(set([institution[0] for institution in topic_cluster_and_data_list]))

    for institution in topic_cluster_list:
        topic_cluster_papers[institution] = []

    for institution, paper in topic_cluster_and_data_list:
        topic_cluster_papers[institution].append(paper)

    with open('./runtime data/topic_cluster_papers', 'wb') as f:
        pickle.dump(topic_cluster_papers, f)
        f.close()


def sub_topic_cluster_papers_export():
    use_topic = [
        'Bacillus Thuringiensis,Aphidoidea,Aphids',
        'Viruses,Mosaic Viruses,Phytoplasma',
        'Phytophthora,Trichoderma,Phytophthora Infestans',
        'Weeds,Herbicides,Weed Control',
        'Arabidopsis,Plants,Genes',
        'Nematoda,Root-Knot Nematodes,Meloidogyne Incognita',
        'China,Fungi,Leaves',
        'Xanthomonas,Ralstonia Solanacearum,Genome',
        'Baculoviridae,Entomopathogenic Nematodes,Nucleopolyhedrovirus',
        'Wheat,Triticum,Triticum Aestivum',
        'Plants,Rhizosphere,Rhizobium',
        'Mycotoxins,Aflatoxins,Ochratoxins',
        'Formicidae,Hymenoptera,Ant',
        'Powdery Mildew,Sugar Beet,Fungicides',
        'Hymenoptera,Galls,Braconidae',
        'Phosphines,Tribolium Castaneum,Coleoptera',
        'Dengue,Viruses,Dengue Virus',
        'Ticks,Lyme Disease,Borrelia Burgdorferi',
        'Fungi,Magnaporthe,Oryza Sativa',
        'Mites,Acari,Oribatida',
        'Beetle,Bark Beetles,Curculionidae',
        'Tephritidae,Fruit Flies,Diptera',
        'Salmonella,Escherichia Coli,Listeria Monocytogenes'
    ]
    data = get_data('topic_cluster_papers')
    use_topic_papers = {}

    for topic in use_topic:
        use_topic_papers[topic] = data[topic]

    with open('./runtime data/sub_topic_cluster_papers', 'wb') as f:
        pickle.dump(use_topic_papers, f)
        f.close()


def sub_institutions_papers_doi_file_export():
    data = get_data('sub_institutions_papers')
    for ins, paper_list in data.items():
        # os.mkdir(f'./wos重点大学/wos paper/{ins}')
        # pd.DataFrame([paper['DOI'] for paper in paper_list if '10' in paper['DOI']]).to_excel(f'./wos重点大学/{ins}.xlsx')
        _file = open(f'./wos重点大学/doi/{ins}', 'w', encoding='utf-8')

        doi_list = [f"(DO={paper['DOI']})" for paper in paper_list if '10' in paper['DOI']]
        _file.write(' OR \n'.join(doi_list))
        _file.close()


def get_data(file_name):
    with open(f'./runtime data/{file_name}', 'rb') as f:
        return pickle.load(f)


def journals_statistics():
    all_data = get_data("all_data")
    journals_list = reduce(
        lambda all_journals, item_list: all_journals + [item['Scopus Source title'] for item in item_list],
        all_data.values(), [])

    print(f"共 {len(journals_list)} 篇文献")
    # print_count(journals_list, 20)

    save_count(journals_list, './中间计算数据/期刊统计.xlsx', name='journal', num=sys.maxsize)


def FWCI_and_TC_statistics(data, _type, file_name):
    """
    求：发文量:总和,占比,每年|被引频次:平均,每年|FWCI:区间(0,5,10),占比,每年,前后五年均值,增长率
    :param data: 传入runtime数据
    :param _type: 统计类型，机构/国家/年份/topic cluster等
    :param file_name: 输出文件名
    :return:
    """
    _list = []
    # 依据EID求去重后总文章数量
    all_len = len(set([paper['EID'] for paper in reduce(lambda _list, _all: _all + _list, data.values(), [])]))

    year_list1 = ['2011', '2012', '2013', '2014', '2015']
    year_list2 = ['2016', '2017', '2018', '2019', '2020']
    all_year_list = ['2011', '2012', '2013', '2014', '2015', '2016', '2017', '2018', '2019', '2020']

    for title, paper_list in data.items():
        print(title)
        year_2011_2015 = [paper for paper in paper_list if str(paper['Year']) in year_list1]
        year_2016_2020 = [paper for paper in paper_list if str(paper['Year']) in year_list2]

        all_year_count = [1/sys.maxsize] * 10
        all_year_FWCI_sum = [0] * 10
        all_year_FWCI_average = []
        all_year_TC_sum = [0] * 10
        all_year_TC_average = []
        all_year_FWCI_growth_rate = []
        average_growth_rate = 0

        FWCI_interval = [0, 0]      # [>10, 5-10]

        for paper in paper_list:
            FWCI = float(paper['Field-Weighted Citation Impact'])
            TC = int(paper['Citations'])
            # 按照年份按位求每年的FWCI和，同时判断FWCI区间
            all_year_count[all_year_list.index(str(paper['Year']))] += 1
            all_year_FWCI_sum[all_year_list.index(str(paper['Year']))] += FWCI
            all_year_TC_sum[all_year_list.index(str(paper['Year']))] += TC

            if FWCI >= 10:
                FWCI_interval[0] += 1
            elif 5 <= FWCI < 10:
                FWCI_interval[1] += 1

        for i in range(len(all_year_FWCI_sum)):
            if all_year_count[i] < 1:
                all_year_FWCI_average.append(0)
                all_year_TC_average.append(0)
            else:
                all_year_FWCI_average.append(round(all_year_FWCI_sum[i] / all_year_count[i], 4))
                all_year_TC_average.append(round(all_year_TC_sum[i] / all_year_count[i], 4))

        have_paper_year_count = 0
        for i in range(len(all_year_FWCI_average) - 1):
            # 上值（i）为0，则跳过当前，当前值为"-"
            if all_year_FWCI_average[i] != 0:
                have_paper_year_count += 1
                growth_rate = (all_year_FWCI_average[i+1] - all_year_FWCI_average[i]) / all_year_FWCI_average[i]
                growth_rate = round(growth_rate, 4)
                all_year_FWCI_growth_rate.append(growth_rate)
                average_growth_rate += growth_rate
            else:
                all_year_FWCI_growth_rate.append('-')

        # 平均增长率
        if have_paper_year_count != 0:
            average_growth_rate /= have_paper_year_count
        else:
            average_growth_rate = '-'

        # 年均复合增长率
        for i in range(len(all_year_count)):
            if all_year_count[i] < 1:
                all_year_count[i] = 0

        FWCI_CAGR = compute_CAGR(all_year_FWCI_average)
        paper_count_CAGR = compute_CAGR(all_year_count)
        paper_count_CAGR_2011_2015 = compute_CAGR(all_year_count[:5])
        paper_count_CAGR_2016_2020 = compute_CAGR(all_year_count[5:])

        FWCI_list = [float(paper['Field-Weighted Citation Impact']) for paper in paper_list]
        FWCI_average = round(reduce(lambda cur, _all: _all + cur, FWCI_list, 0) / len(paper_list), 4)
        TC_list = [int(paper['Citations']) for paper in paper_list]
        TC_average = round(reduce(lambda cur, _all: _all + cur, TC_list, 0) / len(paper_list), 4)

        _2011_2015_FWCI_list = [float(paper['Field-Weighted Citation Impact']) for paper in year_2011_2015]
        _2016_2020_FWCI_list = [float(paper['Field-Weighted Citation Impact']) for paper in year_2016_2020]

        if len(year_2011_2015) != 0:
            _2011_2015_FWCI_average = round(reduce(lambda cur, _all: _all + cur, _2011_2015_FWCI_list, 0) / len(_2011_2015_FWCI_list), 4)
        else:
            _2011_2015_FWCI_average = '-'

        if len(year_2016_2020) != 0:
            _2016_2020_FWCI_average = round(reduce(lambda cur, _all: _all + cur, _2016_2020_FWCI_list, 0) / len(_2016_2020_FWCI_list), 4)
        else:
            _2016_2020_FWCI_average = '-'

        _list.append({
            _type: title,
            '发文量': len(paper_list),
            '发文占比': len(paper_list) / all_len,
            '平均FWCI': FWCI_average,
            '平均被引频次': TC_average,
            '1': ' ',
            'FWCI>=10': FWCI_interval[0],
            '5<=FWCI<10': FWCI_interval[1],
            '2': ' ',
            'FWCI>=10占比': FWCI_interval[0] / len(paper_list),
            '5<=FWCI<10占比': FWCI_interval[1] / len(paper_list),
            '3': ' ',
            '2011-2015发文量': len(year_2011_2015),
            '2016-2020发文量': len(year_2016_2020),
            '4': ' ',
            '2011-2015发文占比': len(year_2011_2015) / len(paper_list),
            '2016-2020发文占比': len(year_2016_2020) / len(paper_list),
            '5': ' ',
            '2011-2015总平均FWCI': _2011_2015_FWCI_average,
            '2016-2020总平均FWCI': _2016_2020_FWCI_average,
            '6': ' ',
            '2011-2020每年发文量': '|'.join([str(int(year_count)) for year_count in all_year_count]),
            '2011-2020每年平均FWCI': '|'.join([str(FWCI_average) for FWCI_average in all_year_FWCI_average]),
            '2011-2020每年平均TC': '|'.join([str(TC_average) for TC_average in all_year_TC_average]),
            # '2012-2020每年FWCI年度增长率': '|'.join([str(round(year, 4)) for year in all_year_FWCI_growth_rate]),
            '2012-2020每年FWCI年度增长率': '|'.join([str(year) for year in all_year_FWCI_growth_rate]),
            '7': ' ',
            'FWCI平均增长率': average_growth_rate,
            '2012-2020年FWCI年均复合增长率': FWCI_CAGR,
            '2011-2015年发文量年均复合增长率': paper_count_CAGR_2011_2015,
            '2016-2020年发文量年均复合增长率': paper_count_CAGR_2016_2020,
            '2012-2020年发文量年均复合增长率': paper_count_CAGR
        })

    _list = sorted(_list, key=lambda x: x['发文量'], reverse=True)

    pd.DataFrame(_list).to_excel(f'./中间计算数据/{file_name}.xlsx')


def topic_cluster_highest_FWCI_country_statistics():
    data = get_data('topic_cluster_papers')
    use_country = ['United States', 'China', 'Brazil', 'India', 'United Kingdom', 'Australia', 'Germany', 'Japan',
                   'France', 'Spain', 'Canada', 'Italy', 'South Korea', 'Iran', 'Netherlands', 'South Africa', 'Poland',
                   'Mexico', 'Pakistan', 'New Zealand', 'Egypt', 'Argentina', 'Switzerland', 'Turkey',
                   'Russian Federation', 'Belgium', 'Czech Republic', 'Thailand', 'Sweden']

    _list = []

    all_len = len(reduce(lambda cur, _all: _all + cur, data.values(), []))

    for title, paper_list in data.items():
        # 初始化
        country_paper_map = {}
        country_paper_FWCI_sum = {}
        country_paper_FWCI_average = {}
        for country in use_country:
            country_paper_map[country] = []
            country_paper_FWCI_sum[country] = 0
            country_paper_FWCI_average[country] = 0

        FWCI_list = [float(paper['Field-Weighted Citation Impact']) for paper in paper_list]
        FWCI_average = round(reduce(lambda cur, _all: _all + cur, FWCI_list, 0) / len(paper_list), 4)

        # 获取当前主题下各有用国家论文集合
        country_text_paper_map_list = [(paper['Country/Region'], paper) for paper in paper_list]
        for country_text, paper in country_text_paper_map_list:
            for country in country_text.split('| '):
                if country in use_country:
                    country_paper_map[country].append(paper)

        # 计算各国家论文FWCI均值
        for country, country_paper_list in country_paper_map.items():
            for paper in country_paper_list:
                country_paper_FWCI_sum[country] += float(paper['Field-Weighted Citation Impact'])
            if len(country_paper_list) != 0:
                country_paper_FWCI_average[country] = country_paper_FWCI_sum[country] / len(country_paper_list)
            else:
                country_paper_FWCI_average[country] = -1

        country_paper_FWCI_average_map_sort = sorted(country_paper_FWCI_average.items(), key=lambda _country_map: _country_map[1], reverse=True)

        _list.append({
            'topic_cluster': title,
            '发文量': len(paper_list),
            '平均FWCI': FWCI_average,
            '发文占比': str(round((len(paper_list) / all_len) * 100, 4)) + "%",
            'FWCI最高国家': country_paper_FWCI_average_map_sort[0][0],
            'FWCI最高国家FWCI': round(country_paper_FWCI_average_map_sort[0][1], 4),
            'FWCI前三高国家': '|'.join([x[0] for x in country_paper_FWCI_average_map_sort[:3]]),
            'FWCI前三高国家FWCI': '|'.join([str(round(x[1], 4)) for x in country_paper_FWCI_average_map_sort[:3]])
        })

    _list = sorted(_list, key=lambda x: x['发文量'], reverse=True)

    pd.DataFrame(_list).to_excel('./中间计算数据/topic_cluster国家统计.xlsx')


def country_cooperate_analysis(data, title_name, exp_file_name):
    # 论文国际合作:数量,前后五年发文量,合作数,合作占比
    result_list = []

    for title, paper_list in data.items():
        year_list1 = ['2011', '2012', '2013', '2014', '2015']
        year_list2 = ['2016', '2017', '2018', '2019', '2020']

        year_2011_2015 = [paper for paper in paper_list if str(paper['Year']) in year_list1]
        year_2016_2020 = [paper for paper in paper_list if str(paper['Year']) in year_list2]

        total_cooperate_num = reduce(
            lambda _sum, paper: _sum + ('|' in paper['Country/Region']),
            paper_list, 0)

        _2011_2015_cooperate_num = reduce(
            lambda _sum, paper: _sum + ('|' in paper['Country/Region']),
            year_2011_2015, 0)

        _2016_2020_cooperate_num = reduce(
            lambda _sum, paper: _sum + ('|' in paper['Country/Region']),
            year_2016_2020, 0)
        # 主题合作有问题
        total_average = total_cooperate_num / len(paper_list)
        if len(year_2011_2015) == 0:
            year_2011_2015_average = '-'
            year_2016_2020_average = _2016_2020_cooperate_num / len(year_2016_2020)
        elif len(year_2016_2020) == 0:
            year_2011_2015_average = _2011_2015_cooperate_num / len(year_2011_2015)
            year_2016_2020_average = '-'
        else:
            year_2011_2015_average = _2011_2015_cooperate_num / len(year_2011_2015)
            year_2016_2020_average = _2016_2020_cooperate_num / len(year_2016_2020)

        print(title, total_average, year_2011_2015_average, year_2016_2020_average)
        result_list.append({
            title_name: title,
            '总发文量': len(paper_list),
            '总合作数': total_cooperate_num,
            '总合作占比': total_average,
            '2011-2015发文量': len(year_2011_2015),
            '2011-2015合作数': _2011_2015_cooperate_num,
            '2011-2015合作占比': year_2011_2015_average,
            '2016-2020发文量': len(year_2016_2020),
            '2016-2020合作数': _2016_2020_cooperate_num,
            '2016-2020合作占比': year_2016_2020_average
        })

        result_list = sorted(result_list, key=lambda x: x['总发文量'], reverse=True)

        pd.DataFrame(result_list).to_excel(f'./中间计算数据/{exp_file_name}.xlsx')


def ASJC_matrix(paper_list, exp_file_name):
    # 输入某一单元文献列表（list），输出ASJC交叉矩阵
    ASJC_list = [paper['All Science Journal Classification (ASJC) field name'] for paper in paper_list]
    kw_items_list = [ASJC_text.split('| ') for ASJC_text in ASJC_list]
    kw_map_list = []
    kw_count = {}

    for kw_co_occurrence in kw_items_list:
        for idx1, kw1 in enumerate(kw_co_occurrence):
            kw_count[kw1] = kw_count[kw1] + 1 if kw_count.get(kw1, 'none') != 'none' else 1
            for idx2 in range(idx1 + 1, len(kw_co_occurrence)):
                kw_map_list.append((kw1, kw_co_occurrence[idx2]))

    kw_count = sorted(kw_count.items(), key=lambda x: x[1], reverse=True)
    kw_list = [kw[0] for kw in kw_count]

    matrix = np.zeros((len(kw_list), len(kw_list)))
    for kw1, kw2 in kw_map_list:
        matrix[kw_list.index(kw1)][kw_list.index(kw2)] += 1
    matrix = matrix + matrix.T

    for idx, (kw, num) in enumerate(kw_count):
        matrix[idx][idx] = num

    matrix = matrix.astype(int)
    df = pd.DataFrame(matrix, index=kw_list, columns=kw_list)
    df.to_excel(f'./中间计算数据/topic cluster ASJC矩阵/{exp_file_name}.xlsx')


def ASJC_and_institution_cross(data, title_name, exp_file_name):
    # 统计ASJC跨学科性和机构交叉情况
    ASJC_TITLE_NAME = 'All Science Journal Classification (ASJC) code'
    INSTITUTION_TITLE_NAME = 'Institutions'
    res_list = []
    all_data_list = reduce(lambda _all, item_list: _all + item_list, data.values(), [])
    # all_len = len(set([paper['EID'] for paper in reduce(lambda _list, _all: _all + _list, data.values(), [])]))
    institutions_not_null_paper_list = [paper for paper in all_data_list if paper[INSTITUTION_TITLE_NAME] != '-']

    all_ASJC_cross_num = reduce(lambda _all, item: _all + len(re.findall('\| ', str(item[ASJC_TITLE_NAME]))) + 1,
                                all_data_list, 0)
    all_ASJC_cross_average = all_ASJC_cross_num / len(all_data_list)

    all_institution_cross_num = reduce(
        lambda _all, item: _all + len(re.findall('\| ', str(item[INSTITUTION_TITLE_NAME]))) + 1,
        institutions_not_null_paper_list, 0)
    all_institution_cross_average = all_institution_cross_num / len(institutions_not_null_paper_list)

    for title, paper_list in data.items():
        topic_all_ASJC_cross_num = 0
        topic_all_institution_cross_num = 0
        sub_institutions_not_null_paper_list = [paper for paper in paper_list if paper[INSTITUTION_TITLE_NAME] != '-']
        for paper in paper_list:
            topic_all_ASJC_cross_num += (len(
                re.findall('\| ', str(paper[ASJC_TITLE_NAME]))) + 1) / all_ASJC_cross_average
            if str(paper[INSTITUTION_TITLE_NAME]) != '-':
                topic_all_institution_cross_num += (len(
                    re.findall('\| ', str(paper[INSTITUTION_TITLE_NAME]))) + 1) / all_institution_cross_average

        ASJC_cross_average = topic_all_ASJC_cross_num / len(paper_list)
        if len(sub_institutions_not_null_paper_list):
            institution_cross_average = topic_all_institution_cross_num / len(sub_institutions_not_null_paper_list)
        else:
            institution_cross_average = '-'

        res_list.append({
            title_name: title,
            'count': len(paper_list),
            # 当前发文占比中总文献量未去重
            # '发文占比': str(round((len(paper_list) / len(all_data_list)) * 100, 4)) + "%",
            'RID': ASJC_cross_average,
            'RCD': institution_cross_average,
        })

    res_list = sorted(res_list, key=lambda x: x['count'], reverse=True)

    pd.DataFrame(res_list).to_excel(f'./中间计算数据/{exp_file_name}.xlsx')


def load_top_IF(top_IF_path):
    top_percent_1 = pd.DataFrame(pd.read_excel(top_IF_path, sheet_name="top1%文献汇总"))
    top_percent_10 = pd.DataFrame(pd.read_excel(top_IF_path, sheet_name="top10%文献汇总"))
    top_percent_25 = pd.DataFrame(pd.read_excel(top_IF_path, sheet_name="top25%文献汇总"))

    return {
        'top_1%': [str(top_percent_1["Full Journal Title"][i].lower()) for i in range(top_percent_1.shape[0])],
        'top_10%': [str(top_percent_10["Full Journal Title"][i].lower()) for i in range(top_percent_10.shape[0])],
        'top_25%': [str(top_percent_25["Full Journal Title"][i].lower()) for i in range(top_percent_25.shape[0])]
    }


def journal_IF_analysis(data, title_name, exp_file_name):
    # 期刊分析：数量:总,有IF|IF:总平均,区间count,区间占比|TOP:count,占比|CNS:count,占比
    journal_IF_map = get_data('journal_IF_map')     # 期刊IF映射
    top_journals = load_top_IF('./journal/top期刊&CNS汇总.xlsx')    # 期刊top区间数据

    # CNS 数据
    main_CNS = ["Science", "Nature", "Cell"]
    sub_CNS = ["Cancer Cell", "Cell Genomics ", "Cell Host & Microbe", "Cell Metabolism", "Cell Reports",
               "Cell Stem Cell", "Current Biology", "Developmental Cell", "Immunity", "Molecular Cell",
               "Nature Biotechnology", "Nature Cell Biology", "Nature Communications", "Nature Genetics",
               "Nature Immunology", "Nature Medicine", "Nature Methods", "Nature Methods", "Nature Microbiology",
               "Nature Neuroscience ", "Nature Reviews Genetics", "Nature Reviews Immunology",
               "Nature Reviews Microbiology", "Nature Reviews Molecular Cell Biology",
               "Nature Structural & Molecular Biology", "Neuron", "Science Advances", "Science Immunology",
               "Science Signaling", "Science Translational Medicine", "Structure"]

    res_list = []

    for title, paper_list in data.items():
        IF_sum = 0
        not_IF_count = 0

        IF_interval_count = [0, 0, 0, 0]   # [>20, 10-20, 5-10, >5]
        IF_top_percent = [0, 0, 0]     # [1%, 10%, 25%]
        CNS_count = [0, 0]      # [主刊, 子刊]

        for paper in paper_list:
            paper_journal_name = str(paper['Scopus Source title']).lower()
            if journal_IF_map.get(paper_journal_name, 'none') != 'none':
                # IF 均值计算
                _IF = journal_IF_map[paper_journal_name]
                IF_sum += _IF

                # IF区间计算
                if _IF >= 20:
                    IF_interval_count[0] += 1
                    IF_interval_count[3] += 1
                elif 10 <= _IF < 20:
                    IF_interval_count[1] += 1
                    IF_interval_count[3] += 1
                elif 5 <= _IF < 10:
                    IF_interval_count[2] += 1
                    IF_interval_count[3] += 1

                # IF TOP百分比计算
                if paper_journal_name in top_journals['top_1%']:
                    IF_top_percent[0] += 1
                if paper_journal_name in top_journals['top_10%']:
                    IF_top_percent[1] += 1
                if paper_journal_name in top_journals['top_25%']:
                    IF_top_percent[2] += 1

                # CNS计算
                if paper_journal_name in map(lambda s: s.lower(), main_CNS):
                    CNS_count[0] += 1
                if paper_journal_name in map(lambda s: s.lower(), sub_CNS):
                    CNS_count[1] += 1
            else:
                not_IF_count += 1

        # IF均值计算
        have_IF_count = len(paper_list) - not_IF_count
        if have_IF_count != 0:
            IF_average = round(IF_sum / have_IF_count, 4)
        else:
            IF_average = '-'

        res_list.append({
            title_name: title,
            'count': len(paper_list),
            'have_IF_count': have_IF_count,
            '0': ' ',
            '平均IF': IF_average,
            'IF>=20': IF_interval_count[0],
            '10<=IF<20': IF_interval_count[1],
            '5<=IF<10': IF_interval_count[2],
            'IF>=5': IF_interval_count[3],
            '1': ' ',
            'IF>=10占比': IF_interval_count[0] / have_IF_count if have_IF_count != 0 else '-',
            '10<=IF<20占比': IF_interval_count[1] / have_IF_count if have_IF_count != 0 else '-',
            '5<=IF<10占比': IF_interval_count[2] / have_IF_count if have_IF_count != 0 else '-',
            'IF>=5占比': IF_interval_count[3] / have_IF_count if have_IF_count != 0 else '-',
            '2': ' ',
            'top_1%': IF_top_percent[0],
            'top_10%': IF_top_percent[1],
            'top_25%': IF_top_percent[2],
            '3': ' ',
            'top_1%占比': IF_top_percent[0] / have_IF_count if have_IF_count != 0 else '-',
            'top_10%占比': IF_top_percent[1] / have_IF_count if have_IF_count != 0 else '-',
            'top_25%占比': IF_top_percent[2] / have_IF_count if have_IF_count != 0 else '-',
            '4': ' ',
            'CNS主刊': CNS_count[0],
            'CNS子刊': CNS_count[1],
            '5': ' ',
            'CNS主刊占比': CNS_count[0] / have_IF_count if have_IF_count != 0 else '-',
            'CNS子刊占比': CNS_count[1] / have_IF_count if have_IF_count != 0 else '-'
        })

    res_list = sorted(res_list, key=lambda x: x['count'], reverse=True)

    pd.DataFrame(res_list).to_excel(f'./中间计算数据/{exp_file_name}.xlsx')


def institution_cooperate(data, exp_file_name):
    # 机构合作与主导力分析：第一作者count|占比；合作前五机构|count
    res_list = []
    year_list1 = ['2011', '2012', '2013', '2014', '2015']
    year_list2 = ['2016', '2017', '2018', '2019', '2020']
    for institution_name, paper_list in data.items():
        # 存放当前机构合作的其他机构名
        cooperate_ins = []
        # 存放机构索引以及count，索引与cooperate_ins中机构位置相对应，索引用于排序后映射
        cooperate_ins_count = []
        one_institution = [0, 0, 0]     # 第一机构, 总，2011-2015，2016-2020
        paper_year_count = [0, 0]       # 2011-2015，2016-2020

        for paper in paper_list:
            institutions_list = paper['Institutions'].split('| ')
            if institutions_list[0] == institution_name:
                one_institution[0] += 1
                if str(paper['Year']) in year_list1:
                    one_institution[1] += 1
                if str(paper['Year']) in year_list2:
                    one_institution[2] += 1

            if str(paper['Year']) in year_list1:
                paper_year_count[0] += 1
            if str(paper['Year']) in year_list2:
                paper_year_count[1] += 1

            # 遍历除自己外其他机构，如果与sub_ins中出现的机构合作则记数字
            for other_institution in set(institutions_list) - {institution_name}:
                if institution_name == 'University of California at Davis' and other_institution == 'University of California Office of the President':
                    continue
                if institution_name != 'University of California at Davis' and other_institution == 'University of California Office of the President':
                    other_institution = 'University of California at Davis'

                if other_institution not in cooperate_ins:
                    cooperate_ins.append(other_institution)
                    cooperate_ins_count.append([len(cooperate_ins_count), 1])
                else:
                    cooperate_ins_count[cooperate_ins.index(other_institution)][1] += 1

        cooperate_ins_count = sorted(cooperate_ins_count, key=lambda x: x[1], reverse=True)
        institution_cooperate_top5 = [cooperate_ins[ins_idx[0]] for ins_idx in cooperate_ins_count[:5]]
        institution_cooperate_top5_count = [str(int(ins_idx[1])) for ins_idx in cooperate_ins_count[:5]]

        res_list.append({
            'institution': institution_name,
            'count': len(paper_list),
            '1': ' ',
            'first_author': one_institution[0],
            'first_author%': one_institution[0] / len(paper_list),
            '2': ' ',
            '2011-2015 first_author': one_institution[1],
            '2011-2015 first_author%': one_institution[1] / paper_year_count[0] if paper_year_count[0] != 0 else '-',
            '3': ' ',
            '2016-2020 first_author': one_institution[2],
            '2016-2020 first_author%': one_institution[2] / paper_year_count[1] if paper_year_count[1] != 0 else '-',
            '4': ' ',
            'institution_cooperate_top5': '|'.join(institution_cooperate_top5),
            'institution_cooperate_top5_count': '|'.join(institution_cooperate_top5_count)
        })

    res_list = sorted(res_list, key=lambda x: x['count'], reverse=True)

    pd.DataFrame(res_list).to_excel(f'./中间计算数据/{exp_file_name}.xlsx')


def top5_country_topic_cluster_export():
    # 输出top5国家主题统计
    country_list = ['China', 'United States', 'Brazil', 'India', 'United Kingdom']
    for country in country_list:
        _list = [paper['Topic Cluster name'] for paper in (get_data('country_papers'))[country]]

        save_count(_list, f'./中间计算数据/{country} topic cluster count.xlsx', name='name', num=sys.maxsize)


def topic_top3_country_and_count():
    top20_topic = [
        'Bacillus Thuringiensis,Aphidoidea,Aphids',
        'Viruses,Mosaic Viruses,Phytoplasma',
        'Phytophthora,Trichoderma,Phytophthora Infestans',
        'Weeds,Herbicides,Weed Control',
        'Arabidopsis,Plants,Genes',
        'Xanthomonas,Ralstonia Solanacearum,Genome',
        'Nematoda,Root-Knot Nematodes,Meloidogyne Incognita',
        'China,Fungi,Leaves',
        'Wheat,Triticum,Triticum Aestivum',
        'Baculoviridae,Entomopathogenic Nematodes,Nucleopolyhedrovirus',
        'Mycotoxins,Aflatoxins,Ochratoxins',
        'Plants,Rhizosphere,Rhizobium',
        'Formicidae,Hymenoptera,Ant',
        'Powdery Mildew,Sugar Beet,Fungicides',
        'Dengue,Viruses,Dengue Virus',
        'Hymenoptera,Galls,Braconidae',
        'Ticks,Lyme Disease,Borrelia Burgdorferi',
        'Phosphines,Tribolium Castaneum,Coleoptera',
        'Tephritidae,Fruit Flies,Diptera',
        'Salmonella,Escherichia Coli,Listeria Monocytogenes'
    ]

    top20_topic_count = [15038, 8450, 7892, 5780, 5139, 4317, 4311, 4003, 3935, 3925, 2992, 2936, 2859, 2603,
                         2440, 2439, 2281, 2194, 2181, 2173]

    res_list = []

    country_list = list((get_data('country_papers')).keys())
    # country_map = [[country, 0, idx] for idx, country in enumerate((get_data('country_papers')).keys())]

    # 第一次输出结果得到
    use_country_list = ['United States', 'China', 'Brazil', 'India', 'Australia', 'Germany', 'France', 'Italy',
                        'United Kingdom']

    topic_cluster_papers = get_data('topic_cluster_papers')

    for topic_idx, topic in enumerate(top20_topic):
        # use_country_map = []
        country_map = [[country, 0, idx] for idx, country in enumerate((get_data('country_papers')).keys())]

        for paper in topic_cluster_papers[topic]:
            for country in paper['Country/Region'].split('| '):
                country_map[country_list.index(country)][1] += 1

        # 前三位国家输出
        #     country_map = sorted(country_map, key=lambda x: x[1], reverse=True)
        #     print(country_map[:3])
        #     res_list.append({
        #         'topic_cluster': topic,
        #         'country1': country_map[0][0],
        #         'country2': country_map[1][0],
        #         'country3': country_map[2][0],
        #         'count1': country_map[0][1],
        #         'count2': country_map[1][1],
        #         'count3': country_map[2][1],
        #     })
        # pd.DataFrame(res_list).to_excel('./中间计算数据/Top20 topic_cluster 国家Top3分布.xlsx')

        cur = {'topic_cluster': topic, 'count': top20_topic_count[topic_idx]}

        for use_country in use_country_list:
            cur[use_country] = country_map[country_list.index(use_country)][1]

        res_list.append(cur)

    pd.DataFrame(res_list).to_excel('./中间计算数据/Top20 topic_cluster 全部Top3国家数量.xlsx')


def scival_to_wos_c1():
    data = get_data('sub_institutions_papers')
    ins_map = {
        'University of Florida': '佛罗里达大',
        'University of California at Davis': '加大戴维斯分校',
        'Universidade de São Paulo': '圣保罗大',
        'Universidade Federal de Viçosa': '维索萨联邦大',
        'Wageningen University & Research': '瓦格宁根大',
        'Cornell University': '康奈尔大',
        'Nanjing Agricultural University': '南农大',
        'Northwest Agriculture and Forestry University': '西北农林',
        'China Agricultural University': '中农大',
        'Zhejiang University': '浙大',
        'South China Agricultural University': '华南农',
        'Huazhong Agricultural University': '华中农'
    }
    # res_text = 'FN Clarivate Analytics Web of Science\nVR 1.0\n'
    #
    # all_paper_list = []
    # for ins_paper_list in data.values():
    #     all_paper_list += [f"{paper['Institutions']}|||{paper['EID']}" for paper in ins_paper_list]
    #
    # all_paper_list = list(set(all_paper_list))
    #
    # for paper_ins_text in [paper.split('|||')[0] for paper in all_paper_list]:
    #     cur_ins_num = 0  # 标记当前有几个合作机构，用于生成C1字段作者姓名
    #     is_one_ins = True
    #     res_text += 'PT J\n'
    #     for ins in paper_ins_text.split('| '):
    #         if ins in ins_map.keys():
    #             ins = ins_map[ins]
    #         # 遍历每个机构
    #         cur_ins_num += 1
    #         if is_one_ins:
    #             is_one_ins = False
    #             res_text += f"C1 [{'a' * cur_ins_num}] {ins}\n"
    #         else:
    #             res_text += f"   [{'a' * cur_ins_num}] {ins}\n"
    #
    #     res_text += 'ER\n\n'
    #
    # res_text += '\nEF'

    res_text = 'FN Clarivate Analytics Web of Science\nVR 1.0\n'
    for _list in data.values():
        # 每个机构的文献列表
        paper_ins_list = [paper['Institutions'].split('| ') for paper in _list]
        for ins_list in paper_ins_list:
            # 遍历每篇文献
            cur_ins_num = 0     # 标记当前有几个合作机构，用于生成C1字段作者姓名
            is_one_ins = True
            res_text += 'PT J\n'

            for ins in ins_list:
                # 遍历每个机构
                if ins in ins_map.keys():
                    ins = ins_map[ins]
                cur_ins_num += 1
                if is_one_ins:
                    is_one_ins = False
                    res_text += f"C1 [{'a'*cur_ins_num}] {ins}\n"
                else:
                    res_text += f"   [{'a' * cur_ins_num}] {ins}\n"

            res_text += 'ER\n\n'
        # break

    res_text += '\nEF'

    exp_text = open('./file.txt', 'w', encoding='utf-8')
    exp_text.write(res_text)
    exp_text.close()


def wos_sub_institutions_analyse():
    SciVal_wos_map = {
        'University of Florida': ['Univ Florida'],
        'University of California at Davis': ['Univ Calif Davis'],
        'Universidade de São Paulo': ['Univ Sao Paulo'],
        'Universidade Federal de Viçosa': ['Univ Fed Vicosa'],
        'Wageningen University & Research': ['Wageningen Univ', 'Wageningen Univ & Res', 'Wageningen Univ & Res Ctr WUR', 'Univ Wageningen & Res Ctr', 'Wageningen UR', 'Wageningen UR Plant Breeding', 'Univ Wageningen & Res Ctr WUR', 'Wageningen Univ & Res Ctr', 'Wageningen UR Appl Plant Res', 'Wageningen Univ Res'],
        'Cornell University': ['Cornell Univ'],
        'Nanjing Agricultural University': ['Nanjing Agr Univ'],
        'Northwest Agriculture and Forestry University': ['Northwest Agr & Forestry Univ', 'Northwest A&F Univ'],
        'China Agricultural University': ['China Agr Univ'],
        'Zhejiang University': ['Zhejiang Univ'],
        'South China Agricultural University': ['South China Agr Univ', 'S China Agr Univ'],
        'Huazhong Agricultural University': ['Huazhong Agr Univ']
    }
    ins_Chinese_map = {
        'University of Florida': '佛罗里达大学',
        'University of California at Davis': '加州大学戴维斯分校',
        'Universidade de São Paulo': '巴西圣保罗大学',
        'Universidade Federal de Viçosa': '维索萨联邦大学',
        'Wageningen University & Research': '瓦格宁根大学',
        'Cornell University': '康奈尔大学',
        'Nanjing Agricultural University': '南京农业大学',
        'Northwest Agriculture and Forestry University': '西北农林科技大学',
        'China Agricultural University': '中国农业大学',
        'Zhejiang University': '浙江大学',
        'South China Agricultural University': '华南农业大学',
        'Huazhong Agricultural University': '华中农业大学'
    }
    print('%10s' % '学校', '%6s' % '第一作者', '%6s' % '通讯作者', '%6s' % '第一或通讯')
    year_2011_2015 = ['2011', '2012', '2013', '2014', '2015']
    year_2016_2020 = ['2016', '2017', '2018', '2019', '2020']
    exp_list = []

    for ins_file in os.listdir('./wos重点大学/sub ins/'):
        if '_St' in ins_file or '~$' in ins_file:
            continue
        ins_name = ins_file[:-5]

        data = pd.read_excel(f'./wos重点大学/sub ins/{ins_file}')

        year_paper = {'2011': 0, '2012': 0, '2013': 0, '2014': 0, '2015': 0, '2016': 0, '2017': 0, '2018': 0, '2019': 0, '2020': 0}
        year_one = {'2011': 0, '2012': 0, '2013': 0, '2014': 0, '2015': 0, '2016': 0, '2017': 0, '2018': 0,'2019': 0, '2020': 0}
        year_RP = {'2011': 0, '2012': 0, '2013': 0, '2014': 0, '2015': 0, '2016': 0, '2017': 0, '2018': 0,'2019': 0, '2020': 0}
        year_one_or_RP = {'2011': 0, '2012': 0, '2013': 0, '2014': 0, '2015': 0, '2016': 0, '2017': 0, '2018': 0,'2019': 0, '2020': 0}

        for i in range(data.shape[0]):
            one_RP_or_ins = False
            if str(data['PY'][i]) == 'nan':
                continue
            year = str(int(data['PY'][i]))
            year = year if year != '2021' else '2020'
            year = year if year != '2010' else '2011'
            year_paper[year] += 1

            RP_list = str(data['RP'][i]).split(', ')
            for idx, cub_RP in enumerate(RP_list):
                if "(corresponding author)" in cub_RP:
                    if RP_list[idx+1] in SciVal_wos_map[ins_name]:
                        one_RP_or_ins = True
                        year_RP[year] += 1
                        break
                    
            C1_list = str(data['C1'][i]).split('   ')
            one_ins_match_list = re.findall('] ([\w\W]+?), ', C1_list[0])
            one_ins = ''
            if len(one_ins_match_list):
                one_ins = one_ins_match_list[0]
            if one_ins in SciVal_wos_map[ins_name]:
                one_RP_or_ins = True
                year_one[year] += 1

            if one_RP_or_ins:
                year_one_or_RP[year] += 1

        paper_count = reduce(lambda _all, num: _all + num, year_paper.values(), 0)
        one_ins_count = reduce(lambda _all, num: _all + num, year_one.values(), 0)
        one_RP_count = reduce(lambda _all, num: _all + num, year_RP.values(), 0)
        one_ins_or_RP_count = reduce(lambda _all, num: _all + num, year_one_or_RP.values(), 0)

        year_2011_2015_paper = reduce(lambda _all, num: _all + num,
            [count for year, count in year_paper.items() if year in year_2011_2015], 0)
        year_2016_2020_paper = reduce(lambda _all, num: _all + num,
            [count for year, count in year_paper.items() if year in year_2016_2020], 0)

        one_2011_2015_ins_count = reduce(lambda _all, num: _all + num,
            [count for year, count in year_one.items() if year in year_2011_2015], 0)
        one_2016_2020_ins_count = reduce(lambda _all, num: _all + num,
            [count for year, count in year_one.items() if year in year_2016_2020], 0)

        one_2011_2015_RP_count = reduce(lambda _all, num: _all + num,
            [count for year, count in year_RP.items() if year in year_2011_2015], 0)
        one_2016_2020_RP_count = reduce(lambda _all, num: _all + num,
            [count for year, count in year_RP.items() if year in year_2016_2020], 0)

        one_2011_2015_ins_or_RP_count = reduce(lambda _all, num: _all + num,
            [count for year, count in year_one_or_RP.items() if year in year_2011_2015], 0)
        one_2016_2020_ins_or_RP_count = reduce(lambda _all, num: _all + num,
            [count for year, count in year_one_or_RP.items() if year in year_2016_2020], 0)

        # print(ins_Chinese_map[ins_name])
        # print(year_2011_2015_paper, year_2016_2020_paper, year_paper)
        # print(one_2011_2015_ins_count, one_2016_2020_ins_count, year_one)
        # print(one_2011_2015_RP_count, one_2016_2020_RP_count, year_RP)
        # print(one_2011_2015_ins_or_RP_count, one_2016_2020_ins_or_RP_count, year_one_or_RP)
        # print('------------------------------')
        print('%10s' % ins_Chinese_map[ins_name], '%f' % round(one_ins_count / paper_count, 4),'%5f' % round(one_RP_count / paper_count, 4), '%f' % round(one_ins_or_RP_count / paper_count, 4))

        exp_list.append({
            '机构': ins_name,
            '机构中文': ins_Chinese_map[ins_name],
            'wos文献量': paper_count,
            '第一作者数量': one_ins_count,
            '通讯作者数量': one_RP_count,
            '第一或通讯数量': one_ins_or_RP_count,
            ' ': ' ',
            '第一作者占比': one_ins_count / paper_count,
            '通讯作者占比': one_RP_count / paper_count,
            '第一或通讯占比': one_ins_or_RP_count / paper_count,
            '  ': ' ',
            '2011-2015第一作者数量': one_2011_2015_ins_count,
            '2011-2015通讯作者数量': one_2011_2015_RP_count,
            '2011-2015第一或通讯数量': one_2011_2015_ins_or_RP_count,
            '   ': ' ',
            '2011-2015第一作者占比': one_2011_2015_ins_count / year_2011_2015_paper,
            '2011-2015通讯作者占比': one_2011_2015_RP_count / year_2011_2015_paper,
            '2011-2015第一或通讯占比': one_2011_2015_ins_or_RP_count / year_2011_2015_paper,
            '    ': ' ',
            '2016-2020第一作者数量': one_2016_2020_ins_count,
            '2016-2020通讯作者数量': one_2016_2020_RP_count,
            '2016-2020第一或通讯数量': one_2016_2020_ins_or_RP_count,
            '     ': ' ',
            '2016-2020第一作者占比': one_2016_2020_ins_count / year_2016_2020_paper,
            '2016-2020通讯作者占比': one_2016_2020_RP_count / year_2016_2020_paper,
            '2016-2020第一或通讯占比': one_2016_2020_ins_or_RP_count / year_2016_2020_paper,
            '      ': ' ',
            '文献数量分布': '|'.join([str(num) for num in year_paper.values()]),
            '第一作者分布': '|'.join([str(num) for num in year_one.values()]),
            '通讯作者分布': '|'.join([str(num) for num in year_RP.values()]),
            '第一或通讯分布': '|'.join([str(num) for num in year_one_or_RP.values()])
        })

        pd.DataFrame(exp_list).to_excel('./wos重点大学/wos主导力分析.xlsx')


if __name__ == '__main__':
    # save_data()
    # journal_export()
    # institutions_papers_export()
    # sub_institutions_papers_export()
    # country_papers_export()
    # sub_country_papers_export()
    # topic_cluster_papers_export()
    # sub_topic_cluster_papers_export()

    ################ 主题topic ASJC矩阵输出 ################
    # data = get_data('topic_cluster_papers')
    # topic_list = [
    #     'Bacillus Thuringiensis,Aphidoidea,Aphids',
    #     'Viruses,Mosaic Viruses,Phytoplasma',
    #     'Phytophthora,Trichoderma,Phytophthora Infestans',
    #     'Weeds,Herbicides,Weed Control',
    #     'Arabidopsis,Plants,Genes',
    #     'Xanthomonas,Ralstonia Solanacearum,Genome',
    #     'Nematoda,Root-Knot Nematodes,Meloidogyne Incognita',
    #     'China,Fungi,Leaves',
    #     'Wheat,Triticum,Triticum Aestivum',
    #     'Baculoviridae,Entomopathogenic Nematodes,Nucleopolyhedrovirus',
    #     'Mycotoxins,Aflatoxins,Ochratoxins',
    #     'Plants,Rhizosphere,Rhizobium',
    #     'Formicidae,Hymenoptera,Ant',
    #     'Powdery Mildew,Sugar Beet,Fungicides',
    #     'Dengue,Viruses,Dengue Virus',
    #     'Hymenoptera,Galls,Braconidae',
    #     'Ticks,Lyme Disease,Borrelia Burgdorferi',
    #     'Phosphines,Tribolium Castaneum,Coleoptera',
    #     'Tephritidae,Fruit Flies,Diptera',
    #     'Salmonella,Escherichia Coli,Listeria Monocytogenes',
    #     'Mites,Acari,Oribatida',
    #     'Beetle,Bark Beetles,Curculionidae',
    #     'Fungi,Magnaporthe,Oryza Sativa'
    # ]
    # for topic in topic_list:
    #     ASJC_matrix(data[topic], topic)


    ################ 论文doi获取 ################
    # data = get_data('all_data')
    # f = open('./doi', 'w')
    # for paper_list in data.values():
    #     for paper in paper_list:
    #         if paper['DOI'] != '-':
    #             f.writelines(f"DO=({str(paper['DOI'])}) OR\n")

    ################ 代码汇总及说明 ################
    # journals_statistics()

    # 求：发文量:总和,占比,每年|被引频次:平均,每年|FWCI:区间(0,5,10),占比,每年,前后五年均值,增长率
    # FWCI_and_TC_statistics(data, _type, file_name)

    # data = get_data('country_papers')
    # FWCI_and_TC_statistics(data, 'institution', '国家FWCI_TC计算')

    # 求topic_cluster FWCI最高的国家
    # topic_cluster_highest_FWCI_country_statistics()

    # 论文国际合作:数量,前后五年发文量,合作数,合作占比
    # country_cooperate_analysis(data, 'country', 'country_cooperate111')

    # 输入某一单元文献列表（list），输出ASJC交叉矩阵
    # ASJC_matrix(paper_list, exp_file_name)

    # 统计ASJC跨学科性和机构交叉情况
    # ASJC_and_institution_cross(data, title_name, exp_file_name):

    # 期刊分析：数量:总,有IF|IF:总平均,区间count,区间占比|TOP:count,占比|CNS:count,占比
    # journal_IF_analysis(data, title_name, exp_file_name):

    # 机构合作与主导力分析：第一作者count|占比；合作前五机构|count
    # institution_cooperate(data, exp_file_name):

    ################ 报告分析代码 ################
    # FWCI_and_TC_statistics(get_data('country_papers'), 'country', '国家分析')
    # topic_cluster_highest_FWCI_country_statistics()
    # country_cooperate_analysis(get_data('country_papers'), 'country', '国家国际合作分析')
    # FWCI_and_TC_statistics(get_data('sub_institutions_papers'), 'institution', '机构分析')
    # journal_IF_analysis(get_data('sub_institutions_papers'), 'institution', '机构期刊分析')
    # institution_cooperate(get_data('sub_institutions_papers'), '机构主导力分析')
    # FWCI_and_TC_statistics(get_data('sub_topic_cluster_papers'), 'topic_cluster', 'topic_cluster FWCI分布')
    # country_cooperate_analysis(get_data('sub_topic_cluster_papers'), 'topic_cluster', 'topic_cluster国际合作分析')
    # ASJC_and_institution_cross(get_data('sub_topic_cluster_papers'), 'topic_cluster', 'topic_cluster跨学科性')

    # country_cooperate_analysis(get_data('topic_cluster_papers'), 'topic_cluster', 'topic_cluster国际合作分析（全）')
    # ASJC_and_institution_cross(get_data('topic_cluster_papers'), 'topic_cluster', 'topic_cluster跨学科性（全）')
    # FWCI_and_TC_statistics(get_data('topic_cluster_papers'), 'topic_cluster', 'topic_cluster FWCI分布（全）')
    # FWCI_and_TC_statistics(get_data('all_data'), 'year', '年份分布')

    # top5_country_topic_cluster_export()
    # topic_top3_country_and_count()

    # ASJC_and_institution_cross(get_data('all_data'), 'year', '文献年度跨学科性')
    # country_cooperate_analysis(get_data('all_data'), 'topic_cluster', '文献年度国际合作分析')

    # data = pd.read_excel('/Users/wangzhuoran/同名作者消歧/中文文献.xlsx')
    # _ = reduce(lambda _all, kw: _all + str(kw).split(';'), data['关键词'], [])
    # save_count(_, '/Users/wangzhuoran/同名作者消歧/关键词.xlsx', name='name', num=10)

    # scival_to_wos_c1()
    # sub_institutions_papers_doi_file_export()

    wos_sub_institutions_analyse()










