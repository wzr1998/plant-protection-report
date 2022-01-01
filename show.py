# coding=utf-8
import pandas as pd
import numpy as np
import squarify
from matplotlib.font_manager import FontProperties
import matplotlib.pyplot as plt


def year_show():
    data = pd.read_excel('./图表绘制.xlsx', sheet_name='年份分布')
    font = FontProperties(fname='./simhei.ttf', size=6.5)
    year = data['year']
    population_by_continent = {
        '全球占比': data['全球占比'],
    }

    fig, ax = plt.subplots()
    ax.stackplot(year, population_by_continent.values(),
                 labels=population_by_continent.keys(), alpha=0.8)
    ax.legend(loc='upper left', prop=font)
    ax.set_title('World population')
    ax.set_xlabel('Year')
    ax.set_ylabel('Number of people (millions)')

    plt.show()


def country_more_1000():
    # 发文量占比超1%的国家的FWCI和RAI关系图（P38, Figure 15）
    fig, ax = plt.subplots()
    data = pd.read_excel('./图表绘制.xlsx', sheet_name='国家发文>1000')

    RAI = data['RAI']

    FWCI = data['FWCI']
    count = data['count']
    country = data['country']
    font = FontProperties(fname='./simhei.ttf', size=7)


    for i in range(len(list(RAI))):
        ax.scatter(RAI[i], FWCI[i], count[i]/20, edgecolors='none', alpha=0.5)
        ax.text(RAI[i], FWCI[i], country[i], horizontalalignment='center', verticalalignment='center', font=font)

    # 删除边框
    ax.spines['right'].set_color('none')
    ax.spines['top'].set_color('none')
    ax.spines['left'].set_color('none')
    ax.spines['bottom'].set_color('none')
    # 删除刻度
    ax.tick_params(bottom=False, top=False, left=False, right=False)

    ax.spines['bottom'].set_position(('axes', 0))  #data表示通过值来设置x轴的位置，将x轴绑定在y=0的位置
    ax.spines['left'].set_position(('axes', 0))  #axes表示以百分比的形式设置轴的位置，即将y轴绑定在x轴50%的位置，也就是x轴的中点

    ax.axhline(y=1, color='gray', linestyle='--', lw=0.8)
    ax.axvline(x=1, color='gray', linestyle='--', lw=0.8)

    # plt.figtext(0.05, 0.4, 'Global average', c='gray', fontsize=6)
    # plt.figtext(0.36, 0.68, 'Global average', c='gray', fontsize=6)

    plt.xlabel("RAI(Relative Actively Index)", size=10)
    plt.ylabel("FWCI", size=10)

    plt.show()

    fig.savefig('./图像输出/国家RAI&FWCI.jpg', dpi=600, pil_kwargs={'quality': 95})


def country_RAI_and_FWCI():
    # 发文量占比超1%的国家前后五年的RAI变动（P39, Figure 16）
    fig, ax = plt.subplots()
    data = pd.read_excel('./图表绘制.xlsx', sheet_name='国家RAI&FWCI')

    country = data['country']
    _2011_2015RAI = data['2011-2015RAI']
    _2016_2020RAI = data['2016-2020RAI']

    font = FontProperties(fname='./simhei.ttf', size=7)

    _len = len(list(country))
    for i in range(_len-1):
        ax.scatter(i-0.5+0.22, _2011_2015RAI[i], 50, c='#87d930', edgecolors='none', alpha=1)
        ax.scatter(i-0.5+0.22, _2016_2020RAI[i], 50, c='#456518', edgecolors='none', alpha=1)
        ax.axvline(x=i+0.22, color='#dddddd', linestyle='--', lw=0.8)

    ax.scatter((_len-1)-0.5+0.22, _2011_2015RAI[_len-1], 50, edgecolors='none', c='#87d930', label='RAI 2011-2015')
    ax.scatter((_len-1)-0.5+0.22, _2016_2020RAI[_len-1], 50, edgecolors='none', c='#456518', label='RAI 2016-2020')

    # 删除边框
    ax.spines['right'].set_color('none')
    ax.spines['top'].set_color('none')
    ax.spines['left'].set_color('#dddddd')
    ax.spines['bottom'].set_color('#dddddd')
    # 删除刻度
    ax.tick_params(bottom=False, top=False, left=False, right=False)

    ax.spines['bottom'].set_position(('axes', 0))  #data表示通过值来设置x轴的位置，将x轴绑定在y=0的位置
    ax.spines['left'].set_position(('axes', 0))  #axes表示以百分比的形式设置轴的位置，即将y轴绑定在x轴50%的位置，也就是x轴的中点

    ax.axhline(y=1, color='#cccccc', linestyle='--', lw=0.8)
    plt.figtext(0.11, 0.256, 'Global average', c='gray', fontsize=6)


    plt.xlabel("Country", size=10)
    plt.ylabel("RAI(Relative Actively Index)", size=10)

    # 自定义坐标轴
    ticks = list(range(_len))
    ax.set_xticks(ticks)
    ax.set_xticklabels(country, rotation=90, horizontalalignment='right', fontsize=6)

    plt.legend(loc=1, markerscale=1, prop=font)

    plt.show()

    fig.savefig('./图像输出/国家RAI区间.jpg', dpi=600, pil_kwargs={'quality': 95})


def ins_prev_and_next_show():
    # 发文量占比超1%的国家前后五年的RAI变动（P39, Figure 16）
    fig, ax = plt.subplots()
    font = FontProperties(fname='./simhei.ttf', size=6.5)
    data = pd.read_excel('./图表绘制.xlsx', sheet_name='国家RAI&FWCI')

    country = ['Univ Florida', 'Calif Davis', 'Univ Sao Paulo', 'Univ Fed Vicosa', 'Wageningen Univ', 'Cornell Univ', 'Nanjing Agr Univ', 'Northwest Agr & Forestry Univ', 'China Agr Univ', 'Zhejiang Univ', 'South China Agr Univ', 'Huazhong Agr Univ']
    _2011_2015RAI = [0.5982, 0.4392, 0.4368, 0.6691, 0.4042, 0.5132, 0.2130, 0.5521, 0.7000, 0.7083, 0.7418, 0.7034]
    _2016_2020RAI = [0.5656, 0.4990, 0.5236, 0.6361, 0.4825, 0.5469, 0.5004, 0.8453, 0.7387, 0.6705, 0.7891, 0.7906]

    font = FontProperties(fname='./simhei.ttf', size=7)

    _len = len(list(country))
    for i in range(_len-1):
        ax.scatter(i-0.5+0.22, _2011_2015RAI[i], 100, c='#87d930', edgecolors='none', alpha=1)
        ax.scatter(i-0.5+0.22, _2016_2020RAI[i], 100, c='#456518', edgecolors='none', alpha=1)
        ax.axvline(x=i+0.22, color='#dddddd', linestyle='--', lw=0.8)

    ax.scatter((_len-1)-0.5+0.22, _2011_2015RAI[_len-1], 100, edgecolors='none', c='#87d930', label='2011-2015')
    ax.scatter((_len-1)-0.5+0.22, _2016_2020RAI[_len-1], 100, edgecolors='none', c='#456518', label='2016-2020')

    # 删除边框
    ax.spines['right'].set_color('none')
    ax.spines['top'].set_color('none')
    ax.spines['left'].set_color('#dddddd')
    ax.spines['bottom'].set_color('#dddddd')
    # 删除刻度
    ax.tick_params(bottom=False, top=False, left=False, right=False)

    ax.spines['bottom'].set_position(('axes', 0))  #data表示通过值来设置x轴的位置，将x轴绑定在y=0的位置
    ax.spines['left'].set_position(('axes', 0))  #axes表示以百分比的形式设置轴的位置，即将y轴绑定在x轴50%的位置，也就是x轴的中点

    # ax.axhline(y=1, color='#cccccc', linestyle='--', lw=0.8)
    # plt.figtext(0.11, 0.256, 'Global average', c='gray', fontsize=6)


    plt.xlabel("Institutions", size=10)
    plt.ylabel("First Author or Corresponding Author %", size=10)
    plt.ylim(0, 1)

    # 自定义坐标轴
    ticks = list(range(_len))
    ax.set_xticks(ticks)
    ax.set_xticklabels(country, rotation=90, horizontalalignment='right', fontsize=6)

    plt.legend(loc=4, markerscale=1, prop=font)
    # plt.legend(loc=4, markerscale=0.3, prop=font)

    plt.show()

    fig.savefig('./图像输出/第一或通讯作者变化.jpg', dpi=600, pil_kwargs={'quality': 95})


def treemap_show(labels, income, jpg_name):

    colors = ['#7cb5ec', '#434348', '#90ed7d', '#f7a35c', '#8085e9', '#f15c80', '#e4d354', '#8085e8', '#8d4653',
              '#91e8e1', '#7cb5ec', '#434348', '#90ed7d', '#f7a35c', '#8085e9', '#f15c80', '#e4d354', '#8085e8',
              '#8d4653', '#91e8e1']

    fig = plt.figure(figsize=(12, 6))
    ax = fig.add_subplot(111)
    plot = squarify.plot(sizes=income,  # 方块面积大小
                         label=labels,  # 指定标签
                         color=colors,  # 指定自定义颜色
                         alpha=0.8,  # 指定透明度
                         # value=income,  # 添加数值标签
                         edgecolor='white',  # 设置边界框
                         linewidth=0.5,  # 设置边框宽度
                         text_kwargs={'fontsize': 10, 'verticalalignment': 'top', 'in_layout': True}
                         )
    ax.axis('off')
    ax.tick_params(top='off', right='off')
    plt.show()
    fig.savefig(f'./图像输出/{jpg_name}.jpg', dpi=600, pil_kwargs={'quality': 95})


def tree_map_draw():
    ######################## China ##################
    labels_China = [
        'Bacillus Thuringiensis,Aphidoidea,Aphids\n2629',
        'Arabidopsis,Plants,Genes\n1630',
        'Viruses,Mosaic Viruses,Phytoplasma\n1287',
        'Phytophthora,Trichoderma,Phytophthora Infestans\n1260',
        'China,Fungi,Leaves\n834',
        'Wheat,Triticum,Triticum Aestivum\n770',
        'Xanthomonas,\nRalstonia Solanacearum,\nGenome\n663',
        'Nematoda,\nRoot-Knot Nematodes,\nMeloidogyne Incognita\n607',
        'Fungi,Magnaporthe,Oryza Sativa\n536',
        'Plants,Rhizosphere,Rhizobium\n502',
        'Baculoviridae,\nEntomopathogenic Nematodes,\nNucleopolyhedrovirus\n425',
        'Powdery Mildew,\nSugar Beet,Fungicides\n418',
        'Weeds,Herbicides,Weed Control\n390',
        'Mycotoxins,\nAflatoxins,\nOchratoxins\n355',
        'Ethylenes,Apples,\nFruit\n291',
        'Human Influenza,\nOrthomyxoviridae,\nInfluenza Vaccines\n282',
        'Mycorrhizal Fungi,\nBasidiomycota,\nFungus\n274',
        'Fungi,Endophytes,Aspergillus\n254',
        'Hymenoptera,Galls,\nBraconidae\n244',
        'Rotavirus,\nNorovirus,\nCoronavirus\n23'
    ]
    income_China = [2629, 1630, 1287, 1260, 834, 770, 663, 607, 536, 502, 425, 418, 390, 355, 291, 282, 274, 254, 244, 237]
    # treemap_show(labels_China, income_China, '中国 Topic')

    ######################## US ##################
    labels_US = [
        'Bacillus Thuringiensis,Aphidoidea,Aphids\n3655',
        'Weeds,Herbicides,Weed Control\n1844',
        'Viruses,Mosaic Viruses,Phytoplasma\n1822',
        'Phytophthora,Trichoderma,\nPhytophthora Infestans\n1776',
        'Xanthomonas,\nRalstonia Solanacearum,\nGenome\n1022',
        'Wheat,Triticum,\nTriticum Aestivum\n993',
        'Hymenoptera,\nGalls,Braconidae\n897',
        'Nematoda,\nRoot-Knot Nematodes,\nMeloidogyne Incognita\n860',
        'Beetle,Bark Beetles,\nCurculionidae\n857',
        'Ticks,Lyme Disease,\nBorrelia Burgdorferi\n845',
        'Arabidopsis,Plants,Genes\n822',
        'Baculoviridae,\nEntomopathogenic Nematodes,\nNucleopolyhedrovirus\n795',
        'Formicidae,Hymenoptera,Ant\n789',
        'Dengue,Viruses,Dengue Virus\n718',
        'Salmonella,\nEscherichia Coli,\nListeria Monocytogenes\n702',
        'Powdery Mildew,Sugar Beet,\nFungicides\n646',
        'Mycotoxins,Aflatoxins,\nOchratoxins\n608',
        'Tephritidae,\nFruit Flies,Diptera\n602',
        'China,Fungi,Leaves\n593',
        'Human Influenza,\nOrthomyxoviridae,\nInfluenza Vaccines\n497'
    ]
    income_US = [3655,1844,1822,1776,1022,993,897,860,857,845,822,795,789,718,702,646,608,602,593,497]
    # treemap_show(labels_US, income_US, '美国 Topic')

    ######################## Brazil ##################
    labels_Brazil = [
        'Bacillus Thuringiensis,Aphidoidea,Aphids\n1457',
        'Weeds,Herbicides,Weed Control\n826',
        'Nematoda,Root-Knot Nematodes,\nMeloidogyne Incognita\n563',
        'China,Fungi,Leaves\n483',
        'Phytophthora,Trichoderma,\nPhytophthora Infestans\n459',
        'Viruses,Mosaic Viruses,\nPhytoplasma\n384',
        'Mites,Acari,Oribatida\n365',
        'Hymenoptera,Galls,Braconidae\n343',
        'Ticks,Lyme Disease,\nBorrelia Burgdorferi\n329',
        'Baculoviridae,\nEntomopathogenic Nematodes,\nNucleopolyhedrovirus\n324',
        'Jatropha,Jatropha Curcas,Brazil\n323',
        'Xanthomonas,\nRalstonia Solanacearum,Genome\n295',
        'Formicidae,Hymenoptera,Ant\n288',
        'Seeds,Germination,Seedlings\n275',
        'Coleoptera,Beetle,\nStaphylinidae\n248',
        'Arabidopsis,\nPlants,Genes\n245',
        'Powdery Mildew,\nSugar Beet,Fungicides\n244',
        'Candida,Infection,\nCandida Albicans\n232',
        'Agriculture,Fruits,\nAgricultural Machinery\n214',
        'Plants,Rhizosphere,\nRhizobium\n212'
    ]
    income_Brazil = [1457,826,563,483,459,384,365,343,329,324,323,295,288,275,248,245,244,232,214,212]
    # treemap_show(labels_Brazil, income_Brazil, 'Brazil Topic')

    ######################## India ##################
    labels_India = [
        'Bacillus Thuringiensis,Aphidoidea,Aphids\n1414',
        'Viruses,Mosaic Viruses,Phytoplasma\n889',
        'Phytophthora,Trichoderma,\nPhytophthora Infestans\n847',
        'Plants,Rhizosphere,Rhizobium\n572',
        'Weeds,Herbicides,Weed Control\n458',
        'Nematoda,Root-Knot Nematodes,\nMeloidogyne Incognita\n393',
        'Arabidopsis,Plants,Genes\n346',
        'Baculoviridae,\nEntomopathogenic Nematodes,\nNucleopolyhedrovirus\n340',
        'Limonins,Meliaceae,Azadirachta\n304',
        'Wheat,Triticum,\nTriticum Aestivum\n302',
        'Xanthomonas,\nRalstonia Solanacearum,Genome\n294',
        'China,Fungi,Leaves\n255',
        'Intercropping,\nNutrient Management,India\n241',
        'Rice,Corn,Wheat\n209',
        'Beans,Fabaceae,\nGenotype\n204',
        'Phosphines,Tribolium Castaneum,\nColeoptera\n192',
        'Fungi,Magnaporthe,\nOryza Sativa\n151',
        'Hemiptera,\nPseudococcidae,\nScale Insects\n144',
        'Mycotoxins,Aflatoxins,\nOchratoxins\n140',
        'Dengue,\nViruses,\nDengue Virus\n129'
    ]
    income_India = [1414,889,847,572,458,393,346,340,304,302,294,255,241,209,204,192,151,144,140,129]
    # treemap_show(labels_India, income_India, 'India Topic')

    ######################## UK ##################
    labels_UK = [
        'Bacillus Thuringiensis,Aphidoidea,Aphids\n683',
        'Phytophthora,Trichoderma,\nPhytophthora Infestans\n371',
        'Viruses,Mosaic Viruses,Phytoplasma\n279',
        'Arabidopsis,Plants,Genes\n274',
        'Wheat,Triticum,Triticum Aestivum\n238',
        'Dengue,Viruses,Dengue Virus\n235',
        'Formicidae,Hymenoptera,Ant\n220',
        'Xanthomonas,\nRalstonia Solanacearum,Genome\n204',
        'Malaria,Plasmodium Falciparum,\nParasites\n193',
        'Birds,Nests,Seabirds\n189',
        'Salmonella,Escherichia Coli,\nListeria Monocytogenes\n180',
        'Nematoda,Root-Knot Nematodes,\nMeloidogyne Incognita\n168',
        'Echinococcosis,Schistosomiasis,\nParasites\n168',
        'Baculoviridae,E\nntomopathogenic Nematodes,\nNucleopolyhedrovirus\n148',
        'Weeds,Herbicides,\nWeed Control\n135',
        'Ticks,Lyme Disease,\nBorrelia Burgdorferi\n121',
        'Carnivores,\nUngulates,Deer\n121',
        'Anti-Bacterial Agents,\nInfection,\nMethicillin-Resistant \nStaphylococcus Aureus\n118',
        'China,Fungi,Leaves\n113',
        'Human Influenza,\nOrthomyxoviridae,\nInfluenza Vaccines\n112'
    ]
    income_UK = [683,371,279,274,238,235,220,204,193,189,180,168,168,148,135,121,121,118,113,112]

    # treemap_show(labels_UK, income_UK, 'UK Topic')


def topic_cluster_show():
    # 绘制 topic_cluster 象限图
    fig, ax = plt.subplots()

    data = pd.read_excel('./图表绘制.xlsx', sheet_name='topic_cluster RCD')

    # RID = [0.9528, 0.9704, 1.0029, 0.9836, 1.0700, 1.0605, 0.9398, 0.9896, 1.1304, 0.9473]
    # RCD = [0.9872, 1.0105, 0.9442, 0.8642, 1.0350, 1.0953, 0.9718, 1.1243, 1.1117, 0.9020]
    # count = [15038, 8450, 7892, 5780, 5139, 4317, 4311, 4003, 3935, 3925]
    # topic_id = [129, 293, 425, 533, 11, 1001, 709, 704, 617, 679]
    # topic_name = [
    #     'Bacillus Thuringiensis,Aphidoidea,Aphids',
    #     'Viruses,Mosaic Viruses,Phytoplasma',
    #     'Phytophthora,Trichoderma,Phytophthora Infestans',
    #     'Weeds,Herbicides,Weed Control',
    #     'Arabidopsis,Plants,Genes',
    #     'Xanthomonas,Ralstonia Solanacearum,Genome',
    #     'Nematoda,Root-Knot Nematodes,Meloidogyne Incognita',
    #     'China,Fungi,Leaves',
    #     'Wheat,Triticum,Triticum Aestivum',
    #     'Baculoviridae,Entomopathogenic Nematodes,Nucleopolyhedrovirus'
    # ]

    RID = data['RID']
    RCD = data['RCD']
    count = data['count']
    topic_id = data['Topic Cluster number']
    topic_name = data['topic_cluster']

    for i in range(10):
        _ = u'%4d' % topic_id[i] + '\t-\t' + topic_name[i]
        ax.scatter(RID[i], RCD[i], count[i]/20, edgecolors='none', label=_, alpha=0.5)
        ax.text(RID[i], RCD[i], topic_id[i], horizontalalignment='center', verticalalignment='center')

    # 删除边框
    ax.spines['right'].set_color('none')
    ax.spines['top'].set_color('none')
    ax.spines['left'].set_color('none')
    ax.spines['bottom'].set_color('none')
    # 删除刻度
    ax.tick_params(bottom=False, top=False, left=False, right=False)

    ax.spines['bottom'].set_position(('axes', 0))  #data表示通过值来设置x轴的位置，将x轴绑定在y=0的位置
    ax.spines['left'].set_position(('axes', 0))  #axes表示以百分比的形式设置轴的位置，即将y轴绑定在x轴50%的位置，也就是x轴的中点

    ax.axhline(y=1, color='gray', linestyle='--', lw=0.8)
    ax.axvline(x=1, color='gray', linestyle='--', lw=0.8)

    plt.figtext(0.075, 0.35, 'RID(relative interdisciplinarity degree)', c='gray', fontsize=6)
    plt.figtext(0.348, 0.68, 'RCD(relative collaboration degree)', c='gray', rotation=90, fontsize=6)

    plt.xticks([0.95, 1, 1.05, 1.1])
    plt.yticks(([0.9, 1, 1.1]))

    font = FontProperties(fname='./simhei.ttf', size=6.5)

    plt.legend(loc=4, markerscale=0.3, prop=font)
    plt.show()

    fig.savefig('./图像输出/topic_cluster RCD 象限图.jpg', dpi=600, pil_kwargs={'quality': 95})


if __name__ == '__main__':
    # topic_cluster_show()
    # year_show()
    # country_more_1000()
    # country_RAI_and_FWCI()
    # tree_map_draw()
    ins_prev_and_next_show()


