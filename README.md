# 植物保护学科发展报告
植物保护发展报告数据与源代码 ——2021情报学方法与技术

## 一、数据：

### 1. 2011-2020 植物保护学科 SciVal 数据

链接：[scival data](./scival%20data)

### 2. 2011-2020 植物保护 WOS 高被引文献

链接：[wos highly cited papers](./wos%20highly%20cited%20papers)

### 3. 代码分析输出数据

链接：[analysis data output](./analysis%20data%20output)

## 二、核心代码：

### 1. [scival analyse.py](./scival%20analyse.py) 

SciVal 数据主要分析代码，重要代码如下：

* _papers_export()

  * 按需求导出文献对象，结构如下：

    ```python
    all_data = {
        'title1': [page_data1, page_data2, ....],
        'title2': [page_data1, page_data2, ....],
        ....
        'title3': [page_data1, page_data2, ....]
    }
    ```

* get_data(path)

  * 获取文献对象

* FWCI_and_TC_statistics(data, _type, file_name)

  * 求：发文量:总和,占比,每年|被引频次:平均,每年|FWCI:区间(0,5,10),占比,每年,前后五年均值,增长率

* topic_cluster_highest_FWCI_country_statistics()

  * 求topic_cluster FWCI最高的国家

* country_cooperate_analysis(data, title_name, exp_file_name)

  * 论文国际合作:数量,前后五年发文量,合作数,合作占比

* ASJC_matrix(paper_list, exp_file_name)

  * 输入某一单元文献列表（list），输出ASJC交叉矩阵

* ASJC_and_institution_cross(data, title_name, exp_file_name)

  * 统计ASJC跨学科性和机构交叉情况

* journal_IF_analysis(data, title_name, exp_file_name):

  * 期刊分析：数量:总,有IF|IF:总平均,区间count,区间占比|TOP:count,占比|CNS:count,占比

* institution_cooperate(data, exp_file_name):

  * 机构合作与主导力分析：第一作者count|占比；合作前五机构|count

* scival_to_wos_c1()

  * 将 scival 文件中 Institutions 字段机构抽取为 wos 题录数据格式，用于WOSViewer可视化分析

### 2. [show.py](./show.py)

SciVal 数据主要分析代码，重要代码如下：

* topic_cluster_show()：按要求输出主题 RID/RCD 分布象限图
* year_show()：输出文献全球占比图
* country_more_1000()：输出发文超过1000国家 FWCI/RAI 分布图
* country_RAI_and_FWCI()：输出国家 RAI/FWCI 前后五年变化对比图
* tree_map_draw()：输出重点国家主题分布矩形树图
* ins_prev_and_next_show()：输出重点机构前后五年第一作者对比图
