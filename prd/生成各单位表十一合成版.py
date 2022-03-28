#!/usr/bin/env python3
# -*- coding:utf-8 -*-
# @Time :2021-12-11 9:05
# @Author:xinggevip
# @File : 生成各单位表十一py
# @Software: PyCharm

import pandas as pd
import time
import datetime


def start():
    start_time_main = time.time()

    print("1. 读取档案表中...")
    print("1.1 正在读取第一张表,主表")
    path = 'D:\\office\\各单位表十一\\数据源\\CAP-HR-03 员工档案表_1.xls'  # 全部
    new_open_path = 'D:\\office\\各单位表十一\\数据源\\52-HR-14 OA账号开通申请表_1.xls'  # 全部
    old_open_path = 'D:\\office\\各单位表十一\\数据源\\53-HR-15 OA账号恢复申请表_1.xls'  # 全部
    new_ruzhi_path = 'D:\\office\\各单位表十一\\数据源\\员工入职申请表（新新员工）黄亚楠_1.xls'  # 流转中
    old_ruzhi_path = 'D:\\office\\各单位表十一\\数据源\\员工入职申请表（新复职员工）黄亚楠_1.xls'  # 流转中
    bian_path = 'D:\\office\\各单位表十一\\数据源\\43-HR-04 岗位变更通知单_1.xls'  # 全部
    zhengchang_li_path = 'D:\\office\\各单位表十一\\数据源\\员工离职审批表 （新）黄亚楠_1.xls'  # 全部
    yichang_li_path = 'D:\\office\\各单位表十一\\数据源\\异常离职员工OA账号注销申请表黄亚楠_1.xls'  # 全部
    zhuanz_path = 'D:\\office\\各单位表十一\\数据源\\员工转正申请表（新）黄亚楠_1.xls'  # 全部

    out_path = "D:\\office\\各单位表十一\\"
    try:
        start_time = time.time()
        data = pd.read_excel(path, '员工档案数据管理', skiprows=2)
        end_time = time.time()
        print("一张表，主表读取完毕,用时", end_time - start_time, "秒")
        # print(data)
        print(data.columns.values)

        print("1.2 正在读取第二张表，家庭成员表")
        start_time = time.time()
        data2 = pd.read_excel(path, 'group1', skiprows=1)
        end_time = time.time()
        print("第二张表，家庭成员表读取完毕,用时", end_time - start_time, "秒")
        # print(data2)
        # print(data2.columns.values)

        print("1.3 正在读取第三张表，社保办理表")
        start_time = time.time()
        data3 = pd.read_excel(path, 'group7', skiprows=1)
        end_time = time.time()
        print("第三张表，社保办理表表读取完毕,用时", end_time - start_time, "秒")
        # print(data3)
        # print(data3.columns.values)

        print("1.4 正在读取第四张表，合同签订表")
        start_time = time.time()
        data4 = pd.read_excel(path, 'group8', skiprows=1)
        end_time = time.time()
        print("第四张表，合同签订表表读取完毕,用时", end_time - start_time, "秒")
        # print(data4)
        # print(data4.columns.values)

        print("1.5 正在读取第五张表，商业险办理表")
        start_time = time.time()
        data5 = pd.read_excel(path, 'group9', skiprows=1)
        end_time = time.time()
        print("第五张表，商业险办理表表读取完毕,用时", end_time - start_time, "秒")
        # print(data5)
        # print(data5.columns.values)

        print("1.6 正在读取第六张表，毕业院校表")
        start_time = time.time()
        school = pd.read_excel(path, 'group2', skiprows=1)
        end_time = time.time()
        print("第六张表，毕业院校表读取完毕,用时", end_time - start_time, "秒")
        # print(data5)
        # print(data5.columns.values)


    except Exception as  result:
        print("档案表读取失败，程序结束运行")
        print(result)
        return

    try:
        print("1.6 正在读取第六张表，新员工入职申请表")
        start_time = time.time()
        data6 = pd.read_excel(new_ruzhi_path, '员工入职申请表（新新员工）黄亚楠', skiprows=2)
        end_time = time.time()
        print("第六张表，新员工入职申请表读取完毕,用时", end_time - start_time, "秒")
        # print(data6)
        print(data6.columns.values)
    except Exception as result:
        print("入职申请表读取失败")
        print(result)
        new_ruzhi_path = ''

    try:
        print("1.7 正在读取第七张表，新复职员工入职申请表")
        start_time = time.time()
        data7 = pd.read_excel(old_ruzhi_path, '员工入职申请表（新复职员工）黄亚楠', skiprows=2)
        end_time = time.time()
        print("第七张表，新复职员工入职申请表读取完毕,用时", end_time - start_time, "秒")
        # print(data6)
        print(data7.columns.values)
    except Exception as result:
        print("复职员工入职申请表读取失败")
        print(result)
        old_ruzhi_path = ''

    try:
        print("1.8 正在读取第八张表，岗位变更申请表")
        start_time = time.time()
        bian = pd.read_excel(bian_path, '43-HR-04 岗位变更通知单', skiprows=2)
        end_time = time.time()
        print("第八张表，新岗位变更申请表读取完毕,用时", end_time - start_time, "秒")
        # print(data6)
        print(bian.columns.values)
    except Exception as result:
        print("岗位变更申请表读取失败")
        print(result)
        bian_path = ''

    try:
        new_open = pd.read_excel(new_open_path, '52-HR-14 OA账号开通申请表', skiprows=2)
    except Exception as result:
        print("OA开通申请表读取失败")
        print(result)
        new_open_path = ''

    try:
        old_open = pd.read_excel(old_open_path, '53-HR-15 OA账号恢复申请表', skiprows=2)
    except Exception as result:
        print("OA恢复申请表读取失败")
        print(result)
        old_open_path = ''
    try:
        zhengchang_li = pd.read_excel(zhengchang_li_path, '员工离职审批表 （新）黄亚楠', skiprows=2)
    except Exception as result:
        print("员工离职申请表读取失败")
        print(result)
        zhengchang_li_path = ''

    try:
        yichang_li = pd.read_excel(yichang_li_path, '异常离职员工OA账号注销申请表黄亚楠', skiprows=2)
    except Exception as result:
        print("员工离职申请表读取失败")
        print(result)
        yichang_li_path = ''

    try:
        zhuanz = pd.read_excel(zhuanz_path, '员工转正申请表（新）黄亚楠', skiprows=2)
        print(zhuanz.columns.values)
    except Exception as result:
        print("员工转正申请表读取失败")
        print(result)
        zhuanz_path = ''

    # 测试语句
    # print("出生日期的数据类型")
    # data["出生日期"] = pd.to_datetime(data["出生日期"], errors='coerce')
    # print(data["出生日期"].dtype)

    '''
    OA新入职写入，已存在则更新异动，否则添加一整行
    OA复制员工入职写入，只更新异动为入职
    OA新复开通，档案已存在只更新异动为入职
    转正
    职级变更，更新异动所有信息
    晋升、变更、
    离职、异常离职
    '''

    print("2.从档案表复制主表基本信息")
    # 得到 未经改列明的数据
    # 待生成字段，家庭联系方式,户口地，紧急联系人/关系/电话，用工性质，人员异动， 调动/离职日期,原调入公司,原调入部门原调入岗位
    # 调出/晋升公司,调出/晋升部门,调出/晋升岗位,异动/离职类型,异动/离职类型,调动原因,离职原因,
    # 合同续签...,档案/合同存放地，劳动合同移交个人（是/否）,是否购买商业险,商业险购买日期,保险机构,保险期限,是否缴纳五险,参保时间
    # 无法生成字段：健康状况

    print("3.生成合成字段")
    # data["紧急联系人/关系/电话"] = data["紧急联系人"].astype('str') + "/" + data["联系人关系"].astype('str') + "/" + data["联系人电话"].astype('str').str[0:-2]

    data["户口地"] = data["户口省"].astype('str') + data["户口市"].astype('str')

    # 遍历家庭表，给主表加一个字段  比较耗时  最后执行
    # data["家庭联系方式"] = ''

    data["人员异动"] = "在职"

    data.rename(columns={'离职类型': '异动/离职类型'}, inplace=True)

    '''
    更新档案内的数据
    入职日期为本月的 把人员异动更新为入职
    岗位变更日期为本月的，本单位为 内部调动 否则看情况给 是调入还是调出
    离职
    '''

    now_year = datetime.datetime.now().year
    now_month = datetime.datetime.now().month
    now_day = datetime.datetime.now().day

    data["离职日期"] = pd.to_datetime(data["离职日期"], errors='coerce').dt.normalize()
    # print(data["离职日期"])

    data["入职日期"] = pd.to_datetime(data["入职日期"], errors='coerce').dt.normalize()
    # data["入职日期"] = data["入职日期"].dt.date

    data['year'] = data['入职日期'].dt.year.fillna(0).astype("int")  # 转化提取年 ,
    # 如果有NaN元素则默认转化float64型，要转换数据类型则需要先填充空值,在做数据类型转换
    data['month'] = data['入职日期'].dt.month.fillna(0).astype("int")  # 转化提取月
    data['day'] = data['入职日期'].dt.day.fillna(0).astype("int")

    print("======================================")

    print(data['year'].dtype)

    print(data[(data["year"] == now_year) & (data["month"] == now_month)].index.tolist())
    for value in data[(data["year"] == now_year) & (data["month"] == now_month)].index.tolist():
        # data["人员异动"][value] = "入职"
        data.loc[value, "人员异动"] = "入职"

    print("档案中当月入职人员信息已处理")
    print(data[(data["year"] == now_year) & (data["month"] == now_month)])

    # 处理离职信息  只能从流程区分是主动离职还是辞退
    # 变更也是从流程来吧

    # 从写入异动数据
    '''
    新员工入职 只导出流转中的，追加到表十一
    然后找出所有入职日期为本月的，更新人员异动 为入职
    '''
    old_open["复职日期"] = pd.to_datetime(old_open["复职日期"], errors='coerce')
    old_open['year'] = old_open['复职日期'].dt.year.fillna(0).astype("int")  # 转化提取年 ,
    # 如果有NaN元素则默认转化float64型，要转换数据类型则需要先填充空值,在做数据类型转换
    old_open['month'] = old_open['复职日期'].dt.month.fillna(0).astype("int")  # 转化提取月
    old_open['day'] = old_open['复职日期'].dt.day.fillna(0).astype("int")

    # OA恢复
    if old_open_path != '':
        # 遍历OA恢复，查看入职日期是否一直，不一致则更新
        for index, row in old_open.iterrows():
            # 获取当前人员编号和入职日期
            bianhao = row["原人员编号"]
            ruzhi_date = str(row["复职日期"])
            # print(data[data["人员编号"] == bianhao].index.tolist()[0])
            dang_bianhao = data["人员编号"][(data[data["人员编号"] == bianhao]["人员编号"].index.tolist()[0])]
            dang_ruzhi_date = str(data["入职日期"][(data[data["人员编号"] == bianhao]["入职日期"].index.tolist()[0])])

            if ruzhi_date != dang_ruzhi_date:
                # 发了OA开通还没有发复职入职，需要更新档案　员工状态、入职日期、
                print("OA恢复表人员编号", bianhao)
                print("OA恢复表入职日期", ruzhi_date)
                # print("档案表人员编号", dang_bianhao)
                print("档案表入职日期", dang_ruzhi_date)
                print("需要更新档案信息")
                index2 = data[data["人员编号"] == bianhao].index.tolist()[0]
                data.loc[index2, "入职日期"] = row["复职日期"]
                data.loc[index2, "人员异动"] = "入职"
                data.loc[index2, "员工状态"] = "试用期"
                data.loc[index2, "工龄年"] = 0
                data.loc[index2, "现职单位"] = row["单位"]
                data.loc[index2, "现职部门"] = row["计划部门"]
                data.loc[index2, "现职岗位"] = row["计划岗位"]
                data.loc[index2, "现职级"] = row["职务级别"]
                data.loc[index2, "手机号码"] = row["手机号码"]
                data.loc[index2, "离职日期"] = ''

                # data[data["人员编号"] == bianhao]["入职日期"] = row["复职日期"]
                # data[data["人员编号"] == bianhao]["用工性质"] = "合同"
                # data[data["人员编号"] == bianhao]["人员异动"] = "入职"
                # data[data["人员编号"] == bianhao]["员工状态"] = "试用期"
                # data[data["人员编号"] == bianhao]["工龄年"] = 0
                # data[data["人员编号"] == bianhao]["现职单位"] = row["单位"]
                # data[data["人员编号"] == bianhao]["现职部门"] = row["计划部门"]
                # data[data["人员编号"] == bianhao]["现职岗位"] = row["计划岗位"]
                # data[data["人员编号"] == bianhao]["现职级"] = row["职务级别"]
                # data[data["人员编号"] == bianhao]["手机号码"] = row["手机号码"]

                print(data.loc[index2, :])
        pass

    # TODO 去重,多个联系方式会导致多行 新员工入职信息
    # 新员工入职
    if new_ruzhi_path != '':
        # 拼接家庭联系方式和紧急联系方式
        # data6["家庭联系方式"] = data6["家庭成员姓名"].astype('str') + "/" + data6["家庭成员关系"].astype('str') + "/" + data6[
        #     "家庭成员联系方式"].astype('str')
        # data6["紧急联系人/关系/电话"] = data6["紧急联系人"].astype('str') + "/" + data6["联系人关系"].astype('str') + "/" + data6[
        #     "联系人电话"].astype('str')
        data6["户口地"] = data6["户口省"].astype('str') + data6["户口市"].astype('str')
        data6["人员异动"] = "入职"
        data6["工龄年"] = 0

        print(data6)
        new_ruzhi_data_need_add = data6.loc[:,
                                  ["单位", "部门", "岗位", "职级", "员工状态", "姓名", "出生日期", "民族", "学历", "身份证号", "手机号码",
                                   "政治面貌", "婚姻状况", "身份证住址", "现居住地", "户口地", "入职日期", "招聘来源", "人员编号", "年龄",
                                   "工龄年", "性别", "社保状态", "人员异动", "紧急联系人", "联系人关系","联系人电话"]]

        # 更改列名和档案表保持一致
        new_ruzhi_data_need_add.rename(columns={'单位': '现职单位', '部门': '现职部门', '岗位': '现职岗位', '职级': '现职级'}, inplace=True)
        print("流转中的新员工")
        print(new_ruzhi_data_need_add)
        new_ruzhi_data_need_add = new_ruzhi_data_need_add.drop_duplicates(subset=['人员编号'], keep='first')

        # 追加到档案表
        data = data.append(new_ruzhi_data_need_add, ignore_index=True)
        print(data)
        # print(data[data["姓名"]=='尚泽斌'])

    # 复职入职信息
    if old_ruzhi_path != '':
        # 拼接家庭联系方式和紧急联系方式
        data7["家庭联系方式"] = data7["家庭成员姓名"].astype('str') + "/" + data7["家庭成员关系"].astype('str') + "/" + data7[
            "家庭成员联系方式"].astype('str')
        data7["紧急联系人/关系/电话"] = data7["紧急联系人"].astype('str') + "/" + data7["联系人关系"].astype('str') + "/" + data7[
            "联系人电话"].astype('str')
        data7["户口地"] = data7["户口省"].astype('str') + data7["户口市"].astype('str')
        data7["人员异动"] = "入职"
        data7["工龄年"] = 0
        #
        data7["离职日期"] = ''

        print(data7)
        old_ruzhi_data_need_add = data7.loc[:,
                                  ["单位", "部门", "岗位", "职级", "员工状态", "姓名", "出生日期", "民族", "学历", "身份证号", "手机号码",
                                   "政治面貌", "婚姻状况", "身份证住址", "现居住地", "户口地", "入职日期", "人员编号", "年龄",
                                   "工龄年", "性别", "社保状态", "人员异动", "紧急联系人", "联系人关系","联系人电话", "离职日期"]]

        # 更改列名和档案表保持一致
        old_ruzhi_data_need_add.rename(columns={'单位': '现职单位', '部门': '现职部门', '岗位': '现职岗位', '职级': '现职级'}, inplace=True)
        print(old_ruzhi_data_need_add)
        old_ruzhi_data_need_add = old_ruzhi_data_need_add.drop_duplicates(subset=['人员编号'], keep='first')
        old_ruzhi_bianhao = old_ruzhi_data_need_add["人员编号"].tolist()

        update_lie = ["现职部门", "现职岗位", "现职级", "员工状态", "姓名", "出生日期", "民族", "学历", "身份证号", "手机号码",
                      "政治面貌", "婚姻状况", "身份证住址", "现居住地", "户口地", "入职日期", "人员编号", "年龄",
                      "工龄年", "性别", "社保状态", "人员异动", "紧急联系人", "联系人关系","联系人电话", "离职日期"]

        for index in old_ruzhi_bianhao:
            print(index)
            value = old_ruzhi_data_need_add[old_ruzhi_data_need_add["人员编号"] == index]
            dao_old_index = old_ruzhi_data_need_add[old_ruzhi_data_need_add["人员编号"] == index].index.tolist()[0]
            data_old_index = data[data["人员编号"] == index].index.tolist()[0]
            # print(value)
            for up_name in update_lie:
                # data[up_name][data_old_index] = value[up_name][dao_old_index]
                data.loc[data_old_index, up_name] = value[up_name][dao_old_index]

        print("流转中流程复职员工信息已更新至档案")
        # print(data[data["姓名"] == "陈灵灵"]["工龄月"])

    # OA开通写入
    if new_open_path != '':
        # 遍历oa开通表，如果编号已经存在于档案表则不执行任何操作，否则 写入到表单中
        for index, row in new_open.iterrows():
            bianhao = row["人员编号"]
            dang_bianhao_arr = data["人员编号"].tolist()
            if bianhao not in dang_bianhao_arr:
                print("人员编号", bianhao, "不存在档案表中，需添加到档案")
                new_open_need_add_arr = new_open.loc[:,
                                        ["单位","计划部门", "计划岗位", "职务级别", "姓名", "性别", "手机号码", "身份证号", "人员编号", "入职日期"]]
                new_open_need_add_arr.rename(columns={'单位':'现职单位','计划部门': '现职部门', '计划岗位': '现职岗位', '职务级别': '现职级'}, inplace=True)
                new_open_need_add_arr["人员异动"] = "入职"
                new_open_need_add_arr["员工状态"] = "试用期"
                new_open_need_add_arr["工龄年"] = 0
                new_open_need_add_arr = new_open_need_add_arr[new_open_need_add_arr["人员编号"] == bianhao]
                data = data.append(new_open_need_add_arr, ignore_index=True)
                print(data[data["人员编号"] == bianhao])
            else:
                print("人员编号", bianhao, "已存在档案表中，无需添加到档案")
    # print(data)
    # print(data[data["人员编号"] == 'RH200801001'])

    # 转正
    if zhuanz_path != '':
        for index, row in zhuanz.iterrows():
            bianhao = row["人员编号"]
            index2 = data[data["人员编号"] == bianhao].index.tolist()[0]
            data.loc[index2, "转正日期"] = row["转正时间"]
            data.loc[index2, "人员异动"] = "转正"
            data.loc[index2, "员工状态"] = "正式"

    # 岗位变更写入到档案
    if bian_path != '':
        '''
        获取流程结束的和流转中的
        已结束的：
            看调动是本单位还是跨单位
            本单位的:
                更新人员异动、异动/离职类型 为内部调动，调动前后单位信息更新，现职单位等信息无需更新
            跨单位的: 
                调入方：
                    更新人员异动、异动/离职类型 为内部调动，调动前后单位信息更新，现职单位等信息无需更新
                调出方：
                    复制一份调入方信息,流程中的原单位信息更新进去
        流转中的：
            本单位的：
                更新人员异动、异动/离职类型 为内部调动，调动前后单位信息更新，现职单位等信息更新
            跨单位的：
                调入方：
                    获取人员编号，获取信息复制一行，更新人员异动、异动/离职类型 为调入，现职单位信息更新，
        '''
        # 获取已结束的
        bian_tong_end_id_arr = bian[(bian["流程状态"] == "已结束") & (bian["原单位"] == bian["调动后单位"])]["人员编号"].tolist()
        bian_yi_end_id_arr = bian[(bian["流程状态"] == "已结束") & (bian["原单位"] != bian["调动后单位"])]["人员编号"].tolist()

        bian_tong_ing_id_arr = bian[(bian["流程状态"] == "未结束") & (bian["原单位"] == bian["调动后单位"])]["人员编号"].tolist()
        bian_yi_ing_id_arr = bian[(bian["流程状态"] == "未结束") & (bian["原单位"] != bian["调动后单位"])]["人员编号"].tolist()

        print("已结束，内部调动")
        print(bian_tong_end_id_arr)
        print("已结束，跨单位调")
        print(bian_yi_end_id_arr)
        print("未结束，内部调动")
        print(bian_tong_ing_id_arr)
        print("未结束，跨单位调")
        print(bian_yi_ing_id_arr)

        # for bianhao in bian_tong_end_id_arr:
        #     index = data[data["人员编号"] == bianhao].index.tolist()
        #     data.loc[index,]

        data["调动/离职日期"] = ''
        data["原调入公司"] = ''
        data["原调入部门"] = ''
        data["原调入岗位"] = ''
        data["调出/晋升公司"] = ''
        data["调出/晋升部门"] = ''
        data["调出/晋升岗位"] = ''
        # data["异动/离职类型"] = ''
        data["调动原因"] = ''
        data["离职原因"] = ''

        if len(bian_tong_end_id_arr) != 0:
            # 更新人员异动、异动/离职类型 为内部调动，调动前后单位信息更新，现职单位等信息无需更新
            print(data.columns.values)
            # 遍历已结束内部调动的人
            for bianhao in bian_tong_end_id_arr:
                bian_index = bian[bian["人员编号"] == bianhao].index.tolist()[0]
                dang_index = data[data["人员编号"] == bianhao].index.tolist()[0]
                data.loc[dang_index, '人员异动'] = "调动"
                data.loc[dang_index, '调动/离职日期'] = bian.loc[bian_index, "调动日期"]
                data.loc[dang_index, '原调入公司'] = bian.loc[bian_index, "原单位"]
                data.loc[dang_index, '原调入部门'] = bian.loc[bian_index, "原部门"]
                data.loc[dang_index, '原调入岗位'] = bian.loc[bian_index, "原岗位"]
                data.loc[dang_index, '调出/晋升公司'] = bian.loc[bian_index, "调动后单位"]
                data.loc[dang_index, '调出/晋升部门'] = bian.loc[bian_index, "调动后部门"]
                data.loc[dang_index, '调出/晋升岗位'] = bian.loc[bian_index, "调动后岗位"]
                data.loc[dang_index, '调动原因'] = bian.loc[bian_index, "调动原因"]
                data.loc[dang_index, '异动/离职类型'] = "内部调动"

        if len(bian_yi_end_id_arr) != 0:
            # 遍历已结束跨单位调动的人
            for bianhao in bian_yi_end_id_arr:
                # 获取变更行
                bian_index = bian[bian["人员编号"] == bianhao].index.tolist()[0]
                bian_row = bian.loc[bian_index, :]
                # 获取档案行
                dang_index = data[data["人员编号"] == bianhao].index.tolist()[0]

                # 操作档案中已存在的调入的
                data.loc[dang_index, '人员异动'] = "调动"
                data.loc[dang_index, '调动/离职日期'] = bian.loc[bian_index, "调动日期"]
                data.loc[dang_index, '原调入公司'] = bian.loc[bian_index, "原单位"]
                data.loc[dang_index, '原调入部门'] = bian.loc[bian_index, "原部门"]
                data.loc[dang_index, '原调入岗位'] = bian.loc[bian_index, "原岗位"]
                data.loc[dang_index, '调出/晋升公司'] = bian.loc[bian_index, "调动后单位"]
                data.loc[dang_index, '调出/晋升部门'] = bian.loc[bian_index, "调动后部门"]
                data.loc[dang_index, '调出/晋升岗位'] = bian.loc[bian_index, "调动后岗位"]
                data.loc[dang_index, '调动原因'] = bian.loc[bian_index, "调动原因"]
                data.loc[dang_index, '异动/离职类型'] = "调入"

                dang_row = data.loc[dang_index, :]

                # 档案表复制一行,把原单位数据写进去
                data = data.append(dang_row, ignore_index=True)
                data_last_index = data.index.tolist()[len(data.index.tolist()) - 1]
                data.loc[data_last_index, "现职单位"] = bian.loc[bian_index, "原单位"]
                data.loc[data_last_index, '现职部门'] = bian.loc[bian_index, "原部门"]
                data.loc[data_last_index, '现职岗位'] = bian.loc[bian_index, "原岗位"]
                data.loc[data_last_index, '现职级'] = bian.loc[bian_index, "原职级"]
                data.loc[data_last_index, '异动/离职类型'] = "调出"

            # print("最后5行")
            # print(data.tail())
            #
            # print(data_last_index)

        if len(bian_tong_ing_id_arr) != 0:
            for bianhao in bian_yi_ing_id_arr:
                bian_index = bian[bian["人员编号"] == bianhao].index.tolist()[0]
                dang_index = data[data["人员编号"] == bianhao].index.tolist()[0]

                data.loc[dang_index, '人员异动'] = "调动"
                data.loc[dang_index, '调动/离职日期'] = bian.loc[bian_index, "调动日期"]
                data.loc[dang_index, '原调入公司'] = bian.loc[bian_index, "原单位"]
                data.loc[dang_index, '原调入部门'] = bian.loc[bian_index, "原部门"]
                data.loc[dang_index, '原调入岗位'] = bian.loc[bian_index, "原岗位"]
                data.loc[dang_index, '调出/晋升公司'] = bian.loc[bian_index, "调动后单位"]
                data.loc[dang_index, '调出/晋升部门'] = bian.loc[bian_index, "调动后部门"]
                data.loc[dang_index, '调出/晋升岗位'] = bian.loc[bian_index, "调动后岗位"]
                data.loc[dang_index, '调动原因'] = bian.loc[bian_index, "调动原因"]
                data.loc[dang_index, '异动/离职类型'] = "内部调动"
                data.loc[dang_index, '现职级'] = bian.loc[bian_index, "调动后职级"]
                data.loc[dang_index, '员工状态'] = bian.loc[bian_index, "调动后员工状态"]

        if len(bian_yi_ing_id_arr) != 0:
            # 未节航速
            for bianhao in bian_yi_ing_id_arr:
                bian_index = bian[bian["人员编号"] == bianhao].index.tolist()[0]
                dang_index = data[data["人员编号"] == bianhao].index.tolist()[0]

                # 操作档案中已存在的调出的
                data.loc[dang_index, '人员异动'] = "调动"
                data.loc[dang_index, '调动/离职日期'] = bian.loc[bian_index, "调动日期"]
                data.loc[dang_index, '原调入公司'] = bian.loc[bian_index, "原单位"]
                data.loc[dang_index, '原调入部门'] = bian.loc[bian_index, "原部门"]
                data.loc[dang_index, '原调入岗位'] = bian.loc[bian_index, "原岗位"]
                data.loc[dang_index, '调出/晋升公司'] = bian.loc[bian_index, "调动后单位"]
                data.loc[dang_index, '调出/晋升部门'] = bian.loc[bian_index, "调动后部门"]
                data.loc[dang_index, '调出/晋升岗位'] = bian.loc[bian_index, "调动后岗位"]
                data.loc[dang_index, '调动原因'] = bian.loc[bian_index, "调动原因"]
                data.loc[dang_index, '异动/离职类型'] = "调出"

                dang_row = data.loc[dang_index, :]

                # 档案表复制一行,把原单位数据写进去
                data = data.append(dang_row, ignore_index=True)
                data_last_index = data.index.tolist()[len(data.index.tolist()) - 1]
                data.loc[data_last_index, "现职单位"] = bian.loc[bian_index, "调动后单位"]
                data.loc[data_last_index, '现职部门'] = bian.loc[bian_index, "调动后部门"]
                data.loc[data_last_index, '现职岗位'] = bian.loc[bian_index, "调动后岗位"]
                data.loc[data_last_index, '现职级'] = bian.loc[bian_index, "调动后职级"]
                data.loc[data_last_index, '员工状态'] = bian.loc[bian_index, "调动后员工状态"]

                data.loc[data_last_index, '异动/离职类型'] = "调入"

                pass

    # 正常离职
    if zhengchang_li_path != '':
        for index, row in zhengchang_li.iterrows():
            bianhao = row["人员编号"]
            dang_index_arr = data[data["人员编号"] == bianhao].index.tolist()
            for dang_index in dang_index_arr:
                data.loc[dang_index, '人员异动'] = "离职"
                data.loc[dang_index, '调动/离职日期'] = row["实际离职日期"]
                data.loc[dang_index, "员工状态"] = "离职"
                data.loc[dang_index, "异动/离职类型"] = row["离职类型"]
                data.loc[dang_index, "离职原因"] = row["任职性质"]
    # 异常离职
    if yichang_li_path != '':
        for index, row in yichang_li.iterrows():
            bianhao = row["人员编号"]
            dang_index_arr = data[data["人员编号"] == bianhao].index.tolist()
            for dang_index in dang_index_arr:
                data.loc[dang_index, '人员异动'] = "离职"
                data.loc[dang_index, '调动/离职日期'] = row["离职日期"]
                data.loc[dang_index, "员工状态"] = "离职"
                data.loc[dang_index, "异动/离职类型"] = row["离职类型"]
                data.loc[dang_index, "离职原因"] = row["任职性质"]

    # 帅选出非离职的和一个月内离职的

    print(data.columns.values)
    print(data["离职日期"])
    # TODO 复制和OA恢复的地方　把离职这仨值设置成0，新入职同理
    data['liyear'] = data['离职日期'].dt.year.fillna(0).astype("int")  # 转化提取年 ,
    # 如果有NaN元素则默认转化float64型，要转换数据类型则需要先填充空值,在做数据类型转换
    data['limonth'] = data['离职日期'].dt.month.fillna(0).astype("int")  # 转化提取月
    data['liday'] = data['离职日期'].dt.day.fillna(0).astype("int")

    zaizhi_emp_arr = data[data["员工状态"] != "离职"]
    lizhi_emp_arr = data[(data["员工状态"] == "离职") & (data['liyear'] == now_year) & (data['limonth'] == now_month)]
    data = zaizhi_emp_arr.append(lizhi_emp_arr, ignore_index=True)
    print(data)

    # print(data)

    # for index, row in data.iterrows():
    #     bianhao1 = row["$"]
    #     for index2, row2 in data2.iterrows():
    #         bianhao2 = row2["$"]
    #         if bianhao1 == bianhao2:
    #             # print(row2["家庭成员联系方式"])
    #             # data["家庭联系方式"][index] = data2["家庭成员姓名"].astype('str')[index2] + "/" + data2["家庭成员关系"].astype('str')[index2] + "/" + data2["家庭成员联系方式"].astype('str')[index2]
    #             # print(data["家庭联系方式"][index])
    #             data.loc[index, "家庭联系方式"] = data2["家庭成员姓名"].astype('str')[index2] + "/" + data2["家庭成员关系"].astype('str')[
    #                 index2] + "/" + data2["家庭成员联系方式"].astype('str')[index2]
    #             # print(data.loc[index,["家庭联系方式","家庭成员姓名","家庭成员关系","家庭成员联系方式"]])
    #             print(data["家庭联系方式"][index])
    #             break

    data["合同签订记录"] = ''
    data["档案/合同存放地"] = ''

    # 生成合同签订记录，

    for index, row in data.iterrows():
        bianhao1 = row["$"]
        i = 1
        for index2, row2 in data4.iterrows():
            bianhao2 = row2["$"]
            if bianhao1 == bianhao2:
                if i == 1:
                    data.loc[index, "合同签订记录"] = str(data.loc[index, "合同签订记录"]) + str(row2["合同签订日期"]) + "," + str(
                        row2["合同到期日期"]) + ","
                    i = i + 1
                else:
                    data.loc[index, "合同签订记录"] = str(data.loc[index, "合同签订记录"]) + str(row2["合同到期日期"]) + ","
                # 合同存放地只取最新的
                data.loc[index, "档案/合同存放地"] = row2["合同存放地"]
        # TODO TMLGBZ 最后一个逗号最后去调
        print(data.loc[index, "合同签订记录"])
    print(data)

    data5.drop_duplicates(subset=["$"], keep='last', inplace=True)
    print(data5)

    # 生成商业险
    data["是否购买商业险"] = '否'
    data["保险期限"] = ""

    data5["保险到期日期"] = pd.to_datetime(data5["保险到期日期"], errors='coerce').dt.normalize()
    data5['xyear'] = data5['保险到期日期'].dt.year.fillna(0).astype("int")  # 转化提取年 ,
    data5['xmonth'] = data5['保险到期日期'].dt.month.fillna(0).astype("int")  # 转化提取月
    data5['xday'] = data5['保险到期日期'].dt.day.fillna(0).astype("int")  # 转化为日

    dtime = datetime.datetime.now()
    now_time_num = int(time.mktime(dtime.timetuple()))

    data["商业险购买日期"] = ""
    data["保险机构"] = ""
    data["缴纳险种"] = ""
    for index, row in data.iterrows():
        bianhao1 = row["$"]
        for index2, row2 in data5.iterrows():
            bianhao2 = row2["$"]
            if bianhao1 == bianhao2:
                # data.loc[index,"是否购买商业险"]
                data.loc[index, "商业险购买日期"] = row2["保险生效日期"]
                data.loc[index, "保险期限"] = row2["保险到期日期"]
                data.loc[index, "保险机构"] = row2["保险公司"]
                data.loc[index, "缴纳险种"] = row2["险种"]
                # 保险日期没到就把是否购买商业险改成是
                t1 = 1
                try:
                    d1 = datetime.date(row2["xyear"], row2["xmonth"], row2["xday"])
                    t1 = int(time.mktime(d1.timetuple()))
                except Exception as result:
                    pass
                if t1 > now_time_num:
                    data.loc[index, "是否购买商业险"] = "是"
                pass
        print(data.loc[index, :])

    data["是否缴纳五险"] = ""
    data["参保时间"] = ""
    data3.dropna(subset=["办理单位"], inplace=True)
    print(data3)
    for index, row in data3.iterrows():
        bianhao1 = row["$"]
        for index2, row2 in data.iterrows():
            bianhao2 = row2["$"]
            if bianhao1 == bianhao2:
                if row["办理类型"] == "新增" or row["办理类型"] == "转移" or row["办理类型"] == "续保":
                    data.loc[index2, "是否缴纳五险"] = "是"
                    data.loc[index2, "参保时间"] = row["办理日期"]
                elif row["办理类型"] == "停保":
                    data.loc[index2, "是否缴纳五险"] = ""
                    data.loc[index2, "参保时间"] = ""
                print(data.loc[index2, :])

    for index, row in data.iterrows():
        canbao_status = row["社保状态"]
        if "已参保" == str(canbao_status):
            data.loc[index,"是否缴纳五险"] = "是"

    data["用工性质"] = "合同"
    # data["档案/合同存放地"] = "否"

    # 处理毕业院校
    # 去空白，去重保留最后一行
    school.dropna(subset=["毕业院校"], inplace=True)
    school.drop_duplicates(subset=["$"], keep='last', inplace=True)
    data["毕业院校"] = ""
    for index, row in school.iterrows():
        bianhao1 = row["$"]
        for index2, row2 in data.iterrows():
            bianhao2 = row2["$"]
            if bianhao1 == bianhao2:
                data.loc[index2, "毕业院校"] = row["毕业院校"]
                print(data.loc[index2, :])

    print("信息处理后的档案所有字段======================================")
    print(data.columns.values)
    print(data)


    # 删除所有调出的行
    data.drop(data[data['异动/离职类型'] == "调出"].index, inplace=True)


    # TODO 如有多个人员编号，进行去重  按照单位分割文件  可以最后分割

    # data.dropna(subset=["现职单位"], inplace=True)
    # data['序号'] = range(1, len(data) + 1)

    col_list = ["现职单位", "现职部门", "现职岗位", "现职级", "员工状态", "姓名", "出生日期", "民族", "学历", "身份证号", "手机号码",
                "政治面貌", "婚姻状况","有无子女", "身份证住址", "户口地", "现居住地", "紧急联系人", "联系人关系","联系人电话",
                "用工性质", "人员异动", "入职日期", "招聘来源", "转正日期", "调动/离职日期", "原调入公司", "原调入部门",
                "原调入岗位", "调出/晋升公司", "调出/晋升部门", "调出/晋升岗位", "异动/离职类型", "调动原因", "离职原因",
                "人员编号", "合同签订记录", "档案/合同存放地", "是否购买商业险", "保险机构", "缴纳险种", "保险期限", "是否缴纳五险",
                "参保时间", "年龄", "工龄年", "性别", "毕业院校"]
    col_list_out = ["序号","现职单位","现职部门", "现职岗位", "现职级", "员工状态", "姓名", "出生日期", "民族", "学历", "身份证号", "手机号码",
                "政治面貌", "婚姻状况","有无子女","身份证住址", "户口地", "现居住地", "紧急联系人", "联系人关系","联系人电话",
                "用工性质", "人员异动", "入职日期", "招聘来源", "转正日期", "调动/离职日期", "原调入公司", "原调入部门",
                "原调入岗位", "调出/晋升公司", "调出/晋升部门", "调出/晋升岗位", "异动/离职类型", "调动原因", "离职原因",
                "人员编号", "合同签订记录", "档案/合同存放地", "是否购买商业险", "保险机构", "缴纳险种", "保险期限", "是否缴纳五险",
                "参保时间", "年龄", "工龄年", "性别", "毕业院校"]
    data = data[col_list]

    #data.sort_values(na_position='first', inplace=True)  # 空值在前
    #data.sort_values(by=['现职单位','现职部门'], inplace=True)

    data.to_excel(out_path + "全集团人事月报.xlsx", sheet_name="sheet1", index=False)

    listType = data['现职单位'].unique()
    print(listType)

    count = 0
    for index, sheet in enumerate(listType):
        # 写入到指定sheet
        if pd.isnull(sheet):
            continue
        count = count + 1
        table = data[data["现职单位"] == sheet].copy()
        table.drop_duplicates(subset=["人员编号"], keep='first', inplace=True)
        table['序号'] = range(1, len(table) + 1)
        table = table[col_list_out]
        print("sheet == ", sheet,type(sheet))
        print("打印table")
        print(table)
        # 将table写入到指定文件夹
        file_name = str(count) + str(sheet) + '.xlsx'
        table.to_excel(out_path + file_name, sheet_name="附表十一人事信息库", index=False)
        pass

    end_time_main = time.time()
    print("表格生成结束，用时", end_time_main - start_time_main, "秒")

    pass


if __name__ == '__main__':
    start()
