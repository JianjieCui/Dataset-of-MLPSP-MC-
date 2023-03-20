import pandas as pd
import random
import copy
import numpy as np
import math


def main(file_name, PMachine_Nums, AMachine_Nums, Product_Kinds, StandardPart_Kinds, CustomPart_Max, Operations_Max,
         PTime_Max,
         PTime_Min, ATime_Min, ATime_Max, CPbatch_Max, CPbatch_Min, SPbatch_Max, SPbatch_Min, DT,
         Earliness_Min, Earliness_Max, Tardiness_Min, Tardiness_Max):
    # 存入DataSize-excel列表的内容
    DataSize_index = [1]
    DataSize_columns = ['加工机器数', '组装机器数', '产品种类数', '通用件种类数']  # 标题
    DataSize_Data = [[PMachine_Nums, AMachine_Nums, Product_Kinds, StandardPart_Kinds]]  # 数据

    # 加工机器信息：机器编号
    PMachine_Id_list = [i + 1 for i in range(PMachine_Nums)]  # 加工机器编号

    '''----------------product information--------------------'''
    Product_Id_list = [i + 1 for i in range(Product_Kinds)]  # 产品编号
    Product_Boom = []  # 存放每种产品包含的定制部件
    Custompart_NO = []  # 存放定制部件编号
    Part_NO = 0  # 定制部件编号
    for i in range(Product_Kinds):
        Product_Boom.append([])
        CustomPart_Kinds = int(random.uniform(1, CustomPart_Max + 1))  # 包含的定制部件种类
        for j in range(CustomPart_Kinds):
            Part_NO += 1
            Product_Boom[i].append(Part_NO)
            Custompart_NO.append(Part_NO)

    for i in range(Product_Kinds):  # 用-1补齐不足CustomPart_Max的产品结构信息
        if len(Product_Boom[i]) < CustomPart_Max:
            add_list = [-1] * (CustomPart_Max - len(Product_Boom[i]))  # 补齐需要的个数
            Product_Boom[i].extend(add_list)
    # 单位组装时间
    for i in range(Product_Kinds):  # 单位定制产品的组装时间
        time = int(random.uniform(ATime_Min, ATime_Max + 1))
        Product_Boom[i].append(time)  # 把组装时间存到Product_Boom列表末尾
    # 生产数量
    for i in range(Product_Kinds):  # 定制产品的生产数量
        Product_Num = int(random.uniform(CPbatch_Min, CPbatch_Max + 1))
        Product_Boom[i].append(Product_Num)

    # store
    ProductInfo_index = Product_Id_list.copy()  # 行索引（定制产品编号
    ProductInfo_index_name = '定制产品编号'
    ProductInfo_columns = ['定制部件编号' for i in range(CustomPart_Max)]  # 列标题
    ProductInfo_columns.append('组装时间')
    ProductInfo_columns.append('生产数量')
    ProductInfo_Data = Product_Boom.copy()  # 数据

    '''----------------Process information for custom parts--------------------'''
    # 定制部件加工路线
    Custompart_Tlist = []  # 定制部件各工序的加工时间
    operation_nums_list = []
    while operation_nums_list == [] or max(operation_nums_list) < Operations_Max:
        Custompart_Tlist = []  # 定制部件各工序的加工时间
        for i in range(len(Custompart_NO)):
            Custompart_Tlist.append([])
            operation_nums = int(random.uniform(1, Operations_Max + 1))  # 工序数
            operation_nums_list.append(operation_nums)
            for j in range(operation_nums):
                operation_machines = int(random.uniform(2, PMachine_Nums + 1))  # 每道工序的可用机器数
                process_machine_list = random.sample(PMachine_Id_list, operation_machines)  # 当前工序的可用机器
                process_time_list = []  # 在各机器上的加工时间
                for t in range(PMachine_Nums):
                    if PMachine_Id_list[t] in process_machine_list:
                        process_time = int(random.uniform(PTime_Min, PTime_Max + 1))  # 加工时间
                        process_time_list.append(process_time)
                    else:  # 如果产品不在该机器上加工，则存入1000（非法值
                        process_time_list.append(1000)
                Custompart_Tlist[i].append(process_time_list)  # 工序在各机器上的加工时间

    # 存入Custompartinfo列表的内容
    Custompart_index = Custompart_NO.copy()  # 行索引（定制部件编号
    Custompart_index_name = '定制部件编号'
    Custompart_columns = []  # 存入列标题，内容为各工序的可用机器
    for i in range(Operations_Max):  # 工序数=最大工序数
        for j in range(PMachine_Nums):  # 可用机器数=加工机器总数
            Custompart_columns.append(PMachine_Id_list[j])
    Custompartinfo_Data_tem = copy.deepcopy(Custompart_Tlist)  # 子表Customproductinfo的数据
    # 用-1补全工序数不足最大工序数的定制部件加工路线，以及可用机器数不足加工机器数的工序
    for i in range(len(Custompart_Tlist)):
        if len(Custompart_Tlist[i]) < Operations_Max:
            m_list = [-1] * PMachine_Nums
            for j in range(Operations_Max - len(Custompart_Tlist[i])):
                Custompartinfo_Data_tem[i].append(m_list)
    # 存入excel时，每一行得是一个列表，不能分为多级
    Custompartinfo_Data = []
    for i in range(len(Custompartinfo_Data_tem)):
        Custompartinfo_Data.append([])
        for j in range(len(Custompartinfo_Data_tem[i])):
            Custompartinfo_Data[i].extend(Custompartinfo_Data_tem[i][j])

    '''-----------------------------------process information of Standard part---------------------------'''
    StandardPart_id_list = [i + 1 for i in range(StandardPart_Kinds)]  # 产品编号
    ## 标准产品加工路线
    StandardP_Tlist = []  # 标准产品工艺信息
    operation_nums_list = []
    while operation_nums_list == [] or max(operation_nums_list) < Operations_Max:
        StandardP_Tlist = []  # 定制部件各工序的加工时间
        for i in range(StandardPart_Kinds):
            StandardP_Tlist.append([])
            operation_nums = int(random.uniform(1, Operations_Max + 1))  # 工序数
            operation_nums_list.append(operation_nums)
            for j in range(operation_nums):
                operation_machines = int(random.uniform(2, PMachine_Nums + 1))  # 每道工序的可用机器数
                process_machine_list = random.sample(PMachine_Id_list, operation_machines)  # 当前工序的可用机器
                process_time_list = []  # 在各机器上的加工时间
                for t in range(PMachine_Nums):
                    if PMachine_Id_list[t] in process_machine_list:
                        process_time = int(random.uniform(PTime_Min, PTime_Max + 1))  # 加工时间
                        process_time_list.append(process_time)
                    else:  # 如果产品不在该机器上加工，则存入1000（非法值
                        process_time_list.append(1000)
                StandardP_Tlist[i].append(process_time_list)  # 工序在各机器上的加工时间

    # 存入SPartinfo列表的内容
    Standardpart_index = StandardPart_id_list.copy()  # 行索引（标准产品编号
    Standardpart_index_name = '标准产品编号'
    Standardpart_columns = []  # 存入列标题，内容为各工序的可用机器
    for i in range(Operations_Max):  # 工序数=最大工序数
        for j in range(PMachine_Nums):  # 可用机器数=加工机器总数
            Standardpart_columns.append(PMachine_Id_list[j])
    SPartinfo_Data_tem = copy.deepcopy(StandardP_Tlist)  # 子表SProductinfo的数据
    # 用-1补全工序数不足最大工序数的标准产品加工路线，以及可用机器数不足加工机器数的工序
    for i in range(len(StandardP_Tlist)):
        if len(StandardP_Tlist[i]) < Operations_Max:
            m_list = [-1] * PMachine_Nums
            for j in range(Operations_Max - len(StandardP_Tlist[i])):
                SPartinfo_Data_tem[i].append(m_list)
    SPartinfo_Data = []  # 子表SPartinfo的数据
    for i in range(len(SPartinfo_Data_tem)):
        SPartinfo_Data.append([])
        for j in range(len(SPartinfo_Data_tem[i])):
            SPartinfo_Data[i].extend(SPartinfo_Data_tem[i][j])
    # 各标准件生产数量
    for i in range(StandardPart_Kinds):
        SP_Num = int(random.uniform(SPbatch_Min, SPbatch_Max + 1))  # 标准产品生产数量
        SPartinfo_Data[i].append(SP_Num)  # 生产数量
    Standardpart_columns.append('生产数量')

    '''---------------------------due date and early/delay cost of product ------------------------------'''
    ProductInfo_columns.append('交货期')
    ProductInfo_columns.append('提前系数')
    ProductInfo_columns.append('延迟系数')
    for i in range(Product_Kinds):
        product_num = ProductInfo_Data[i][-1]  # 产品生产数量
        product_asstime = ProductInfo_Data[i][-2]  # 产品单位组装时间
        product_boom = ProductInfo_Data[i][:-2]  # 产品结构信息
        custompart_process_time = []  # 存放构成产品的定制部件的总加工时间
        for custompart_id in product_boom[:-1]:  # 最后一位是组装时间，不用遍历
            process_time = 0
            if custompart_id == -1:  # 若定制部件编号为-1，则说明后面的都是为补齐长度而生成的无用数据
                break
            custom_index = Custompart_NO.index(custompart_id)  # 定制部件索引
            for process_time_list in Custompart_Tlist[custom_index]:
                list_tem1 = []
                for time in process_time_list:
                    if time != 1000:
                        list_tem1.append(time)
                process_time += np.mean(list_tem1)
            custompart_process_time.append(process_time)
        total_process_time = max(custompart_process_time)  # 产品加工阶段的总处理时间=各定制部件总加工时间的最大值
        total_batch_time = (product_asstime + total_process_time) * product_num  # 总批量时间
        DT = random.uniform(1.2, 2)
        print(DT)
        order_due_date = math.ceil(total_batch_time * DT)  # 交货期（向上取整
        earliness_rate = round(random.uniform(Earliness_Min, Earliness_Max), 1)  # 提前期单位时间成本
        tardiness_rate = round(random.uniform(Tardiness_Min, Tardiness_Max), 1)  # 拖期单位时间成本
        ProductInfo_Data[i].append(order_due_date)
        ProductInfo_Data[i].append(earliness_rate)
        ProductInfo_Data[i].append(tardiness_rate)

    writer = pd.ExcelWriter("../" + file_name + '.xlsx')  # 存放位置
    # DataSize子表
    pd_look1 = pd.DataFrame(DataSize_Data, index=DataSize_index, columns=DataSize_columns)
    pd_look1.index.name = '1'
    pd_look1.to_excel(writer, sheet_name="DataSize")
    # ProductBoom子表
    pd_look3 = pd.DataFrame(ProductInfo_Data, index=ProductInfo_index, columns=ProductInfo_columns)
    pd_look3.index.name = ProductInfo_index_name
    pd_look3.to_excel(writer, sheet_name="ProductInfo")
    # CProductinfo子表
    pd_look4 = pd.DataFrame(Custompartinfo_Data, index=Custompart_index, columns=Custompart_columns)
    pd_look4.index.name = Custompart_index_name
    pd_look4.to_excel(writer, sheet_name="CPartinfo")
    # CommonpartM子表
    pd_look6 = pd.DataFrame(SPartinfo_Data, index=Standardpart_index, columns=Standardpart_columns)
    pd_look6.index.name = Standardpart_index_name
    pd_look6.to_excel(writer, sheet_name="SPartinfo")
    # 最后保存写入，并释放
    writer.save()
    writer.close()

