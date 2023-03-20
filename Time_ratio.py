import numpy as np
import xlrd

"""计算算例中通用件与定制件的总处理时间，以及时间比例，用于对算例进行分类"""


def main(file_name, operations_max):

    '''获取算法求解需要的实验算例'''

    excel_path = "../" + file_name + '.xlsx'  # 存放位置

    # 打开文件，获取excel文件的workbook（工作簿）对象
    excel = xlrd.open_workbook(excel_path, encoding_override="utf-8")

    # 获取sheet对象
    all_sheet = excel.sheets()
    # 机器信息
    machinep_nums = int(all_sheet[0].cell(1, 1).value)  # 加工机器数
    Machine_ProId_list = [i + 1 for i in range(machinep_nums)]  # 机器编号列表
    print('Machine_ProId_list:', Machine_ProId_list)

    # 定制产品信息
    # Product_Id_list = list(map(lambda x: int(x), all_sheet[1].col_values(0)[1:]))  # 产品编号
    Product_CustomId_list = []  # 各产品组装所需的定制部件，[[1,2],[3,4],……]，第一个列表内是第一种产品，第二个是第二种
    Product_Batch_list = list(map(lambda x: int(x), all_sheet[1].col_values(4)[1:]))  # 生产数量
    for each_row in range(all_sheet[1].nrows - 1):
        Product_CustomId_list.append([])
        for i in all_sheet[1].row_values(each_row + 1)[1:3]:
            id = int(i)
            if id == -1:  # 若ID=-1，则说明是凑数的，后面都是无用信息
                break
            else:
                Product_CustomId_list[-1].append(id)
    print('Product_CustomId_list:', Product_CustomId_list)
    print('Product_Batch_list', Product_Batch_list)

    CustomPart_Id_list = list(map(lambda x: int(x), all_sheet[2].col_values(0)[1:]))  # 定制件编号集
    CustomP_OperationMac_list = []  # 定制产品各工序可用机器
    CustomP_OperationTim_list = []  # 定制产品各工序在机器上的加工时间
    for i in range(len(CustomPart_Id_list)):
        CustomP_OperationTim_list.append([])
        CustomP_OperationMac_list.append([])
        list_tem = list(map(lambda x: int(x), all_sheet[2].row_values(i + 1)[1:]))  # 当前行的所有信息
        for j in range(operations_max):
            index1 = j * machinep_nums
            if list_tem[index1] == -1:
                break
            CustomP_OperationTim_list[i].append([])
            CustomP_OperationMac_list[i].append([])
            for k in range(index1, index1 + len(Machine_ProId_list)):
                if list_tem[k] != 1000:
                    m_id = Machine_ProId_list[k - index1]  # 加工机器编号
                    CustomP_OperationTim_list[i][j].append(list_tem[k])
                    CustomP_OperationMac_list[i][j].append(m_id)
    print('CPart_Id_list:', CustomPart_Id_list)
    print('CustomP_OperationTim_list:', CustomP_OperationTim_list)
    print('CustomP_OperationMac_list:', CustomP_OperationMac_list)

    # 标准产品信息
    StandradP_Id_list = list(map(lambda x: int(x), all_sheet[3].col_values(0)[1:]))  # 标准产品编号
    StandradP_OperationMac_list = []  # 各通用部件各工序可用机器，[[],[],……]
    StandradP_OperationTim_list = []  # 各通用部件在各机器上的加工时间,[[],[],……]
    for i in range(len(StandradP_Id_list)):
        StandradP_OperationTim_list.append([])
        StandradP_OperationMac_list.append([])
        list_tem = list(map(lambda x: int(x), all_sheet[3].row_values(i + 1)[1:-1]))  # 当前行的所有信息
        for j in range(operations_max):
            index1 = j * machinep_nums
            if list_tem[index1] == -1:
                break
            StandradP_OperationTim_list[i].append([])
            StandradP_OperationMac_list[i].append([])
            for k in range(index1, index1 + len(Machine_ProId_list)):
                if list_tem[k] != 1000:
                    m_id = Machine_ProId_list[k - index1]  # 加工机器编号
                    StandradP_OperationTim_list[i][j].append(list_tem[k])
                    StandradP_OperationMac_list[i][j].append(m_id)
    print('StandradP_Id_list:', StandradP_Id_list)
    print('StandradP_OperationMac_list:', StandradP_OperationMac_list)
    print('StandradP_OperationTim_list:', StandradP_OperationTim_list)

    # 通用部件生产任务信息
    StandardP_Num_list = list(map(lambda x: int(x), all_sheet[3].col_values(-1)[1:]))  # 通用部件生产数量
    print('StandardP_Num_list:', StandardP_Num_list)

    common_total_ptime = 0  # 通用件总处理时间
    for i in range(len(StandradP_Id_list)):
        total_ptime = sum(
            [np.mean(process_time) for process_time in StandradP_OperationTim_list[i]])  # 通用件各工序总加工时间
        total_ptime *= StandardP_Num_list[i]  # 批加工时间
        common_total_ptime += total_ptime

    custom_total_time = 0  # 定制件总处理时间
    for i in range(len(Product_Batch_list)):
        batch = Product_Batch_list[i]  # 生产数量
        total_ptime = 0
        for custom_id in Product_CustomId_list[i]:
            time_tem = sum([np.mean(process_time) for process_time in CustomP_OperationTim_list[custom_id - 1]])
            total_ptime += time_tem  # 把组成产品的所有定制件平均加工时间加起来
        total_ptime *= batch  # 乘上批量
        custom_total_time += total_ptime  # 所有产品总处理时间
    ratio = common_total_ptime / custom_total_time  # 通用件/定制件
    print(f'通用件总处理时间：{common_total_ptime}')
    print(f'定制件总处理时间：{custom_total_time}')
    print(f'通用件/定制件：{ratio}')
    return [common_total_ptime, custom_total_time, ratio]

# common_total_ptime, custom_total_time, ratio = main('S01', 4)
