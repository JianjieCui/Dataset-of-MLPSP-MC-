import CreatMethod
import Time_ratio

# 机器信息
PMachine_Nums = 15  # 加工机器数
AMachine_Nums = 5  # 组装机器数

# 生产信息
Operations_Max = 10  # 工序数最大值
Product_Kinds = 20  # 定制产品种类
StandardPart_Kinds = 8  # 通用件种类

CustomPart_Max = 2  # 每种产品包含的定制件最大种类
PTime_Max = 10  # 加工时间最大值
PTime_Min = 5  # 加工时间最小值
ATime_Min = 5  # 组装时间最小值
ATime_Max = 10  # 加工时间最小值
# 生产任务信息
CPbatch_Max = 5  # 产品最大生产数量
CPbatch_Min = 1  # 产品最小生产数量
SPbatch_Max = 30  # 通用部件最大生产数量
SPbatch_Min = 10  # 通用部件最小生产数量
DT = 1  # 交货期松紧度
Earliness_Min = 0.1  # 提前惩罚最小值
Earliness_Max = 0.5  # 提前惩罚最大值
Tardiness_Min = 0.1  # 拖期惩罚最小值
Tardiness_Max = 0.5  # 拖期惩罚最大值
standard_total_ptime = 0
custom_total_time = 0
ratio = 0
file_name = 'L11'  # 算例名称

while ratio <= 0 or ratio >= 0.8:
    CreatMethod.main(file_name, PMachine_Nums, AMachine_Nums, Product_Kinds, StandardPart_Kinds, CustomPart_Max, Operations_Max, PTime_Max,
         PTime_Min, ATime_Min, ATime_Max, CPbatch_Max, CPbatch_Min, SPbatch_Max, SPbatch_Min, DT,
         Earliness_Min, Earliness_Max, Tardiness_Min, Tardiness_Max)
    standard_total_ptime, custom_total_time, ratio = Time_ratio.main(file_name, Operations_Max)
