# -*- coding: utf-8 -*-
import numpy as np
import geatpy as ea
import xlrd
import win32com.client as com  # VISSIM COM
import time

"""
该案例展示了一个简单的连续型决策变量最大化目标的单目标优化问题。
max f = x * np.sin(10 * np.pi * x) + 2.0
s.t.
-1 <= x <= 2
"""


class MyProblem(ea.Problem):  # 继承Problem父类
    flag = False

    def __init__(self):
        name = 'MyProblem'  # 初始化name（函数名称，可以随意设置）
        M = 1  # 初始化M（目标维数）
        maxormins = [1]  # 初始化maxormins（目标最小最大化标记列表，1：最小化该目标；-1：最大化该目标）
        Dim = 3  # 初始化Dim（决策变量维数）
        varTypes = [0] * Dim  # 初始化varTypes（决策变量的类型，元素为0表示对应的变量是连续的；1表示是离散的）
        lb = [1.2, 1.0, 2]  # 决策变量下界 CC0推荐：1.2-1.7
        ub = [1.7, 2.0, 7]  # 决策变量上界
        lbin = [1, 1, 1]  # 决策变量下边界
        ubin = [1, 1, 1]  # 决策变量上边界
        # 调用父类构造方法完成实例化
        ea.Problem.__init__(self, name, M, maxormins, Dim, varTypes, lb, ub, lbin, ubin)

    def aimFunc(self, pop):  # 目标函数
        x = pop.Phen  # 得到决策变量矩阵 x=50
        x0 = x[:, [0]]
        x1 = x[:, [1]]
        x2 = x[:, [2]]
        # x1 = x[:, 1]
        # x2 = x[:, 2]
        # print(x)
        if self.flag == False:
            self.Data = xlrd.open_workbook('resultToCali2.xlsx')
            self.Table = self.Data.sheet_by_name(u'Sheet1')
            self.flag = True
            self.V1 = self.Table.col_values(0)  # spd for segment 1
            self.V2 = self.Table.col_values(1)  # spd for segment 2
            self.V3 = self.Table.col_values(2)  # spd for segment 2
            self.V4 = self.Table.col_values(3)  # spd for segment 2
            self.relativeFlow = self.Table.col_values(4)  # relative flow
            self.volume = self.Table.col_values(5)  # 流量
        totalResult = []
        for k in range(30):
            # # =================================================================================
            # # VISSIM Configurations
            # # Load VISSIM Network
            self.Vissim = com.Dispatch("Vissim.Vissim");
            dir = "E:\\PycharmProjects\\GA-calibration\\test.inp"
            self.Vissim.LoadNet(dir)
            # Define Simulation Configurations
            graphics = self.Vissim.Graphics
            graphics.SetAttValue("VISUALIZATION", False)  ## 设为 不可见 提高效率
            self.Sim = self.Vissim.Simulation
            self.Net = self.Vissim.Net

            # G = self.Vissim.Graphics
            dbpss = self.Net.DrivingBehaviorParSets  # Driving behavior module
            dbps = dbpss(3)
            # # Set Simulation Parameters
            TotalPeriod = 82802  # Define total simulation period
            WarmPeriod = 900  # Define warm period 10 minutes
            Random_Seed = k  # Define random seed
            step_time = 1  # Define Step Time
            self.Sim.Period = TotalPeriod
            self.Sim.RandomSeed = 42
            # self.Sim.RunIndex= 1
            self.Sim.Resolution = step_time

            # Each scenario run 5 times
            # =================================================================================
            # Data Collection Variables
            t1 = []
            t2 = []
            t3 = []
            t4 = []
            nVeh = []
            dbps.SetAttValue('CC0', x0[k][0])  ## 天坑！ 不能写x0[k]！！！！！！
            dbps.SetAttValue('CC1', x1[k][0])
            dbps.SetAttValue('CC2', x2[k][0])
            print("第", k, "组参数： CC0=", x0[k][0], ",CC1=", x1[k][0], ",CC2=", x2[k][0])
            eval = self.Vissim.Evaluation
            eval.SetAttValue("TRAVELTIME", True)
            eval.SetAttValue("DATACOLLECTION", True)
            TT1 = self.Net.TravelTimes(1)
            TT2 = self.Net.TravelTimes(2)
            dataCollections = self.Vissim.Net.DataCollections
            dt1 = dataCollections(1)
            dt2 = dataCollections(2)
            dt3 = dataCollections(3)
            dt4 = dataCollections(4)
            composition = self.Net.TrafficCompositions(1)
            vehicleInput = self.Net.VehicleInputs(1)
            # self.Sim.RunContinuous()
            for j in range(1, TotalPeriod):
                if (j % 900 == 0) and (j >= 1801):
                    composition.SetAttValue1("RELATIVEFLOW", 100, 1 - self.relativeFlow[int(j / 900-3)])
                    composition.SetAttValue1("RELATIVEFLOW", 200, self.relativeFlow[int(j / 900-3)])
                    vehicleInput.SetAttValue("VOLUME", self.volume[int(j / 900-3)])
                    t1.append(dt1.GetResult("speed", "mean", 0))  # 车道1
                    t2.append(dt2.GetResult("speed", "mean", 0))  # 车道2
                    t3.append(dt3.GetResult("speed", "mean", 0))  # 车道3
                    t4.append(dt4.GetResult("speed", "mean", 0))  # 车道4
                    nVeh.append(4 * (dt1.GetResult("NVEHICLES", "sum", 0)
                                     + dt2.GetResult("NVEHICLES", "sum",0)
                                     + dt3.GetResult("NVEHICLES", "sum", 0)
                                     + dt4.GetResult("NVEHICLES", "sum", 0)))
                self.Sim.RunSingleStep()

            spdTotal = np.array(0.25 * sum(abs(np.array(t1) - np.array(self.V1)) / np.array(self.V1)
                                           + abs(np.array(t2) - np.array(self.V2)) / np.array(self.V1)
                                           + abs(np.array(t3) - np.array(self.V3)) / np.array(self.V1)
                                           + abs(np.array(t4) - np.array(self.V4)) / np.array(self.V1)))
            tTimeTotal = np.array(sum(abs(np.array(nVeh) - np.array(self.volume)) / np.array(self.volume)))
            totalResult.append(spdTotal+tTimeTotal)
            print("总误差为：", totalResult[k])
            self.Sim.Stop()
        pop.ObjV = np.vstack(totalResult)  # 计算目标函数值，赋值给pop种群对象的ObjV属性
