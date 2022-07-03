import sys
import clr
import numpy as np
import pandas as pd
from argparse import ArgumentParser

sys.path.append("C:/Windows/Microsoft.NET/assembly/GAC_MSIL/OpenTDv62/v4.0_6.2.0.7__65e6d95ed5c2e178/")
clr.AddReference("OpenTDv62")
sys.path.append("C:/Windows/Microsoft.NET/assembly/GAC_64/OpenTDv62.Results/v4.0_6.2.0.0__b62f614be6a1e14a/")
clr.AddReference("OpenTDv62.Results")

from System.Collections.Generic import List
import OpenTDv62 as otd

class SaveFile(otd.Results.Dataset.SaveFile):
    """SaveFile Class

    OpenTDv62.Results.Dataset.SaveFileを継承したクラス。OpenTDv62.Results.Dataset.SaveFileクラスについては、マニュアル(OpenTD 62 Class Reference.chm)を参照。

    """
    
    def __init__(self, sav_path):
        super().__init__(sav_path)
        self.times = self.GetTimes().GetValues()[:]
    
    def get_submodels(self):
        return self.GetThermalSubmodels()
    
    def get_node_ids(self, submodel_name):
        return self.GetNodeIds(submodel_name)
    
    def get_node_info(self):
        info = {}
        submodel_list = self.GetThermalSubmodels()
        for i in range(len(submodel_list)):
            node_id_list = self.GetNodeIds(submodel_list[i])
            if node_id_list == []:
                continue
            info[f'{submodel_list[i]}'] = node_id_list
        return info
    
    def get_node_names(self, submodel_name='all', option=''):
        node_names = []
        submodel_list = self.GetThermalSubmodels()
        for submodel in submodel_list:
            node_id_list = self.GetNodeIds(submodel)
            for node_id in node_id_list:
                node_names.append(f'{submodel}.{option}{node_id}')
        return node_names
    
    def get_data(self, node_list):
        """時系列データ（結果）の取得

        Args:
            node_list (list): ノードのリスト
        Returns:
            pandas.core.frame.DataFrame: 時系列データ。ノードリストの情報に加えて、時間情報も追加されている。
        Note:
            node_listの要素は、"AAA.1"ではなく、"AAA.T1"や"AAA.Q1"という形式で指定する。
        """
        data_td = self.GetData(node_list)
        data = []
        for i in range(data_td.Count):
            tmp = data_td[i].GetValues()[:]
            if '.T' in node_list[i]:
                tmp = [a-273.15 for a in tmp]
            data.append(tmp)
        df = pd.DataFrame(np.array(data).transpose(), columns=node_list)
        times = self.GetTimes().GetValues()[:]
        df_times = pd.DataFrame(times, columns=['Times'])
        return pd.concat([df_times, df], axis=1)
    
    def get_all_temperature(self):
        """全ノードの温度データ取得
        """
        node_list = self.get_node_names(option='T')
        return self.get_data(node_list)
    
    def get_all_heatrate(self):
        """全ノードの熱入力取得
        
        Note:
            外部熱入力（太陽熱入力）も含む。
        """
        node_list = self.get_node_names(option='Q')
        return self.get_data(node_list)
