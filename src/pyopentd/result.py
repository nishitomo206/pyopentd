import sys
import clr
import numpy as np
import pandas as pd
import typing as tp
from argparse import ArgumentParser

sys.path.append(
    "C:/Windows/Microsoft.NET/assembly/GAC_MSIL/OpenTDv62/v4.0_6.2.0.7__65e6d95ed5c2e178/"
)
clr.AddReference("OpenTDv62")
sys.path.append(
    "C:/Windows/Microsoft.NET/assembly/GAC_64/OpenTDv62.Results/v4.0_6.2.0.0__b62f614be6a1e14a/"
)
clr.AddReference("OpenTDv62.Results")

from System.Collections.Generic import List
import OpenTDv62 as otd


class SaveFile(otd.Results.Dataset.SaveFile):
    """SaveFile Class

    OpenTDv62.Results.Dataset.SaveFileを継承したクラス。OpenTDv62.Results.Dataset.SaveFileクラスについては、マニュアル(OpenTD 62 Class Reference.chm)を参照。
    また、self.topologyは対応するsavPCSファイル読み込み、OpenTDv62.Results.Dataset.Topologyを継承したクラスとなっている。

    """

    def __init__(self, sav_path):
        super().__init__(sav_path)
        self.times = self.GetTimes().GetValues()[:]
        self.sav_path = sav_path
        self.get_pcs_file()

    def get_pcs_file(self, record_index=-1):
        """モデルのトポロジー情報が格納されているsavPCSファイルを読み出す

        Args:
            record_index:savファイルの中の時系列の何番目の記録をベースにトポロジーを読み出すかの指定
            デフォルトでは最後の記録を読み出すようにしている（TDの設定で時系列の最後のデータは詳細情報を出力している設定にしていることが多いため）
        Returns:
            topology (OpenTDv62.Results.Dataset.Topology):OpenTD 62 Class Reference.chmを参照。IDatasetTopologyの中に詳細な構成が書かれている。
        """
        record_list = self.GetRecordNumbers()
        times = self.GetTimes().GetValues()[:]
        if record_index < 0:
            record_index = len(times) + record_index
        record_number = record_list[record_index]
        pcs_path = self.sav_path + "PCS"
        self.topology = otd.Results.Dataset.Topology.DatasetTopology.Load(
            self, record_number, pcs_path
        )

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
            info[f"{submodel_list[i]}"] = node_id_list
        return info

    def get_node_names(self, submodel_name="all", option=""):
        node_names = []
        submodel_list = self.GetThermalSubmodels()
        for submodel in submodel_list:
            node_id_list = self.GetNodeIds(submodel)
            for node_id in node_id_list:
                node_names.append(f"{submodel}.{option}{node_id}")
        return node_names

    def get_data(self, node_list):
        """時系列データ（結果）の取得（pandas.core.frame.DataFrameで出力）

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
            if ".T" in node_list[i]:
                tmp = [a - 273.15 for a in tmp]
            data.append(tmp)
        df = pd.DataFrame(np.array(data).transpose(), columns=node_list)
        times = self.GetTimes().GetValues()[:]
        df_times = pd.DataFrame(times, columns=["Times"])
        return pd.concat([df_times, df], axis=1)

    def get_data_array(self, node_list):
        """時系列データ（結果）の取得（numpy.arrayで出力）

        Args:
            node_list (list): ノードのリスト
        Returns:
            numpy.array (len(node_list) x len(times)): 時系列データ。
        Note:
            node_listの要素は、"AAA.1"ではなく、"AAA.T1"や"AAA.Q1"という形式で指定する。
        """
        data_td = self.GetData(node_list)
        data = []
        for i in range(data_td.Count):
            tmp = data_td[i].GetValues()[:]
            if ".T" in node_list[i]:
                tmp = [a - 273.15 for a in tmp]
            data.append(tmp)
        return np.array(data)

    def get_data_value(self, node_list, time_index):
        """単一の時系列データ（結果）の取得（numpy.arrayで出力）

        Args:
            node_list (list): ノードのリスト
            time_id (int): 時刻のインデックス
        Returns:
            numpy.array (len(node_list)): time_idで指定した時刻のデータ
        Note:
            node_listの要素は、"AAA.1"ではなく、"AAA.T1"や"AAA.Q1"という形式で指定する。
        """
        data_td = self.GetData(node_list)
        data = []
        times = self.GetTimes().GetValues()[:]
        if time_index < 0:
            time_index = len(times) + time_index
        for i in range(data_td.Count):
            tmp = data_td[i].GetValues()[time_index]
            if ".T" in node_list[i]:
                tmp = [a - 273.15 for a in tmp]
            data.append(tmp)
        return np.array(data)

    def get_all_temperature(self):
        """全ノードの温度データ取得"""
        node_list = self.get_node_names(option="T")
        return self.get_data(node_list)

    def get_all_heatrate(self):
        """全ノードの熱入力取得

        Note:
            外部熱入力（太陽熱入力）も含む。
        """
        node_list = self.get_node_names(option="Q")
        return self.get_data(node_list)

    def get_heat_capacity(self, time_index=-1):
        """全ノードの熱容量取得
        Args:
            record_index:savファイルの中の時系列の何番目の記録をベースにトポロジーを読み出すかの指定
            デフォルトでは最後の記録を読み出すようにしている（TDの設定で時系列の最後のデータは詳細情報を出力している設定にしていることが多いため）
        Returns:
            pandas.core.frame.DataFrame
        Note:
            基本的に熱容量は不変であることを想定しているが、時変で設定している場合はget_dataで取得可能。
        """
        node_list = self.get_node_names(option="C")
        capacity_array = self.get_data_value(node_list, time_index)
        df = pd.DataFrame(capacity_array.T, columns=["capacity"])
        return df

    def get_all_cond(self, time_index=-1):
        """全ての熱コンダクタンス・放射結合をリストで取得（pandas.core.frame.DataFrameで出力）"""

        cond_key_list = []
        from_node_list = []
        to_node_list = []
        is_rad_list = []
        node_id_list = []

        i_node_list = self.topology.Nodes
        for i_node in i_node_list:
            i_cond_list = i_node.Conductors
            for i_cond in i_cond_list:
                cond_key = i_cond.SindaName.Key
                if cond_key not in cond_key_list:
                    from_node = i_cond.FromNode.SindaName.Key
                    to_node = i_cond.ToNode.SindaName.Key
                    is_rad = i_cond.IsRad
                    node_id = (
                        cond_key.split(".")[0] + ".G" + cond_key.split(".")[1]
                    )

                    node_id_list.append(node_id)
                    cond_key_list.append(cond_key)
                    from_node_list.append(from_node)
                    to_node_list.append(to_node)
                    is_rad_list.append(is_rad)

        cond_value_list = self.get_data_value(node_id_list, time_index)

        list_all = list(
            zip(
                cond_key_list,
                from_node_list,
                to_node_list,
                is_rad_list,
                cond_value_list,
            )
        )
        df = pd.DataFrame(
            list_all, columns=["Key", "From", "To", "Is Rad", "Value"]
        )
        df_rad = df[df["Is Rad"] == True]
        df_cond = df[df["Is Rad"] == False]
        return df_rad, df_cond

    def get_cond_matrix(self, time_index=-1):
        """全ての熱コンダクタンス・放射結合を行列形式で取得（pandas.core.frame.DataFrameで出力）"""

        i_node_list = self.topology.Nodes
        cond_matrix = np.zeros((len(i_node_list), len(i_node_list)))
        rad_matrix = np.zeros((len(i_node_list), len(i_node_list)))
        node_index_dict = {}
        node_list = []

        for i in range(len(i_node_list)):
            node_index_dict[i_node_list[i].SindaName.Key] = i
            node_list.append(i_node_list[i].SindaName.Key)

        df_rad, df_cond = self.get_all_cond(time_index)

        for _, row in df_rad.iterrows():
            from_node = row["From"]
            to_node = row["To"]
            value = row["Value"]

            from_index = node_index_dict[from_node]
            to_index = node_index_dict[to_node]

            rad_matrix[from_index][to_index] = value
            rad_matrix[to_index][from_index] = value

        for _, row in df_cond.iterrows():
            from_node = row["From"]
            to_node = row["To"]
            value = row["Value"]

            from_index = node_index_dict[from_node]
            to_index = node_index_dict[to_node]

            cond_matrix[from_index][to_index] = value
            cond_matrix[to_index][from_index] = value

        df_rad_matrix = pd.DataFrame(
            rad_matrix, index=node_list, columns=node_list
        )
        df_cond_matrix = pd.DataFrame(
            rad_matrix, index=node_list, columns=node_list
        )

        return df_rad_matrix, df_cond_matrix
