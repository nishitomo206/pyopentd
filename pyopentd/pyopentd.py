import sys
import clr
import numpy as np
import pandas as pd
from argparse import ArgumentParser

sys.path.append("C:/Windows/Microsoft.NET/assembly/GAC_MSIL/OpenTDv62/v4.0_6.2.0.7__65e6d95ed5c2e178/")
clr.AddReference("OpenTDv62")
sys.path.append("C:/Windows/Microsoft.NET/assembly/GAC_64/OpenTDv62.Results/v4.0_6.2.0.0__b62f614be6a1e14a/")
clr.AddReference("OpenTDv62.Results")

# from System import *
from System.Collections.Generic import List
import OpenTDv62 as otd

def open_dwg(dwg_path):
    td = otd.ThermalDesktop()
    dwg_path_obj = otd.Utility.RootedPathname(dwg_path)
    td.ConnectConfig.DwgPathname = dwg_path_obj
    td.Connect()
    return td

def create_orbit(td, csv_path, orbit_name="new_orbit"):
    # 新しい軌道の作成
    orbit_name = "new_orbit"
    orbit = td.CreateOrbit(orbit_name)
    orbit.OrbitType = otd.RadCAD.Orbit.OrbitTypes.TRAJECTORY # orbitタイプの変更。
    
    # csv読み込み
    header = ['time', 'sun_bf_x', 'sun_bf_y', 'sun_bf_z', 'earth_bf_x', 'earth_bf_y', 'earth_bf_z', 'radius']
    df = pd.read_csv(csv_path, header=None)
    df.columns = header
    
    # 軌道情報作成
    solar_vector_list = List[otd.Vector3d]() # Solar Vectorの履歴
    planet_vector_list = List[otd.Vector3d]() # Vector to Earthの履歴
    radius_list = List[float]() # radiusの履歴
    for index, row in df.iterrows():
        solar_vector_list.Add(otd.Vector3d(row[1]/1000,row[2]/1000,row[3]/1000))
        planet_vector_list.Add(otd.Vector3d(row[4]/1000,row[5]/1000,row[6]/1000))
        radius_list.Add(row[7])
    time_list = list(df['time'].values.astype(float))
    time_list = otd.Dimension.DimensionalList[otd.Dimension.Time](time_list)
    # 軌道情報更新
    orbit.HrSunVecArray = solar_vector_list
    orbit.HrPlanetVecArray = planet_vector_list
    orbit.HrTimeArray = time_list
    orbit.HrOrbitRadiusArray = radius_list
    orbit.Update()
    return orbit

def update_symbol(case, symbol_csv_path):
    symbol_name_list = case.SymbolNames
    symbol_value_list = case.SymbolValues
    
    df = pd.read_csv(symbol_csv_path)
    for index, row in df.iterrows():
        if row['name'] in symbol_name_list:
            index = symbol_name_list.index(row['name'])
            symbol_value_list[index] = str(row['value'])
        else:
            # print("シンボルが見つかりませんでした、追加します。")
            symbol_name_list.append(row['name'])
            symbol_value_list.append(row['value'])
    symbol_names = List[str]()
    symbol_values = List[str]()
    for i in range(len(symbol_name_list)):
        symbol_names.Add(symbol_name_list[i])
        symbol_values.Add(symbol_value_list[i])
    # 更新
    case.SymbolNames = symbol_names
    case.SymbolValues = symbol_values
    case.Update()
    return case

def update_orbit(case, orbit_name): #TODO ここの入力をorbitで対応できるようにする。
    # 軌道参照先変更
    for i in range(len(case.RadiationTasks)):
        case.RadiationTasks[i].OrbitName = orbit_name
    case.Update()
    return case

def change_sav_name(case, sav_name):
    case.SindaOptions.SaveFilename = sav_name
    case.Update()
    return case

def run_one_case(td, caseset_group_name, caseset_name, orbit_csv_path="", symbol_csv_path="", sav_name=""):
    # ケースセットの読み込み
    case = td.GetCaseSet(caseset_name, caseset_group_name)
    
    # 軌道の変更
    if orbit_csv_path != "":
        orbit_name = "new_orbit"
        orbit = create_orbit(td, orbit_csv_path, orbit_name)
        orbit.Update()
        # 軌道参照先変更
        for i in range(len(case.RadiationTasks)):
            case.RadiationTasks[i].OrbitName = orbit_name
        case.Update()
    
    # シンボルの更新
    if symbol_csv_path != "":
        case = update_symbol(case, symbol_csv_path)
        case.Update()
    
    # 保存先の更新
    if sav_name != "":
        case.SindaOptions.SaveFilename = sav_name
        case.Update()
    
    case.Run()
    return


def main(dwg_path, caseset_group_name, caseset_name, orbit_csv_path="", symbol_csv_path="", sav_name=""):
    td = otd.ThermalDesktop()
    dwg_path_obj = otd.Utility.RootedPathname(dwg_path)
    td.ConnectConfig.DwgPathname = dwg_path_obj
    td.Connect()
    
    run_one_case(td, caseset_group_name, caseset_name, orbit_csv_path, symbol_csv_path, sav_name)
    
    return

def read_savefile(sav_path):
    """savefileの読み込み

    Args:
        sav_path (str): savファイルへのpath。
    Returns:
        OpenTDv62.Results.Dataset.SaveFile: savefileクラスのオブジェクト
    """
    return otd.Results.Dataset.SaveFile(sav_path)

def get_submodels(savefile):
    return savefile.GetThermalSubmodels()

def get_node_ids(savefile, submodel_name):
    return savefile.GetNodeIds(submodel_name)

def get_node_info(savefile):
    info = {}
    submodel_list = savefile.GetThermalSubmodels()
    for i in range(len(submodel_list)):
        node_id_list = savefile.GetNodeIds(submodel_list[i])
        if node_id_list == []:
            continue
        info[f'{submodel_list[i]}'] = node_id_list
    return info

def get_node_names(savefile, submodel_name='all', option=''):
    node_names = []
    submodel_list = savefile.GetThermalSubmodels()
    for submodel in submodel_list:
        node_id_list = savefile.GetNodeIds(submodel)
        for node_id in node_id_list:
            node_names.append(f'{submodel}.{option}{node_id}')
    return node_names

def get_data(savefile, node_list):
    """時系列データ（結果）の取得

    Args:
        savefile (OpenTDv62.Results.Dataset.SaveFile): save file
        node_list (list): ノードのリスト
    Returns:
        pandas.core.frame.DataFrame: 時系列データ。ノードリストの情報に加えて、時間情報も追加されている。
    Note:
        node_listの要素は、"AAA.1"ではなく、"AAA.T1"や"AAA.Q1"という形式で指定する。
    """
    data_td = savefile.GetData(node_list)
    data = []
    for i in range(data_td.Count):
        data.append(data_td[i].GetValues()[:])
    df = pd.DataFrame(np.array(data).transpose(), columns=node_list) -273.15
    times = savefile.GetTimes().GetValues()[:]
    df_times = pd.DataFrame(times, columns=['Times'])
    return pd.concat([df_times, df], axis=1)

def get_all_temperature(savefile):
    node_list = get_node_names(savefile, option='T')
    return get_data(savefile, node_list)
