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

class ThermalDesktop(otd.ThermalDesktop):
    """ThermalDesktop Class

    OpenTDv62.ThermalDesktopを継承したクラス。OpenTDv62.ThermalDesktopクラスについては、マニュアル(OpenTD 62 Class Reference.chm)を参照。

    """
    
    def __init__(self, dwg_path):
        super().__init__()
        dwg_path_obj = otd.Utility.RootedPathname(dwg_path)
        self.ConnectConfig.DwgPathname = dwg_path_obj
        self.Connect()
    
    def get_casesets(self):
        cases_td = self.GetCaseSets()
        cases = []
        for case_td in cases_td:
            cases.append(case_td.ToString().split('.'))
        df = pd.DataFrame(cases, columns=['group_name', 'caseset_name'])
        return df
    
    def get_orbit(self, orbit_name):
        orbit = self.GetOrbit(orbit_name)
        # 太陽方向vec
        sun_vecs = []
        for HrSunVec in orbit.HrSunVecArray:
            x = HrSunVec.X.GetValueSI()
            y = HrSunVec.Y.GetValueSI()
            z = HrSunVec.Z.GetValueSI()
            sun_vecs.append([x, y, z])
        df_sun = pd.DataFrame(sun_vecs, columns=['sun_x', 'sun_y', 'sun_z']) * 1000
        # 惑星（地球）方向vec
        planet_vecs = []
        for HrPlanetVec in orbit.HrPlanetVecArray:
            x = HrPlanetVec.X.GetValueSI()
            y = HrPlanetVec.Y.GetValueSI()
            z = HrPlanetVec.Z.GetValueSI()
            planet_vecs.append([x, y, z])
        df_planet = pd.DataFrame(planet_vecs, columns=['planet_x', 'planet_y', 'planet_z']) * 1000
        # 時間列
        times = []
        for HrTime in orbit.HrTimeArray:
            t = HrTime.GetValueSI()
            times.append(t)
        df_times = pd.DataFrame(times, columns=['Times'])
        # 距離時系列
        radiuses = orbit.HrOrbitRadiusArray
        df_radiuses = pd.DataFrame(radiuses, columns=['radius'])
        return pd.concat([df_times, df_sun, df_planet, df_radiuses], axis=1)
    
def get_orbit_name_from_caseset(case):
    orbit_name = ''
    rad_tasks = case.RadiationTasks
    for rad_task in rad_tasks:
        if orbit_name == '':
            orbit_name = rad_task.OrbitName
        elif orbit_name != rad_task.OrbitName:
            print('MyWarning: RadiationTasksのOrbitが共通ではありません。')
    return orbit_name
    
    

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
