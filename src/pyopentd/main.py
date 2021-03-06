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
import System

class ThermalDesktop(otd.ThermalDesktop):
    """ThermalDesktop Class

    OpenTDv62.ThermalDesktopを継承したクラス。OpenTDv62.ThermalDesktopクラスについては、マニュアル(OpenTD 62 Class Reference.chm)を参照。

    """
    
    def __init__(self, dwg_path, visible=True):
        super().__init__()
        dwg_path_obj = otd.Utility.RootedPathname(dwg_path)
        self.ConnectConfig.DwgPathname = dwg_path_obj
        self.ConnectConfig.AcadVisible = visible
        self.Connect()
    
    def get_casesets(self):
        cases_td = self.GetCaseSets()
        cases = []
        for case_td in cases_td:
            cases.append(case_td.ToString().split('.'))
        df = pd.DataFrame(cases, columns=['group_name', 'caseset_name'])
        return df
    
    def get_caseset(self, caseset_name, group_name):
        return Case(self.GetCaseSet(caseset_name, group_name))
    
    def get_heatloads(self):
        #* アプライ先、センサー先のノードは1つであると仮定。
        #* ノードヒーターであると仮定。
        #TODO 色々なタイプのノードに対応させる。
        heatloads = self.GetHeatLoads()
        heatloads_list = []
        for heatload in heatloads:
            apply_node = None
            name = heatload.Name
            submodel = heatload.Submodel
            handle = heatload.Handle
            if heatload.ApplyConnections != []:
                node = self.GetNode(heatload.ApplyConnections[0].Handle)
                apply_node = f'{node.Submodel}.{node.Id}'
            if len(heatload.ApplyConnections) >= 2: print("MYWARNING: 1つのヒーターが2つ以上のノードに適用されています。")
            heatloads_list.append([name, submodel, handle, apply_node, heatload])
        header = ['Name', 'submodel', 'handle', 'apply_node', 'original_object']
        df_heatloads = pd.DataFrame(heatloads_list, columns=header)
        return df_heatloads
    
    def get_heaters(td):
        #* アプライ先、センサー先のノードは1つであると仮定。
        #* ノードヒーターであると仮定。
        #TODO 色々なタイプのノードに対応させる。
        heaters = td.GetHeaters()
        heaters_list = []
        for heater in heaters:
            apply_node = None
            sensor_node = None
            name = heater.Name
            submodel = heater.Submodel
            handle = heater.Handle
            if heater.ApplyConnections != []:
                node = td.GetNode(heater.ApplyConnections[0].Handle)
                apply_node = f'{node.Submodel}.{node.Id}'
            if len(heater.ApplyConnections) >= 2: print("MYWARNING: 1つのヒーターが2つ以上のノードに適用されています。")
            if heater.SensorConnections != []:
                node = td.GetNode(heater.SensorConnections[0].Handle)
                sensor_node = f'{node.Submodel}.{node.Id}'
            if len(heater.SensorConnections) >= 2: print("MYWARNING: 1つのヒーターが2つ以上のノードの温度を監視してます。")
            heaters_list.append([name, submodel, handle, apply_node, sensor_node, heater])
        header = ['Name', 'submodel', 'handle', 'apply_node', 'sensor_node', 'original_object']
        df_heaters = pd.DataFrame(heaters_list, columns=header)
        return df_heaters
    
    def get_heater(self, handle):
        return self.GetHeater(handle)
    
    def get_orbits(td):
        orbits_list = []
        orbits = td.GetOrbits()
        for orbit in orbits:
            orbits_list.append([orbit.Name, orbit.OrbitType, orbit])
        df_orbits = pd.DataFrame(orbits_list, columns=['Name', 'OrbitType', 'original_object'])
        return df_orbits
    
    def get_orbit(self, orbit_name):
        """軌道データの所得"""
        orbit = self.GetOrbit(orbit_name)
        # 太陽方向vec
        sun_vecs = []
        for HrSunVec in orbit.HrSunVecArray:
            x = HrSunVec.X.GetValueSI()
            y = HrSunVec.Y.GetValueSI()
            z = HrSunVec.Z.GetValueSI()
            sun_vecs.append([x, y, z])
        df_sun = pd.DataFrame(sun_vecs, columns=['sun_x', 'sun_y', 'sun_z'])
        if df_sun.max().max() < 0.0011:
            df_sun *= 1000
        if df_sun.max().max() > 1.1:
            df_sun /= 1000
        # 惑星（地球）方向vec
        planet_vecs = []
        for HrPlanetVec in orbit.HrPlanetVecArray:
            x = HrPlanetVec.X.GetValueSI()
            y = HrPlanetVec.Y.GetValueSI()
            z = HrPlanetVec.Z.GetValueSI()
            planet_vecs.append([x, y, z])
        df_planet = pd.DataFrame(planet_vecs, columns=['planet_x', 'planet_y', 'planet_z']) * 1000
        if df_planet.max().max() < 0.0011:
            df_planet *= 1000
        if df_planet.max().max() > 1.1:
            df_planet /= 1000
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
    
    def get_nodes(self, printif=False):
        nodes = self.GetNodes()
        if printif:
            for node in nodes:
                print('----------')
                print(node)
                print('Handle: ', node.Handle)
                print('Submodel: ', node.Submodel)
                print('Id: ', node.Id)
        return nodes
    
    def get_node(self, submodel, id, printif=False):
        nodes = self.GetNodes()
        node = 0
        for node_i in nodes:
            if node_i.Submodel.ToString()==submodel and node_i.Id==int(id):
                node = self.GetNode(node_i.Handle)
                break
        if node == 0:
            print('MYERROR (in pyopentd.main.get_node): 指定したnodeが見つかりませんでした。')
            return
        if printif:
            print(node)
            print('Handle: ', node.Handle)
            print('Submodel: ', node.Submodel)
            print('Id: ', node.Id)
        return node
    
    def create_caseset(self, name, group, steady, transient, time_end=0, run_dir='', sumodels_not_built=[], restart_file=''):
        case = self.CreateCaseSet(name, group, name)
        case = Case(case)
        case.origin.SaveQ = 1 # outputにincident heatを含める。
        case.origin.SteadyState = steady
        case.origin.Transient = transient
        if transient == 1:
            if time_end == 0: print('MYERROR: 終了時刻(time_end)を指定してください。')
            case.origin.SindaControl.timendExp.Value = str(time_end)
        
        # リスタートファイルを使用する
        if restart_file != '':
            case.origin.UseRestartFile = 1
            case.origin.RestartFile = restart_file
        
        # BuildTypeの指定（シミュレーションしないサブモデルの指定）
        if sumodels_not_built != []:
            case.origin.BuildType = 2
            tmp_list = List[str]()
            for submodel in sumodels_not_built:
                tmp_list.Add(submodel)
            case.origin.SubmodelsNotBuilt = tmp_list
        
        # use_run_directoryの設定
        if run_dir != '':
            case.origin.UseUserDirectory = 1
            case.origin.UserDirectory = run_dir
        case.update()
        return case
    
    def create_heatload(self, submodel, id, transient_type, value=-1, time_array=[], value_array=[], name='', layer='', enable_exp=''):
        node = self.get_node(submodel, id) # ノードの取得
        heatload = self.CreateHeatLoad(otd.Connection(node)) # heatloadの作成
        
        if transient_type == 0: # constant heatload の作成
            # typeの変更
            heatload.HeatLoadTransientType = 0
            heatload.TimeDependentSteadyStateType = 0
            if value == -1:
                print('MYWARNING (in pyopentd.main.create_heatload): ヒートロードのvalueが指定されていません。')
            # valueの変更
            if type(value) == str:
                heatload.ValueExp.Value = value
            else:
                heatload.Value = value
        elif transient_type == 1: # time dependent heatload の作成
            # チェック
            if len(time_array) != len(value_array):
                print('MYERROR (in pyopentd.main.create_heatload): time_arrayとvalue_arrayの長さが異なります。')
            # typeの変更
            heatload.HeatLoadTransientType = 1
            heatload.TimeDependentSteadyStateType = 1
            # time arrayの変更
            TimeArray = List[System.String]()
            for time in time_array:
                TimeArray.Add(str(time))
            heatload.TimeArrayExp.expression = TimeArray
            # value arrayの変更
            ValueArray = List[System.String]()
            for val in value_array:
                ValueArray.Add(str(val))
            heatload.ValueArrayExp.expression = ValueArray
        else:
            print('MYERROR (in pyopentd.main.create_heatload): transient_typeが正しくありません。')
        
        heatload.Submodel.Name = submodel # hwatloadのSubmodelを変更
        if name != '': heatload.Name = name
        if layer != '': heatload.Layer = layer
        if enable_exp != '': heatload.EnabledExp.Value = enable_exp
        
        heatload.Update()
        return heatload
    
    def add_symbol(self, name, value, group=''):
        symbol = self.CreateSymbol(name, str(value))
        symbol.Group = group
        symbol.Update()

    def create_orbit(self, df_orbit, orbit_name="new_orbit"):
        """新規軌道作成
        
        Args:
            df_orbit (pandas.core.frame.DataFrame): 軌道のDataframe(カラムは['time', 'sun_x', 'sun_y', 'sun_z', 'earth_x', 'earth_y', 'earth_z', 'radius'])
            orbit_name (str): 軌道の名前
        """
        # カラム名チェック
        column = ['Times', 'sun_x', 'sun_y', 'sun_z', 'planet_x', 'planet_y', 'planet_z', 'radius']
        if df_orbit.columns.values.tolist() != column:
            print("MYERROR (in pyopentd.main.create_orbit): DataFrameのカラム名が正しくありません。['Times', 'sun_x', 'sun_y', 'sun_z', 'planet_x', 'planet_y', 'planet_z', 'radius']に変更して下さい。")
        
        # 新しい軌道の作成
        orbit = self.CreateOrbit(orbit_name)
        orbit.OrbitType = otd.RadCAD.Orbit.OrbitTypes.TRAJECTORY # orbitタイプの変更。
        
        # 軌道情報作成
        solar_vector_list = List[otd.Vector3d]() # Solar Vectorの履歴
        planet_vector_list = List[otd.Vector3d]() # Vector to Earthの履歴
        radius_list = List[float]() # radiusの履歴
        for index, row in df_orbit.iterrows():
            solar_vector_list.Add(otd.Vector3d(row['sun_x']/1000,row['sun_y']/1000,row['sun_z']/1000))
            planet_vector_list.Add(otd.Vector3d(row['planet_x']/1000,row['planet_y']/1000,row['planet_z']/1000))
            radius_list.Add(row['radius'])
        time_list = list(df_orbit['Times'].values.astype(float))
        time_list = otd.Dimension.DimensionalList[otd.Dimension.Time](time_list)
        # 軌道情報更新
        orbit.HrSunVecArray = solar_vector_list
        orbit.HrPlanetVecArray = planet_vector_list
        orbit.HrTimeArray = time_list
        orbit.HrOrbitRadiusArray = radius_list
        orbit.Update()
        return orbit

class Case():
    """Case Class
    
    ケースセットのクラス。self.originにOpenTDv62.CaseSetクラスのインスタンスを持つ。
    OpenTDv62.CaseSetクラスについては、マニュアル(OpenTD 62 Class Reference.chm)を参照。

    """
    
    def __init__(self, case):
        self.origin = case
    
    def update(self):
        self.origin.Update()
        return

    def run(self):
        self.origin.Run()
        return
    
    def get_orbit_name(self):
        orbit_name = ''
        rad_tasks = self.origin.RadiationTasks
        for rad_task in rad_tasks:
            if orbit_name == '':
                orbit_name = rad_task.OrbitName
            elif orbit_name != rad_task.OrbitName:
                print('MyWarning: RadiationTasksのOrbitが共通ではありません。')
        return orbit_name
    
    def get_symbols(self):
        symbols = np.array([self.origin.SymbolNames,
                            self.origin.SymbolValues,
                            self.origin.SymbolComments]).T
        df = pd.DataFrame(symbols, columns=['name', 'value', 'comment'])
        return df
    
    def update_symbol(self, df_symbol):
        """変数の変更・追加
        
        Args:
            df_symbol (pandas.core.frame.DataFrame): 変更したい変数のDataframe('name'と'value'をカラムにもつ)
        """
        symbol_name_list = self.origin.SymbolNames
        symbol_value_list = self.origin.SymbolValues
        symbol_comment_list = self.origin.SymbolComments
        
        for index, row in df_symbol.iterrows():
            if row['name'] in symbol_name_list:
                index = symbol_name_list.index(row['name'])
                symbol_value_list[index] = str(row['value'])
            else:
                # print("シンボルが見つかりませんでした、追加します。")
                symbol_name_list.append(row['name'])
                symbol_value_list.append(str(row['value']))
                symbol_comment_list.append('')
        symbol_names = List[str]()
        symbol_values = List[str]()
        symbol_comments = List[str]()
        for i in range(len(symbol_name_list)):
            symbol_names.Add(symbol_name_list[i])
            symbol_values.Add(symbol_value_list[i])
            symbol_comments.Add(symbol_comment_list[i])
        # 更新
        self.origin.SymbolNames = symbol_names
        self.origin.SymbolValues = symbol_values
        self.origin.SymbolComments = symbol_comments
        self.update()
        return
    
    def update_orbit(self, orbit_name): #TODO ここの入力をorbitで対応できるようにする。
        # 軌道参照先変更
        for i in range(len(self.origin.RadiationTasks)):
            self.origin.RadiationTasks[i].OrbitName = orbit_name
        self.update()
        return
    
    def change_sav_name(self, sav_name):
        self.origin.SindaOptions.SaveFilename = sav_name
        self.update()
        return
    
    def add_radiation_task(self, calc_type, orbit_name='', analysis_group='BASE'):
        # rad task の作成
        rad_task = otd.RadiationTaskData()
        rad_task.AnalGroup = analysis_group # Analysis Groupの指定
        rad_task.TypeCalc = calc_type # [0: Radks, 1: Heating Rates, 2: Articulating Radks]
        if calc_type == 1 or calc_type == 2:
            rad_task.OrbitName = orbit_name
        
        rad_task.RkFilename = f'{self.origin.GroupName}_{self.origin.Name}.k'
        rad_task.RkSubmodel = f'{self.origin.GroupName}_{self.origin.Name}'
        rad_task.HrFilename = f'{self.origin.GroupName}_{self.origin.Name}.hr'
        rad_task.HrSubmodel = f'{self.origin.GroupName}_{self.origin.Name}'
        rad_task.FfFilename = f'{self.origin.GroupName}_{self.origin.Name}.dat'
        rad_task.OutputTrackerDataFile = f'{self.origin.GroupName}_{self.origin.Name}.dat'
        
        # 追加
        if self.origin.RadiationTasks == []:
            rad_task_list = List[otd.RadiationTaskData]()
            rad_task_list.Add(rad_task)
            self.origin.RadiationTasks = rad_task_list
        else:
            original_rad_task_list = self.origin.RadiationTasks
            rad_task_list = List[otd.RadiationTaskData]()
            for task in original_rad_task_list:
                rad_task_list.Add(task)
            rad_task_list.Add(rad_task)
            self.origin.RadiationTasks = rad_task_list
        self.update()
        return


# def run_one_case(td, caseset_group_name, caseset_name, orbit_csv_path="", symbol_csv_path="", sav_name=""):
#     # ケースセットの読み込み
#     case = td.GetCaseSet(caseset_name, caseset_group_name)
    
#     # 軌道の変更
#     if orbit_csv_path != "":
#         orbit_name = "new_orbit"
#         orbit = create_orbit(td, orbit_csv_path, orbit_name)
#         orbit.Update()
#         # 軌道参照先変更
#         for i in range(len(case.RadiationTasks)):
#             case.RadiationTasks[i].OrbitName = orbit_name
#         case.Update()
    
#     # シンボルの更新
#     if symbol_csv_path != "":
#         case = update_symbol(case, symbol_csv_path)
#         case.Update()
    
#     # 保存先の更新
#     if sav_name != "":
#         case.SindaOptions.SaveFilename = sav_name
#         case.Update()
    
#     case.Run()
#     return

# def create_orbit(td, csv_path, orbit_name="new_orbit"):
#     # 新しい軌道の作成
#     orbit_name = "new_orbit"
#     orbit = td.CreateOrbit(orbit_name)
#     orbit.OrbitType = otd.RadCAD.Orbit.OrbitTypes.TRAJECTORY # orbitタイプの変更。
    
#     # csv読み込み
#     header = ['time', 'sun_bf_x', 'sun_bf_y', 'sun_bf_z', 'earth_bf_x', 'earth_bf_y', 'earth_bf_z', 'radius']
#     df = pd.read_csv(csv_path, header=None)
#     df.columns = header
    
#     # 軌道情報作成
#     solar_vector_list = List[otd.Vector3d]() # Solar Vectorの履歴
#     planet_vector_list = List[otd.Vector3d]() # Vector to Earthの履歴
#     radius_list = List[float]() # radiusの履歴
#     for index, row in df.iterrows():
#         solar_vector_list.Add(otd.Vector3d(row[1]/1000,row[2]/1000,row[3]/1000))
#         planet_vector_list.Add(otd.Vector3d(row[4]/1000,row[5]/1000,row[6]/1000))
#         radius_list.Add(row[7])
#     time_list = list(df['time'].values.astype(float))
#     time_list = otd.Dimension.DimensionalList[otd.Dimension.Time](time_list)
#     # 軌道情報更新
#     orbit.HrSunVecArray = solar_vector_list
#     orbit.HrPlanetVecArray = planet_vector_list
#     orbit.HrTimeArray = time_list
#     orbit.HrOrbitRadiusArray = radius_list
#     orbit.Update()
#     return orbit

# def get_orbit_name_from_caseset(case): #TODO 不要なので近い将来消す？
#     orbit_name = ''
#     rad_tasks = case.RadiationTasks
#     for rad_task in rad_tasks:
#         if orbit_name == '':
#             orbit_name = rad_task.OrbitName
#         elif orbit_name != rad_task.OrbitName:
#             print('MyWarning: RadiationTasksのOrbitが共通ではありません。')
#     return orbit_name

# def open_dwg(dwg_path): #TODO 不要なので近い将来消す
#     td = otd.ThermalDesktop()
#     dwg_path_obj = otd.Utility.RootedPathname(dwg_path)
#     td.ConnectConfig.DwgPathname = dwg_path_obj
#     td.Connect()
#     return td

# def update_symbol(case, symbol_csv_path): #TODO 不要なので近い将来消す
#     symbol_name_list = case.SymbolNames
#     symbol_value_list = case.SymbolValues
    
#     df = pd.read_csv(symbol_csv_path)
#     for index, row in df.iterrows():
#         if row['name'] in symbol_name_list:
#             index = symbol_name_list.index(row['name'])
#             symbol_value_list[index] = str(row['value'])
#         else:
#             # print("シンボルが見つかりませんでした、追加します。")
#             symbol_name_list.append(row['name'])
#             symbol_value_list.append(row['value'])
#     symbol_names = List[str]()
#     symbol_values = List[str]()
#     for i in range(len(symbol_name_list)):
#         symbol_names.Add(symbol_name_list[i])
#         symbol_values.Add(symbol_value_list[i])
#     # 更新
#     case.SymbolNames = symbol_names
#     case.SymbolValues = symbol_values
#     case.Update()
#     return case

# def update_orbit(case, orbit_name): #TODO 不要なので近い将来消す
#     # 軌道参照先変更
#     for i in range(len(case.RadiationTasks)):
#         case.RadiationTasks[i].OrbitName = orbit_name
#     case.Update()
#     return case

# def change_sav_name(case, sav_name): #TODO 不要なので近い将来消す
#     case.SindaOptions.SaveFilename = sav_name
#     case.Update()
#     return case

# def main(dwg_path, caseset_group_name, caseset_name, orbit_csv_path="", symbol_csv_path="", sav_name=""):  #TODO 不要なので近い将来消す
#     td = otd.ThermalDesktop()
#     dwg_path_obj = otd.Utility.RootedPathname(dwg_path)
#     td.ConnectConfig.DwgPathname = dwg_path_obj
#     td.Connect()
    
#     run_one_case(td, caseset_group_name, caseset_name, orbit_csv_path, symbol_csv_path, sav_name)
#     return
