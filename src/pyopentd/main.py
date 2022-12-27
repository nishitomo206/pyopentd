import sys
import clr
import numpy as np
import pandas as pd
from argparse import ArgumentParser

sys.path.append(
    "C:/Windows/Microsoft.NET/assembly/GAC_MSIL/OpenTDv62/v4.0_6.2.0.7__65e6d95ed5c2e178/"
)
clr.AddReference("OpenTDv62")
sys.path.append(
    "C:/Windows/Microsoft.NET/assembly/GAC_64/OpenTDv62.Results/v4.0_6.2.0.0__b62f614be6a1e14a/"
)
clr.AddReference("OpenTDv62.Results")

# from System import *
from System.Collections.Generic import List
import OpenTDv62 as otd
import System
from OpenTDv62 import Dimension


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
            original_object = Case(case_td)
            cases.append(
                [
                    case_td.GroupName,
                    case_td.Name,
                    original_object.get_orbit_name(),
                    original_object,
                ]
            )
        header = [
            "group_name",
            "caseset_name",
            "orbit_name",
            "original_object",
        ]
        df = pd.DataFrame(cases, columns=header)
        return df

    def get_caseset(self, caseset_name, group_name):
        return Case(self.GetCaseSet(caseset_name, group_name))

    def get_heatloads(self):
        # * アプライ先、センサー先のノードは1つであると仮定。
        # * ノードヒーターであると仮定。
        # TODO 色々なタイプのノードに対応させる。
        heatloads = self.GetHeatLoads()
        heatloads_list = []
        for heatload in heatloads:
            apply_node = None
            name = heatload.Name
            submodel = heatload.Submodel.ToString()
            handle = heatload.Handle
            value = heatload.Value
            value_exp = heatload.ValueExp
            transient_type = heatload.HeatLoadTransientType
            enabled_exp = heatload.EnabledExp.ToString()
            if heatload.ApplyConnections != []:
                node = self.GetNode(heatload.ApplyConnections[0].Handle)
                apply_node = f"{node.Submodel}.{node.Id}"
            if len(heatload.ApplyConnections) >= 2:
                print("MYWARNING: 1つのヒーターが2つ以上のノードに適用されています。")
            heatloads_list.append(
                [
                    name,
                    submodel,
                    handle,
                    apply_node,
                    transient_type,
                    value,
                    value_exp,
                    enabled_exp,
                    heatload,
                ]
            )
        header = [
            "Name",
            "submodel",
            "handle",
            "apply_node",
            "transient_type",
            "value",
            "value_exp",
            "enabled_exp",
            "original_object",
        ]
        df_heatloads = pd.DataFrame(heatloads_list, columns=header)
        return df_heatloads

    def get_heaters(self):
        # * アプライ先、センサー先のノードは1つであると仮定。
        # * ノードヒーターであると仮定。
        # TODO 色々なタイプのノードに対応させる。
        heaters = self.GetHeaters()
        heaters_list = []
        for heater in heaters:
            apply_node = None
            sensor_node = None
            name = heater.Name
            submodel = heater.Submodel
            handle = heater.Handle
            value = heater.Value
            value_exp = heater.ValueExp
            on_temp = heater.OnTemp
            on_temp_exp = heater.OnTempExp
            off_temp = heater.OffTemp
            off_temp_exp = heater.OffTempExp
            enabled_exp = heater.EnabledExp.ToString()
            if heater.ApplyConnections != []:
                node = self.GetNode(heater.ApplyConnections[0].Handle)
                apply_node = f"{node.Submodel}.{node.Id}"
            if len(heater.ApplyConnections) >= 2:
                print("MYWARNING: 1つのヒーターが2つ以上のノードに適用されています。")
            if heater.SensorConnections != []:
                node = self.GetNode(heater.SensorConnections[0].Handle)
                sensor_node = f"{node.Submodel}.{node.Id}"
            if len(heater.SensorConnections) >= 2:
                print("MYWARNING: 1つのヒーターが2つ以上のノードの温度を監視してます。")
            heaters_list.append(
                [
                    name,
                    submodel,
                    handle,
                    apply_node,
                    sensor_node,
                    value,
                    value_exp,
                    on_temp,
                    on_temp_exp,
                    off_temp,
                    off_temp_exp,
                    enabled_exp,
                    heater,
                ]
            )
        header = [
            "Name",
            "submodel",
            "handle",
            "apply_node",
            "sensor_node",
            "value",
            "value_exp",
            "on_temp",
            "on_temp_exp",
            "off_temp",
            "off_temp_exp",
            "enabled_exp",
            "original_object",
        ]
        df_heaters = pd.DataFrame(heaters_list, columns=header)
        return df_heaters

    def get_heater(self, handle):
        return self.GetHeater(handle)

    def get_orbits(td):
        orbits_list = []
        orbits = td.GetOrbits()
        for orbit in orbits:
            orbits_list.append([orbit.Name, orbit.OrbitType, orbit])
        df_orbits = pd.DataFrame(
            orbits_list, columns=["Name", "OrbitType", "original_object"]
        )
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
        df_sun = pd.DataFrame(sun_vecs, columns=["sun_x", "sun_y", "sun_z"])
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
        df_planet = (
            pd.DataFrame(
                planet_vecs, columns=["planet_x", "planet_y", "planet_z"]
            )
            * 1000
        )
        if df_planet.abs().max().max() < 0.0011:
            df_planet *= 1000
        if df_planet.abs().max().max() > 1.1:
            df_planet /= 1000
        # 時間列
        times = []
        for HrTime in orbit.HrTimeArray:
            t = HrTime.GetValueSI()
            times.append(t)
        df_times = pd.DataFrame(times, columns=["Times"])
        # 距離時系列
        radiuses = orbit.HrOrbitRadiusArray
        df_radiuses = pd.DataFrame(radiuses, columns=["radius"])
        return pd.concat([df_times, df_sun, df_planet, df_radiuses], axis=1)

    def get_nodes(self):
        nodes = self.GetNodes()
        nodes_list = []
        for node in nodes:
            nodes_list.append([node.Submodel, node.Id, node.Handle, node])
        header = ["submodel", "id", "handle", "original_object"]
        df_nodes = pd.DataFrame(nodes_list, columns=header)
        return df_nodes

    def get_node(self, submodel, id, printif=False):
        nodes = self.GetNodes()
        node = 0
        for node_i in nodes:
            if node_i.Submodel.ToString() == submodel and node_i.Id == int(id):
                node = self.GetNode(node_i.Handle)
                break
        if node == 0:
            print("MYERROR (in pyopentd.main.get_node): 指定したnodeが見つかりませんでした。")
            return
        if printif:
            print(node)
            print("Handle: ", node.Handle)
            print("Submodel: ", node.Submodel)
            print("Id: ", node.Id)
        return node

    # TODO: get_solidsやget_thin_shellsでまとめる関数があるとよさそう。現状抽出してる項目同じなので。

    # TODO: 抽出する項目は要検討
    def get_solid_bricks(self):
        solid_bricks = self.GetSolidBricks()
        solid_bricks_list = []
        for solid_brick in solid_bricks:
            solid_bricks_list.append(
                [
                    solid_brick.StartSubmodel,
                    solid_brick.CondSubmodel,
                    solid_brick.NodeNames,
                    solid_brick.AttachedNodeHandles,
                    solid_brick.OutsideOpticalProperties,
                    solid_brick.Handle,
                    solid_brick.Comment,
                    solid_brick.XMax,
                    solid_brick.YMax,
                    solid_brick.ZMax,
                    solid_brick,
                ]
            )
        header = [
            "start_submodel",
            "cond_submodel",
            "node_names",
            "attached_node_handles",
            "outside_optical_properties",
            "handle",
            "comment",
            "x_max",
            "y_mMax",
            "z_mMax",
            "original_object",
        ]
        df_solid_bricks = pd.DataFrame(solid_bricks_list, columns=header)
        return df_solid_bricks

    # TODO: 抽出する項目は要検討
    def get_solid_cylinders(self):
        solid_cylinders = self.GetSolidCylinders()
        solid_cylinders_list = []
        for solid_cylinder in solid_cylinders:
            solid_cylinders_list.append(
                [
                    solid_cylinder.StartSubmodel,
                    solid_cylinder.CondSubmodel,
                    solid_cylinder.NodeNames,
                    solid_cylinder.AttachedNodeHandles,
                    solid_cylinder.OutsideOpticalProperties,
                    solid_cylinder.Handle,
                    solid_cylinder.Comment,
                    solid_cylinder.Height,
                    solid_cylinder.Rmin,
                    solid_cylinder.Rmax,
                    solid_cylinder.StartAngle,
                    solid_cylinder.EndAngle,
                    solid_cylinder,
                ]
            )
        header = [
            "start_submodel",
            "cond_submodel",
            "node_names",
            "attached_node_handles",
            "outside_optical_properties",
            "handle",
            "comment",
            "height",
            "r_min",
            "r_max",
            "start_angle",
            "end_angle",
            "original_object",
        ]
        df_solid_cylinders = pd.DataFrame(solid_cylinders_list, columns=header)
        return df_solid_cylinders

    # TODO: 抽出する項目は要検討
    def get_rectangles(self):
        rectangles = self.GetRectangles()
        rectangles_list = []
        for rectangle in rectangles:
            rectangles_list.append(
                [
                    rectangle.TopStartSubmodel,
                    rectangle.BotStartSubmodel,
                    rectangle.CondSubmodel,
                    rectangle.TopNodeNames,
                    rectangle.BotNodeNames,
                    rectangle._AttachedNodeHandles,
                    rectangle.TopOpticalProp,
                    rectangle.BotOpticalProp,
                    rectangle.Handle,
                    rectangle.Comment,
                    rectangle.XMax,
                    rectangle.YMax,
                    rectangle.TopThickness,
                    rectangle.BotThickness,
                    rectangle,
                ]
            )
        header = [
            "top_start_submodel",
            "bot_start_submodel",
            "cond_submodel",
            "top_node_names",
            "bot_node_names",
            "_attached_node_handles",
            "top_optical_prop",
            "bot_optical_prop",
            "handle",
            "comment",
            "x_max",
            "y_max",
            "top_thickness",
            "bot_thickness",
            "original_object",
        ]
        df_rectangles = pd.DataFrame(rectangles_list, columns=header)
        return df_rectangles

    # TODO: 抽出する項目は要検討
    def get_polygons(self):
        polygons = self.GetPolygons()
        polygons_list = []
        for polygon in polygons:
            polygons_list.append(
                [
                    polygon.TopStartSubmodel,
                    polygon.BotStartSubmodel,
                    polygon.CondSubmodel,
                    polygon.TopNodeNames,
                    polygon.BotNodeNames,
                    polygon._AttachedNodeHandles,
                    polygon.TopOpticalProp,
                    polygon.BotOpticalProp,
                    polygon.Handle,
                    polygon.Comment,
                    polygon.TopThickness,
                    polygon.BotThickness,
                    polygon,
                ]
            )
        header = [
            "top_start_submodel",
            "bot_start_submodel",
            "cond_submodel",
            "top_node_names",
            "bot_node_names",
            "_attached_node_handles",
            "top_optical_prop",
            "bot_optical_prop",
            "handle",
            "comment",
            "top_thickness",
            "bot_thickness",
            "original_object",
        ]
        df_polygons = pd.DataFrame(polygons_list, columns=header)
        return df_polygons

    # TODO: 抽出する項目は要検討
    def get_cylinders(self):
        cylinders = self.GetCylinders()
        cylinders_list = []
        for cylinder in cylinders:
            cylinders_list.append(
                [
                    cylinder.TopStartSubmodel,
                    cylinder.BotStartSubmodel,
                    cylinder.CondSubmodel,
                    cylinder.TopNodeNames,
                    cylinder.BotNodeNames,
                    cylinder._AttachedNodeHandles,
                    cylinder.TopOpticalProp,
                    cylinder.BotOpticalProp,
                    cylinder.Handle,
                    cylinder.Comment,
                    cylinder.Height,
                    cylinder.Radius,
                    cylinder.StartAngle,
                    cylinder.EndAngle,
                    cylinder,
                ]
            )
        header = [
            "top_start_submodel",
            "bot_start_submodel",
            "cond_submodel",
            "top_node_names",
            "bot_node_names",
            "_attached_node_handles",
            "top_optical_prop",
            "bot_optical_prop",
            "handle",
            "comment",
            "height",
            "radius",
            "start_angle",
            "end_angle",
            "original_object",
        ]
        df_cylinders = pd.DataFrame(cylinders_list, columns=header)
        return df_cylinders

    def create_caseset(
        self,
        name,
        group,
        steady,
        transient,
        time_end=0,
        run_dir=None,
        sumodels_not_built=[],
        restart_file=None,
        force_reset=False,
    ):
        # 既にケースセットがあるかの確認
        if self.GetCaseSet(name, group):
            if force_reset:
                self.DeleteCaseSet(name, group)
            else:
                while True:
                    choice = input(
                        f"既に {name} が {group} にありますが、削除して新たにケースセットを作成しますか？ [y/N]: "
                    ).lower()
                    if choice in ["y", "ye", "yes"]:
                        self.DeleteCaseSet(name, group)
                        break
                    elif choice in ["n", "no"]:
                        print(
                            f"Error: 既に{name} が {group} にあります", file=sys.stderr
                        )
                        sys.exit(1)
        # 以下、ケースセットの作成
        case = self.CreateCaseSet(name, group, name)
        case = Case(case)
        case.origin.SaveQ = 1  # outputにincident heatを含める。
        case.origin.SteadyState = steady
        case.origin.Transient = transient
        if transient == 1:
            if time_end == 0:
                print("MYERROR: 終了時刻(time_end)を指定してください。")
            case.origin.SindaControl.timendExp.Value = str(time_end)

        # リスタートファイルを使用する
        if restart_file != "" and restart_file != None:
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
        if run_dir != "" and run_dir != None:
            case.origin.UseUserDirectory = 1
            case.origin.UserDirectory = run_dir
        case.update()
        return case

    def create_heatload(
        self,
        submodel,
        apply_node_sub,
        apply_node_id,
        transient_type,
        value=-1,
        time_array=[],
        value_array=[],
        name="",
        layer="",
        enable_exp="",
    ):
        node = self.get_node(apply_node_sub, apply_node_id)  # ノードの取得
        heatload = self.CreateHeatLoad(otd.Connection(node))  # heatloadの作成

        if transient_type == 0:  # constant heatload の作成
            # typeの変更
            heatload.HeatLoadTransientType = 0
            heatload.TimeDependentSteadyStateType = 0
            if value == -1:
                print(
                    "MYWARNING (in pyopentd.main.create_heatload): ヒートロードのvalueが指定されていません。"
                )
            # valueの変更
            if type(value) == str:
                heatload.ValueExp.Value = value
            else:
                heatload.Value = value
        elif transient_type == 1:  # time dependent heatload の作成
            # チェック
            if len(time_array) != len(value_array):
                print(
                    "MYERROR (in pyopentd.main.create_heatload): time_arrayとvalue_arrayの長さが異なります。"
                )
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
            print(
                "MYERROR (in pyopentd.main.create_heatload): transient_typeが正しくありません。"
            )

        heatload.Submodel.Name = submodel  # hwatloadのSubmodelを変更
        if name != "":
            heatload.Name = name
        if layer != "":
            heatload.Layer = layer
        if enable_exp != "":
            heatload.EnabledExp.Value = enable_exp

        heatload.Update()
        return heatload

    def create_heater(
        self,
        apply_node_sub,
        apply_node_id,
        sensor_node_sub,
        sensor_node_id,
        value,
        on_temp,
        off_temp,
        name="",
        submodel="MAIN",
        enabled_exp="",
        layer=None,
        ss_method=None,
        ss_power_per=None,
        ss_proportional_damp=None,
        time_list=None,
        scale_list=None,
    ):
        # ヒーターの作成
        apply_node = self.get_node(apply_node_sub, apply_node_id)  # ノードの取得
        apply_list = List[otd.Connection]()
        apply_list.Add(otd.Connection(apply_node))
        sensor_node = self.get_node(sensor_node_sub, sensor_node_id)  # ノードの取得
        sensor_list = List[otd.Connection]()
        sensor_list.Add(otd.Connection(sensor_node))
        heater = self.CreateHeater(apply_list, sensor_list)

        # ヒーターのプロパティ設定
        heater.ValueExp.Value = str(value)
        heater.OnTempExp.Value = str(on_temp)
        heater.OffTempExp.Value = str(off_temp)
        heater.Name = name
        heater.Submodel.Name = submodel
        heater.EnabledExp.Value = enabled_exp
        if layer != None:
            heater.Layer = layer

        # TODO ss_methodの指定
        heater.SSMethod = 1
        heater.SSPowerPer = 0

        # times, scalesの指定
        if time_list != None or scale_list != None:
            heater.UseTransientScaling = 1
            times = Dimension.DimensionalList[Dimension.Time](time_list)
            scales = List[float]()
            for scale in scale_list:
                scales.Add(scale)
            heater.Times = times
            heater.Scales = scales

        heater.Update()
        return heater

    def get_symbol_names(self):
        symbols_list = []
        for symbol in self.GetSymbols():
            symbols_list.append(symbol.Name)
        return symbols_list

    def get_symbols(self):
        symbols = []
        for symbol in self.GetSymbols():
            symbols.append(
                [
                    symbol.Name,
                    symbol.Value,
                    symbol.Description,
                    symbol.Group,
                    symbol.Type,
                ]
            )
        df = pd.DataFrame(
            symbols, columns=["name", "value", "description", "group", "type"]
        )
        return df

    def get_symbols_each_case(self, group_name=None, output_type=None):
        """各ケースセットのシンボルの一括取得

        Args:
            output_type (int): アウトプットのタイプを指定。（現状はケースセットで指定されている変数のみを表示）
        """
        cases = self.get_casesets()
        df_output = pd.DataFrame([])
        for i in range(cases.shape[0]):
            row = cases.loc[i, :]
            case = cases.loc[i, "original_object"]
            if not group_name is None:
                if row["group_name"] != group_name:
                    continue
            df_symbol = case.get_symbols()
            ds_symbol = df_symbol.set_index("name")["value"]
            ds_symbol.index.name = None
            ds_symbol.name = f'{row["group_name"]}.{row["caseset_name"]}'
            df_output = pd.concat([df_output, ds_symbol], axis=1)
        df_output = df_output.sort_index()
        df_output = df_output.T
        return df_output

    def add_symbol(self, name, value, group="td_tool"):
        symbol_names_list = self.get_symbol_names()
        if name in symbol_names_list:
            print(f"MYWARNING: {name}は既にグローバルで宣言されています。スキップします。")
        else:
            symbol = self.CreateSymbol(name, str(value))
            symbol.Group = group
            symbol.Update()

    def create_orbit(
        self, df_orbit, orbit_name="new_orbit", solar_flux=None, albedo=None
    ):
        """新規軌道作成

        Args:
            df_orbit (pandas.core.frame.DataFrame): 軌道のDataframe(カラムは['time', 'sun_x', 'sun_y', 'sun_z', 'earth_x', 'earth_y', 'earth_z', 'radius'])
            orbit_name (str): 軌道の名前
        """
        # カラム名チェック
        column = [
            "Times",
            "sun_x",
            "sun_y",
            "sun_z",
            "planet_x",
            "planet_y",
            "planet_z",
            "radius",
        ]
        if df_orbit.columns.values.tolist() != column:
            print(
                "MYERROR (in pyopentd.main.create_orbit): DataFrameのカラム名が正しくありません。['Times', 'sun_x', 'sun_y', 'sun_z', 'planet_x', 'planet_y', 'planet_z', 'radius']に変更して下さい。"
            )

        # 新しい軌道の作成
        orbit = self.CreateOrbit(orbit_name)
        orbit.OrbitType = (
            otd.RadCAD.Orbit.OrbitTypes.TRAJECTORY
        )  # orbitタイプの変更。

        # 軌道情報作成
        solar_vector_list = List[otd.Vector3d]()  # Solar Vectorの履歴
        planet_vector_list = List[otd.Vector3d]()  # Vector to Earthの履歴
        radius_list = List[float]()  # radiusの履歴
        for index, row in df_orbit.iterrows():
            solar_vector_list.Add(
                otd.Vector3d(
                    row["sun_x"] / 1000,
                    row["sun_y"] / 1000,
                    row["sun_z"] / 1000,
                )
            )
            planet_vector_list.Add(
                otd.Vector3d(
                    row["planet_x"] / 1000,
                    row["planet_y"] / 1000,
                    row["planet_z"] / 1000,
                )
            )
            radius_list.Add(row["radius"])
        time_list = list(df_orbit["Times"].values.astype(float))
        time_list = otd.Dimension.DimensionalList[otd.Dimension.Time](
            time_list
        )
        # 軌道情報更新
        orbit.HrSunVecArray = solar_vector_list
        orbit.HrPlanetVecArray = planet_vector_list
        orbit.HrTimeArray = time_list
        orbit.HrOrbitRadiusArray = radius_list
        if not solar_flux is None:
            orbit.SolarFluxExp.Value = str(solar_flux)
        if not albedo is None:
            orbit.AlbedoExp.Value = str(albedo)
        orbit.Update()
        return orbit


class Case:
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
        orbit_name = ""
        rad_tasks = self.origin.RadiationTasks
        for rad_task in rad_tasks:
            if orbit_name == "":
                orbit_name = rad_task.OrbitName
            elif orbit_name != rad_task.OrbitName:
                print("MyWarning: RadiationTasksのOrbitが共通ではありません。")
        return orbit_name

    def get_symbols(self):
        symbols = np.array(
            [
                self.origin.SymbolNames,
                self.origin.SymbolValues,
                self.origin.SymbolComments,
            ]
        ).T
        df = pd.DataFrame(symbols, columns=["name", "value", "comment"])
        return df

    def update_symbols(self, td, df_symbol, reset_symbols=False):
        """複数変数の追加・変更

        複数変数の追加・変更用のメソッド。1つの変数ならadd_symbolのほうが引数が分かりやすい（やってることは同じ）。多くの変数を追加するときに、add_symbolだと時間がかかりそうなので、このメソッドを作成。

        Args:
            df_symbol (pandas.core.frame.DataFrame): 変更したい変数のDataframe('name'と'value'をカラムにもつ)
            reset_symbols (bool): 現在の変数を全てリセットしてから、変数を更新する。
        """

        global_symbol_names_list = td.get_symbol_names()

        if reset_symbols:
            symbol_name_list = []
            symbol_value_list = []
            symbol_comment_list = []
        else:
            symbol_name_list = self.origin.SymbolNames
            symbol_value_list = self.origin.SymbolValues
            symbol_comment_list = self.origin.SymbolComments

        for index, row in df_symbol.iterrows():
            if row["name"] in symbol_name_list:
                index = symbol_name_list.index(row["name"])
                symbol_value_list[index] = str(row["value"])
            else:
                # print("シンボルが見つかりませんでした、追加します。")
                symbol_name_list.append(row["name"])
                symbol_value_list.append(str(row["value"]))
                symbol_comment_list.append("")
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

    def add_symbol(self, name, value):
        """1変数の追加・変更

        多くの変数を追加・変更する場合はupdate_symbolsの使用を検討してください。

        """
        symbol_name_list = self.origin.SymbolNames
        symbol_value_list = self.origin.SymbolValues
        symbol_comment_list = self.origin.SymbolComments
        if name in symbol_name_list:
            index = symbol_name_list.index(name)
            symbol_value_list[index] = str(value)
        else:
            # print("シンボルが見つかりませんでした、追加します。")
            symbol_name_list.append(name)
            symbol_value_list.append(str(value))
            symbol_comment_list.append("")
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

    def update_orbit(self, orbit_name):  # TODO ここの入力をorbitで対応できるようにする。
        # 軌道参照先変更
        for i in range(len(self.origin.RadiationTasks)):
            self.origin.RadiationTasks[i].OrbitName = orbit_name
        self.update()
        return

    def change_sav_name(self, sav_name):
        self.origin.SindaOptions.SaveFilename = sav_name
        self.update()
        return

    def add_radiation_task(
        self, calc_type, orbit_name="", analysis_group="BASE"
    ):
        # rad task の作成
        rad_task = otd.RadiationTaskData()
        rad_task.AnalGroup = analysis_group  # Analysis Groupの指定
        rad_task.TypeCalc = (
            calc_type  # [0: Radks, 1: Heating Rates, 2: Articulating Radks]
        )
        if calc_type == 1 or calc_type == 2:
            rad_task.OrbitName = orbit_name

        rad_task.RkFilename = (
            f"{self.origin.GroupName}_{self.origin.Name}_{analysis_group}.k"
        )
        submodel_name = (
            f"{self.origin.GroupName}_{self.origin.Name}_{analysis_group}"
        )
        n_rad_task = len(self.origin.RadiationTasks)
        submodel_name = (
            submodel_name[:29] + f"_{n_rad_task}"
        )  # submodel nameは32文字までしか入らないので、制限をかける。
        rad_task.RkSubmodel = submodel_name
        rad_task.HrFilename = (
            f"{self.origin.GroupName}_{self.origin.Name}_{analysis_group}.hr"
        )
        rad_task.HrSubmodel = submodel_name
        rad_task.FfFilename = (
            f"{self.origin.GroupName}_{self.origin.Name}_{analysis_group}.dat"
        )
        rad_task.OutputTrackerDataFile = (
            f"{self.origin.GroupName}_{self.origin.Name}_{analysis_group}.dat"
        )

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
