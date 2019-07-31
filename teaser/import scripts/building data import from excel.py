# -*- coding: utf-8 -*-
# @Author: Martin Rätz
# @Date:   2019-02-19 18:41:56
# @Last Modified by: Martin Rätz
# @Last Modified time:

"""
Limitations and assumptions:
- Outer and inner wall area depend on the calculations done in the excel
- Ground floor area is only as big the respective net area of the heated room volume (NetArea)
- Floor area is only as big the respective net area of the heated room volume (NetArea)
- Rooftop area is only as big the respective net area of the heated room volume (NetArea)
- Rooftops are flat and not tilted, see "RooftopTilt"
- Ceiling area is only as big the respective net area of the heated room volume (NetArea)
- Ceiling, floor and inner walls are only respected by half their area, since they belong half to the respective
and half to the adjacent zone

-respective construction types have to be added to the TypeBuildingElements.xml
-respective UsageTypes for Zones have to be added to the UseConditions.xml
-excel file format has to be as shown in the "Excel_Import_Example_File.xlsx" #TODO: Set this file
-yellowed columns are input to teaser -> don´t change column header, keep value names consistent. Additional columns
may be used for Zoning
"""

import os
from teaser.project import Project
from teaser.logic.buildingobjects.building import Building
from teaser.logic.buildingobjects.thermalzone import ThermalZone
from teaser.logic.buildingobjects.boundaryconditions.boundaryconditions \
    import BoundaryConditions
from teaser.logic.buildingobjects.buildingphysics.outerwall import OuterWall
from teaser.logic.buildingobjects.buildingphysics.floor import Floor
from teaser.logic.buildingobjects.buildingphysics.rooftop import Rooftop
from teaser.logic.buildingobjects.buildingphysics.groundfloor\
    import GroundFloor
from teaser.logic.buildingobjects.buildingphysics.ceiling import Ceiling
from teaser.logic.buildingobjects.buildingphysics.window import Window
from teaser.logic.buildingobjects.buildingphysics.innerwall import InnerWall
import pandas as pd
import numpy as np
import warnings

def import_data(FileName=None, SheetNames=[0, 1, 2, 3, 4]):
    ''' Import data from the building data excel file and perform some preprocessing for nan and empty cells.
    If several sheets are imported, the data is concatenated to one dataframe'''

    Path = os.path.join(
        os.path.dirname(os.path.dirname(__file__)),
        'data',
        'input',
        'inputdata',
        'buildingdata',
        FileName)

    # process an import of a single sheet as well as several sheets, which will be concatenated with an continuous index
    if type(SheetNames) == list:
        Data = pd.DataFrame()
        _Data = pd.read_excel(io=Path, sheet_name=SheetNames, header=0, index_col=None)
        for sheet in SheetNames:
            Data = Data.append(_Data[sheet], sort=False)
        Data = Data.reset_index(drop=True)
    else:
        Data = pd.read_excel(io=Path, sheet_name=SheetNames, header=0, index_col=None)

    # Cut of leading or tailing white spaces from any string in the dataframe
    Data = Data.applymap(lambda x: x.strip() if type(x) is str else x)

    # Convert every N/A, nan, empty strings and strings called N/a, n/A, NAN, nan, na, Na, nA or NA to np.nan
    Data = Data.replace(["", "N/a", "n/A", "NAN", "nan", "na", "Na", "nA", "NA"], np.nan, regex=True)
    Data = Data.fillna(np.nan)

    return Data

def get_list_of_present_entries(List):
    ''' Extracts a list of all in the list available entries,
        discarding "None" and "nan" entries
    '''
    _List = []
    for x in List:
        if not x in _List:
            if not None:
                if not pd.isna(x):
                    _List.append(x)
    return _List

# Block: Zoning methodologies (define your zoning function here)
# -------------------------------------------------------------
def zoning_1(Data):
    '''
    -UsageType has to be empty in the case that the respective line does not represent an other room but a different
    orientated  wall or window belonging to a room that is already declared once in the excel file
    '''
    # just apply some values since the excel sheet is not yet filled with good values
    Data["WindowConstruction"] = "EnEv"
    Data["OuterWallConstruction"] = "heavy"
    Data["InnerWallConstruction"] = "heavy"
    Data["FloorConstruction"] = "heavy"
    Data["GroundFloorConstruction"] = "heavy"
    Data["CeilingConstruction"] = "heavy"
    Data["RooftopConstruction"] = "heavy"

    # account all outer walls not adjacent to the ambient to the entity "inner wall"
    # !right now the wall construction of the added wall is not respected, the same wall construction as regular inner wall is set
    for index, line in Data.iterrows():
        if not pd.isna(line["WallAdjacentTo"]):
            Data.at[index, "InnerWallArea[m²]"] = Data.at[index, "OuterWallArea[m²]"]\
                                                  + Data.at[index, "WindowArea[m²]"] \
                                                  + Data.at[index, "InnerWallArea[m²]"]
            Data.at[index, "WindowOrientation[°]"] = np.NaN
            Data.at[index, "WindowArea[m²]"] = np.NaN
            Data.at[index, "WindowConstruction"] = np.NaN
            Data.at[index, "OuterWallOrientation[°]"] = np.NaN
            Data.at[index, "OuterWallArea[m²]"] = np.NaN
            Data.at[index, "OuterWallConstruction"] = np.NaN

    # make all rooms that belong to a certain room have the same room identifier
    _list = []
    for index, line in Data.iterrows():
        if pd.isna(line["BelongsToIdentifier"]):
            _list.append(line["RoomIdentifier"])
        else:
            _list.append(line["BelongsToIdentifier"])
    Data["RoomCluster"] = _list

    # check for lines in which the net area is zero, marking an second wall or window
    # element for the respective room, and in which there is still stated a UsageType which is wrong
    # and should be changed in the file
    for i, row in Data.iterrows():
        if row["NetArea[m²]"] == 0 and not pd.isna(row["UsageType"]):
            warnings.warn("In line %s the net area is zero, marking an second wall or window element for the respective room, "
                  "and in which there is still stated a UsageType which is wrong and should be changed in the file"
                  %i)

    # make all rooms of the cluster having the usage type of the main usage type
    _groups = Data.groupby(["RoomCluster"])
    for index, cluster in _groups:
        count = 0
        for line in cluster.iterrows():
            if pd.isna(line[1]["BelongsToIdentifier"]) and not pd.isna(line[1]["UsageType"]):
                main_usage = line[1]["UsageType"]
                for i, row in Data.iterrows():
                    if row["RoomCluster"] == line[1]["RoomCluster"]:
                        Data.at[i, "RoomClusterUsage"] = main_usage
                count += 1
        if count != 1:
            warnings.warn("This cluster has more than one main usage type or none, check your excel file for mistakes! \n"
                          "Common mistakes: \n"
                          "-NetArea of a wall is not equal to 0 \n"
                          "-UsageType of a wall is not empty \n"
                          "Explanation: Rooms may have outer walls/windows on different orientations. That is why, "
                          "we denote an empty slot under UsageType, marks another direction of an outer wall or window entity "
                          "of the same room. The connection of the same room is defined by an equal RoomIdentifier."
                            "Cluster = %s" %cluster)



    # name usage types after usage types available in the xml
    #Todo: For now just done anything, check for meaning later!
    usage_to_xml_usage = {"IsolationRoom" : "Bed room",
                          "Aisle" : "Corridors in the general care area",
                          "Technical room" : "Stock, technical equipment, archives", # only correct if technical rooms have no devices producing heat
                          "Washing" : "WC and sanitary rooms in non-residential buildings",
                          "Stairway" : "Corridors in the general care area",
                          "WC" : "WC and sanitary rooms in non-residential buildings",
                          "Storage" : "Stock, technical equipment, archives",
                          "Lounge" : "Meeting, Conference, seminar",
                          "Office": "Meeting, Conference, seminar",
                          "Treatment room" : "Examination- or treatment room",
                          "StorageChemical" : "Stock, technical equipment, archives",
                          "EquipmentServiceAndRinse" : "WC and sanitary rooms in non-residential buildings",}

    # rename all zone names from the excel to the according zone name which is in the BoundaryConditions.xml files
    usages = get_list_of_present_entries(Data["RoomClusterUsage"])
    for usage in usages:
        Data["RoomClusterUsage"] = np.where(Data["RoomClusterUsage"] == usage, usage_to_xml_usage[usage], Data["RoomClusterUsage"])

    # name the column where the zones are defined "Zone"
    Data["Zone"] = Data["RoomClusterUsage"]

    return Data
# -------------------------------------------------------------

def import_building_from_excel(prj, building_name, construction_age, excel_name, SheetNames):
    '''Import building data from excel, convert it via the respective zoning and feed it to teasers logic classes.'''

    bldg = Building(parent=prj)
    bldg.name = building_name
    bldg.year_of_construction = construction_age
    bldg.with_ahu = True #Todo: avoid hard coding, import this from the excel sheet!?
    if bldg.with_ahu is True:
        #profile by Mans yearly don´t work for current development version of aixlib
        #profile_temperature_ahu_summer = 7 * [293.15] + 12 * [287.15] + 5 * [
        #    293.15]
        #profile_temperature_ahu_winter = 24 * [299.15]
        #bldg.central_ahu.profile_temperature = 120 * \
        #                                       profile_temperature_ahu_winter \
        #                                       + 124 * \
        #                                       profile_temperature_ahu_summer \
        #                                       + 121 * profile_temperature_ahu_winter
        #bldg.central_ahu.profile_min_relative_humidity = (8760 * [0.45])
        #bldg.central_ahu.profile_max_relative_humidity = (8760 * [0.65])
        #bldg.central_ahu.profile_v_flow = (
        #                                          7 * [0.0] + 12 * [1.0] + 5 * [0.0]) * 365

        #literature values
        #bldg.central_ahu.profile_temperature = (
        #    7 * [293.15] + 12 * [295.15] + 6 * [293.15])
        ## according to :cite:`DeutschesInstitutfurNormung.2016`
        #bldg.central_ahu.profile_min_relative_humidity = (25 * [0.45])
        ##  according to :cite:`DeutschesInstitutfurNormung.2016b`  and
        ## :cite:`DeutschesInstitutfurNormung.2016`
        #bldg.central_ahu.profile_max_relative_humidity = (25 * [0.65])
        #bldg.central_ahu.profile_v_flow = (
        #    7 * [0.0] + 12 * [1.0] + 6 * [0.0])
        ## according to user
        ## profile in :cite:`DeutschesInstitutfurNormung.2016`

        #HUS values
        bldg.central_ahu.heat_recovery = True
        bldg.central_ahu.efficiency_recovery = 0.35 # Trox mentioned 0,73 but due to a dynamic behavior i assume 0.6
        bldg.central_ahu.profile_temperature = (25 * [273.15 + 18])
        bldg.central_ahu.profile_min_relative_humidity = (25 * [0])
        bldg.central_ahu.profile_max_relative_humidity = (25 * [1])
        bldg.central_ahu.profile_v_flow = (25 * [1])


    #Parameters to be hard coded in teasers logic classes
    #1. "use_set_back" needs hard coding at aixlib.py in the init; defines if the in the useconditions stated heating_time with the
    #   respective set_back_temp should be applied. use_set_back = false -> all hours of the day have same set_temp_heat
    #   actual value: use_set_back = false
    #2. HeaterOn, CoolerOn, hHeat, lCool, etc. can be hard coded in the text file
    #   "teaser / data / output / modelicatemplate / AixLib / AixLib_ThermalZoneRecord_TwoElement"
    #   actual changes:
    #           hHeat = ${zone.model_attr.heat_load}, to hHeat = 100000000000,

    #Parameters to be set for each and every zone
    #-----------------------------
    OutWallTilt = 90
    WindowTilt = 90
    GroundFloorTilt = 0
    FloorTilt = 0
    CeilingTilt = 0
    RooftopTilt = 0
    GroundFloorOrientation = -2
    FloorOrientation = -2
    RooftopOrientation = -1
    CeilingOrientation = -1
    #-----------------------------

    # load_building_data from excel_to_pandas dataframe:
    Data = import_data(excel_name, SheetNames)

    # informative print
    UsageTypes = get_list_of_present_entries(Data["UsageType"])
    print("List of present UsageTypes in the original Data set: \n%s" %UsageTypes)

    # define the zoning methodology/function
    Data = zoning_1(Data)

    # informative print
    UsageTypes = get_list_of_present_entries(Data["Zone"])
    print("List of Zones after the zoning is applied: \n%s" %UsageTypes)

    # aggregate all rooms of each Zone and for each set general parameter, boundary conditions
    # and parameter regarding the building physics
    Zones = Data.groupby(["Zone"])
    for name, Zone in Zones:

        # Block: Thermal Zone (general parameter)
        tz = ThermalZone(parent=bldg)
        tz.name = str(name)
        tz.area = Zone["NetArea[m²]"].sum()
        # room vice calculation of volume plus summing those
        tz.volume = (np.array(Zone["NetArea[m²]"]) * np.array(Zone["HeatedRoomHeight[m]"])).sum()

        # Block: Boundary Conditions
        # load UsageOperationTime, Lighting, RoomClimate and InternalGains from the "UseCondition.xml"
        tz.use_conditions = BoundaryConditions(parent=tz)
        tz.use_conditions.load_use_conditions(Zone["Zone"].iloc[0], prj.data)

        # Block: Building Physics
        # Grouping by orientation and construction type
        # aggregating and feeding to the teaser logic classes
        Grouped = Zone.groupby(["OuterWallOrientation[°]", "OuterWallConstruction"])
        for name, group in Grouped:
            # looping through a groupby object automatically discards the groups where one of the attributes is nan
            # additionally check for strings, since the value must be of type int or float
            if not isinstance(group["OuterWallOrientation[°]"].iloc[0], str):
                out_wall = OuterWall(parent=tz)
                out_wall.name = "outer_wall_" + \
                                str(group["OuterWallOrientation[°]"].iloc[0]) + \
                                str(group["OuterWallConstruction"].iloc[0])
                out_wall.area = group["OuterWallArea[m²]"].sum()
                out_wall.tilt = OutWallTilt
                out_wall.orientation = group["OuterWallOrientation[°]"].iloc[0]
                # load wall properties from "TypeBuildingElements.xml"
                out_wall.load_type_element(
                    year=bldg.year_of_construction,
                    construction=group["OuterWallConstruction"].iloc[0])
            else:
                warnings.warn(
                    "In Zone \"%s\" the OuterWallOrientation \"%s\" is neither float nor int, "
                    "hence this building element is not added."
                    % (group["Zone"].iloc[0], group["OuterWallOrientation[°]"].iloc[0]))

        Grouped = Zone.groupby(["WindowOrientation[°]", "WindowConstruction"])
        for name, group in Grouped:
            # looping through a groupby object automatically discards the groups where one of the attributes is nan
            # additionally check for strings, since the value must be of type int or float
            if not isinstance(group["OuterWallOrientation[°]"].iloc[0], str):
                window = Window(parent=tz)
                window.name = "window_" + \
                              str(group["WindowOrientation[°]"].iloc[0]) + \
                              str(group["WindowConstruction"].iloc[0])
                window.area = group["WindowArea[m²]"].sum()
                window.tilt = WindowTilt
                window.orientation = group["WindowOrientation[°]"].iloc[0]
                # load wall properties from "TypeBuildingElements.xml"
                window.load_type_element(
                    year=bldg.year_of_construction,
                    construction=group["WindowConstruction"].iloc[0])
            else:
                warnings.warn(
                    "In Zone \"%s\" the window orientation \"%s\" is neither float nor int, "
                    "hence this building element is not added."
                    % (group["Zone"].iloc[0], group["WindowOrientation[°]"].iloc[0]))

        Grouped = Zone.groupby(["IsGroundFloor", "FloorConstruction"])
        for name, group in Grouped:
            if group["NetArea[m²]"].sum() != 0: # to avoid devision by 0
                if group["IsGroundFloor"].iloc[0] == 1:
                    ground_floor = GroundFloor(parent=tz)
                    ground_floor.name = "ground_floor" + \
                                    str(group["FloorConstruction"].iloc[0])
                    ground_floor.area = group["NetArea[m²]"].sum()
                    ground_floor.tilt = GroundFloorTilt
                    ground_floor.orientation = GroundFloorOrientation
                    # load wall properties from "TypeBuildingElements.xml"
                    ground_floor.load_type_element(
                        year=bldg.year_of_construction,
                        construction=group["FloorConstruction"].iloc[0])
                elif group["IsGroundFloor"].iloc[0] == 0:
                    floor = Floor(parent=tz)
                    floor.name = "floor" + \
                                    str(group["FloorConstruction"].iloc[0])
                    floor.area = group["NetArea[m²]"].sum() / 2 # only half of the floor belongs to this story
                    floor.tilt = FloorTilt
                    floor.orientation = FloorOrientation
                    # load wall properties from "TypeBuildingElements.xml"
                    floor.load_type_element(
                        year=bldg.year_of_construction,
                        construction=group["FloorConstruction"].iloc[0])
                else:
                    warnings.warn("Values for IsGroundFloor have to be either 0 or 1, for no or yes respectively")
            else:
                warnings.warn("Zone \"%s\" with IsGroundFloor \"%s\" and construction type \"%s\" "
                              "has no floor nor groundfloor, since the area equals 0."
                              %(group["Zone"].iloc[0], group["IsGroundFloor"].iloc[0],
                                group["FloorConstruction"].iloc[0]))

        Grouped = Zone.groupby(["IsRooftop", "CeilingConstruction"])
        for name, group in Grouped:
            if group["NetArea[m²]"].sum() != 0: # to avoid devision by 0
                if group["IsRooftop"].iloc[0] == 1:
                    rooftop = Rooftop(parent=tz)
                    rooftop.name = "rooftop" + \
                                    str(group["CeilingConstruction"].iloc[0])
                    rooftop.area = group["NetArea[m²]"].sum()   #sum up area of respective rooftop parts
                    rooftop.tilt = RooftopTilt
                    rooftop.orientation = RooftopOrientation
                    # load wall properties from "TypeBuildingElements.xml"
                    rooftop.load_type_element(
                        year=bldg.year_of_construction,
                        construction=group["CeilingConstruction"].iloc[0])
                elif group["IsRooftop"].iloc[0] == 0:
                    ceiling = Ceiling(parent=tz)
                    ceiling.name = "ceiling" + \
                                    str(group["CeilingConstruction"].iloc[0])
                    ceiling.area = group["NetArea[m²]"].sum() / 2 # only half of the ceiling belongs to a story, the other half to the above
                    ceiling.tilt = CeilingTilt
                    ceiling.orientation = CeilingOrientation
                    # load wall properties from "TypeBuildingElements.xml"
                    ceiling.load_type_element(
                        year=bldg.year_of_construction,
                        construction=group["CeilingConstruction"].iloc[0])
                else:
                    warnings.warn("Values for IsRooftop have to be either 0 or 1, for no or yes respectively")
            else:
                warnings.warn("Zone \"%s\" with IsRooftop \"%s\" and construction type \"%s\" "
                              "has no ceiling nor rooftop, since the area equals 0."
                              %(group["Zone"].iloc[0], group["IsRooftop"].iloc[0],
                                group["CeilingConstruction"].iloc[0]))

        Grouped = Zone.groupby(["InnerWallConstruction"])
        for name, group in Grouped:
            if group["InnerWallArea[m²]"].sum() != 0: # to avoid devision by 0
                in_wall = InnerWall(parent=tz)
                in_wall.name = "inner_wall" + \
                                str(group["InnerWallConstruction"].iloc[0])
                in_wall.area = group["InnerWallArea[m²]"].sum() / 2 # only half of the wall belongs to each room, the other half to the adjacent
                # load wall properties from "TypeBuildingElements.xml"
                in_wall.load_type_element(
                    year=bldg.year_of_construction,
                    construction=group["InnerWallConstruction"].iloc[0])
            else:
                warnings.warn("Zone \"%s\" with inner wall construction \"%s\" has no inner walls, since area = 0."
                              %(group["Zone"].iloc[0], group["InnerWallConstruction"].iloc[0]))

        # Block: AHU and infiltration #Todo: Attention hard coding
        tz.use_conditions.with_ahu = True
        #tz.use_conditions.min_ahu = 13.5
        #tz.use_conditions.max_ahu = 13.5
        tz.use_conditions.use_constant_ach_rate = True
        tz.infiltration_rate = 0.2

        #Todo: maybe better to read from excel!?
        ahu_dict_TK03 = {"Bedroom" : [15.778, 15.778],
                         "Corridorsinthegeneralcarearea" : [5.2941, 5.2941],
                         "Examinationortreatmentroom" : [15.743, 15.743],
                         "MeetingConferenceseminar" : [16.036, 16.036],
                         "Stocktechnicalequipmentarchives" : [20.484, 20.484],
                         "WCandsanitaryroomsinnonresidentialbuildings" : [27.692, 27.692],}
        _i = 0
        for key in ahu_dict_TK03:
            if tz.name == key:
                tz.use_conditions.min_ahu = ahu_dict_TK03[key][0]
                tz.use_conditions.max_ahu = ahu_dict_TK03[key][1]
                _i = 1
        if _i == 0:
            warnings.warn("The zone %s could not be found in your ahu_dict. Hence, no AHU flow is defined. The default value is 0 (min_ahu = 0; max_ahu=0" %tz.name)

    return prj, Data

if __name__ == '__main__':
    result_path = "D:\Teaser_exports\HUS"

    prj = Project(load_data=True)
    prj.name = "TK03_statisticElements_AHU_infiniteHeat_realACH_TEEEEEEEEEETTTTTTTTTTTTTTTTT"
    prj.data.load_uc_binding()
    prj.weather_file_path = os.path.join(
        os.path.dirname(os.path.dirname(__file__)),
        'data',
        'input',
        'inputdata',
        'weatherdata',
        'FIN_Helsinki.029740_IWEC.mos')
    prj.modelica_info.weekday = 0  # 0-Monday, 6-Sunday
    prj.modelica_info.simulation_start = 0  # start time for simulation
    prj, Data = import_building_from_excel(prj, "HUS", 2000, "HUS_Etage_7_Aktuell.xlsx", SheetNames = ["Area1"])
    prj.modelica_info.current_solver = "dassl"
    prj.calc_all_buildings(raise_errors=True)
    prj.export_aixlib(
        internal_id=None,
        path=result_path)

    # if wished, export the zoned dataframe which is finally used to parameterize the teaser classes
    Data.to_excel(os.path.join(result_path, prj.name, "ZonedInput.xlsx"))

    print("%s: That's it :)" %prj.name)
    
    

