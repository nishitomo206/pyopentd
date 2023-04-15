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

from System.Collections.Generic import List
import OpenTDv62 as otd


def get_properties(obj):
    for prop in obj.GetType().GetProperties():
        val = getattr(obj, prop.Name)
        print("----------")
        print("Name: ", prop.Name)
        print("val : ", val)
        if str(type(val)) == "<class 'OpenTDv62.ExpressionArrayClassData'>":
            print(val.expression)
    return
