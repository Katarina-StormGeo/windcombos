import json
import pandas as pd
import time
from itertools import dropwhile
import sys 
import os
import subprocess
import requests
import argparse
from io import StringIO
import logging
from logging.handlers import TimedRotatingFileHandler
import numpy as np
import traceback
import operator
import re
from dateutil.relativedelta import relativedelta
import glob
from datetime import datetime, timedelta, date
from calendar import monthrange, month_name
import xlwings as xw
import locale
from dateutil.parser import parse



def get_ec_run(run):
    if run == 0:
        wind_ops = "M:/Database/PowerWeather/WindPowerForecastEC00.xls"
        sol_ops = "M:/Database/PowerWeather/SolarPowerForecastEC00.xls"

        wind_ens = "M:/Database/PowerWeather/WindPowerForecastECEPS00.xls"
        sol_ens = "M:/Database/PowerWeather/SolarPowerForecastEM00.xls"
    if run == 12:
        wind_ops = "M:/Database/PowerWeather/WindPowerForecastEC12.xls"
        sol_ops = "M:/Database/PowerWeather/SolarPowerForecastEC12.xls"

        wind_ens = "M:/Database/PowerWeather/WindPowerForecastECEPS12.xls"
        sol_ens = "M:/Database/PowerWeather/SolarPowerForecastEM12.xls"
    
    return wind_ops, wind_ens, sol_ops, sol_ens

def get_ec_wind(ops, ens, combo):


    wind_op = ops

    operational = pd.read_excel(wind_op, header=1)
    operational.columns.values[0] = 'time'

    operational = operational.iloc[operational.iloc[:,1].first_valid_index():]  #finding the first non-nan value, choosing this as the first val
    operational = operational.set_index('time')

    op_start = datetime.now().strftime("%Y-%m-%d")
    op_end = (datetime.now() + timedelta(days=combo)).strftime("%Y-%m-%d")

    op = operational[op_start:op_end]

    wind_ens = ens

    ensamble = pd.read_excel(wind_ens, header=1)
    ensamble.columns.values[0] = 'time'
    ensamble = ensamble.set_index('time')

    ens_start = (datetime.now() + timedelta(days=combo)).strftime("%Y-%m-%d")
    ens_end = ensamble.iloc[:,1].last_valid_index()

    ens = ensamble[ens_start:ens_end]



    wind = pd.concat([op, ens])

    if combo == 0:
        wind = ens


    wind = wind.drop(wind.columns[26:],axis = 1)

    wind.index.name = None
    return wind


def get_ec_solar(ops, ens, combo):

    solar_op = ops

    operational = pd.read_excel(solar_op, header=1)
    operational.columns.values[0] = 'time'

    operational = operational.iloc[operational.iloc[:,1].first_valid_index():]  #finding the first non-nan value, choosing this as the first val
    operational = operational.set_index('time')

    op_start = datetime.now().strftime("%Y-%m-%d")
    op_end = (datetime.now() + timedelta(days=combo)).strftime("%Y-%m-%d")

    op = operational[op_start:op_end]


    solar_ens = ens

    ensamble = pd.read_excel(solar_ens, header=1)
    ensamble.columns.values[0] = 'time'
    ensamble = ensamble.set_index('time')

    ens_start = (datetime.now() + timedelta(days=combo)).strftime("%Y-%m-%d")
    ens_end = ensamble.iloc[:,1].last_valid_index()

    ens = ensamble[ens_start:ens_end]


    solar = pd.concat([op, ens])

    if combo == 0:
        solar = ens

    solar.index.name = None
    

    return solar

def write_to_excel(run):
    wind_ops, wind_ens, sol_ops, sol_ens = get_ec_run(run)

    it = [0,1,2,3,4,5,6,7,8,9,10]

    wind_combos = []
    solar_combos = []
    
    for i in it:
        wind_combo = get_ec_wind(wind_ops, wind_ens, i)
        wind_combos.append(wind_combo)

        solar_combo = get_ec_solar(sol_ops, sol_ens, i)
        solar_combos.append(solar_combo)

    if run == 0:
        runname = '00'
    if run == 12:
        runname = '12'
    
    with pd.ExcelWriter(f'{runname}_wind.xlsx') as writer:
        for i, combo in enumerate(wind_combos):
            combo.to_excel(writer, sheet_name=f'split {i} days')

    with pd.ExcelWriter(f'{runname}_solar.xlsx') as writer:
        for i, combo in enumerate(solar_combos):
            combo.to_excel(writer, sheet_name=f'split {i} days')
    

write_to_excel(0)

