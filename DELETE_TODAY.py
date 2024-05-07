from pyModbusTCP.client import ModbusClient
import time
import math
import os
import pandas as pd
import numpy as np
import sys
import win32com.client
import week8finalcode
import matplotlib.pyplot as plt

import psutil
import collections

pts_n = 100
x = []
y = []
num=0
# Create the figure and axis outside the loop
fig, ax = plt.subplots()
line, = ax.plot(x, y, linestyle="--")

#my_process = psutil.Process(os.getpid())
#t_start = time.time()


dss_ema = week8finalcode.DSS_EMA()
dss_ema.check()
dss_ema.set_excel_filename("C:\\Users\\DELL\\Desktop\\FYP-main\\IEEE-13 excel sheet.xlsx")
dss_ema.load_data_from_excel()


def fpfrom754(first, second):
    """ Convert two 16-bit integers (IEEE 754 format) back to a floating-point number. """
    full = (first << 16) | second
    if full == 0:
        return 0.0
    sign = -1 if (full >> 31) else 1
    exponent = ((full >> 23) & 0xFF) - 127
    mantissa = (full & 0x7FFFFF) | 0x800000  # The implicit leading bit
    return sign * mantissa / (1 << 23) * (2 ** exponent)



# Setup the client
client = ModbusClient(host='10.20.2.243', port=1201, auto_open=True)

try:
    index = 0  # Start index of registers to read from
    while True:
        # Read two registers at a time for each floating-point value
        regs = client.read_holding_registers(index, 2)
        if regs and len(regs) == 2:
            value = fpfrom754(regs[0], regs[1])
            print(f"Value at indices {index}, {index+1}: {value}")
            a = dss_ema.solve_snapLV_real(load_mult=value)
            total_power = a["TotalPower"][0]
            print("this is total power",total_power)
            num+=1
            x.append(num)
            y.append(total_power*-1)

            line.set_xdata(x)
            line.set_ydata(y)
            ax.relim()
            ax.autoscale_view()
            plt.pause(0.1)

            index += 2
        else:
            print(f"Read error at registers: {index}, {index+1}")
            index += 2  # Increment index even in case of error to try the next value

        time.sleep(2)  # Sleep for 3 seconds before reading the next value

        # Reset index to 0 if it reaches the end of the expected range of registers
        # assuming you are cycling through a fixed set of floating-point values
        if index >= 2 * 24:  # Adjust 24 based on the actual number of floats you expect
            index = 0

except KeyboardInterrupt:
    print("Client stopped by user.")
finally:
    client.close()
    plt.close()