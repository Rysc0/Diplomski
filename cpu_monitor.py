import csv
import sys
from datetime import datetime
import time

CPU_TEMP_PATH = '/sys/class/hwmon/hwmon2/temp1_input'
CPU_FREQUENCY_PATH = '/sys/devices/system/cpu/cpu0/cpufreq/scaling_cur_freq'

# time in seconds
TIME_BETWEEN_MEASUREMENTS = int(sys.argv[1])
# time in minutes
NUM_OF_MEASUREMENTS = int(sys.argv[2])
# NUM_OF_MEASUREMENTS = 10
# TIME_BETWEEN_MEASUREMENTS = 2
READINGS = []

while NUM_OF_MEASUREMENTS > -1:
    with open(CPU_TEMP_PATH, 'r') as temperature_reading, open(CPU_FREQUENCY_PATH, 'r') as frequency_reading:
        # print(temperature_reading.read())
        current_time = datetime.now().strftime("%d-%m-%Y %H:%M:%S")
        current_temp = temperature_reading.read()[:-1]  # Read the temperature from a file and cut off new line from the end ('\n')
        current_freq = frequency_reading.read()[:-1]
        READINGS.append((current_time, current_temp, current_freq))
        print(f"TIME: {current_time} -> TEMP: {current_temp} -> FREQ: {current_freq}" )
        NUM_OF_MEASUREMENTS -= 1
        if NUM_OF_MEASUREMENTS == -1:
            break
        time.sleep(TIME_BETWEEN_MEASUREMENTS)
    # print(TIME_BETWEEN_MEASUREMENTS, ' -> ', DURATION)

print("READINGS: ", READINGS)

with open("temp_freq_raw.csv", 'w') as output:
    writer = csv.writer(output)
    writer.writerow(["Time", "Temperature", "Frequency"]) # HEADER
    writer.writerows(READINGS)


