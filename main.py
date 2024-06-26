import time

from aardvark_py import *
from array import array
import pandas as pd


def bar(progress, total):
    proportion = progress / total
    filled_length = int(proportion * 100)
    filled = '#' * filled_length

    blank_length = 100 - filled_length
    blank = ' ' * blank_length

    return f'[{filled}{blank}] {filled_length}%'


print("Starting\n")
print("Reading EMC2104_eeprom_prog_table.xlsx", end='')
dataframe = pd.read_excel('EMC2104_eeprom_prog_table.xlsx', sheet_name='DCB_Modified')
addresses = dataframe['Address Offset (hex)'].values
hex_data = dataframe['Hex'].values
data = {address: data for address, data in zip(addresses, hex_data)}
print(" - DONE")

print("Configuring Aardvark", end='')
aardvark = aa_open(0)
if aardvark <= 0:
    print("\nUnable to open Aardvark device on port %d" % 0)
    print("Error code = %d" % aardvark)
    exit()
else:
    aa_configure(aardvark, AA_CONFIG_SPI_I2C)
    bitrate = aa_i2c_bitrate(aardvark, 100)  # Set the bitrate to 100 kHz
    bus_timeout = aa_i2c_bus_timeout(aardvark, 150)  # Set the timeout to 150 ms
    print(" - DONE")

# Set the slave device address

serial = ''
while True:
    serial = input("\nEnter Serial No: N34083-101-").strip()
    if len(serial) == 3:
        break
    print("Must be 3 characters.")

data['8B'] = format(ord(serial[0]), '02X')
data['8C'] = format(ord(serial[1]), '02X')
data['8D'] = format(ord(serial[2]), '02X')

errors = []

if aa_i2c_write(aardvark, 0x71, AA_I2C_NO_FLAGS, array('B', [0x08])) < 0:
    errors.append('71')
time.sleep(.02)

print()
for i, (addr, val) in enumerate(data.items()):
    if pd.isnull(addr) or pd.isnull(val):
        continue

    if aa_i2c_write(aardvark, 0x50, AA_I2C_NO_FLAGS, array('B', [int(addr, 16), int(val, 16)])) < 1:
        errors.append(addr)
        print(f'\rError writing 0x{val} to 0x{addr}     {bar(i + 1, len(data))}', end='')
    else:
        print(f'\rWrote 0x{val} to 0x{addr}     {bar(i + 1, len(data))}', end='')

    time.sleep(.01)

print()

if aa_i2c_write(aardvark, 0x55, AA_I2C_NO_FLAGS, array('B', [int('15', 16), int('01', 16)])) < 0:
    errors.append('0x55-0x15')
time.sleep(.5)

if len(errors) > 0:
    print(f"Error writing at 0x{errors}")

aa_i2c_write(aardvark, 0x50, AA_I2C_NO_STOP, array('B', [0x80]))
(result, data_in) = aa_i2c_read(aardvark, 0x50, AA_I2C_NO_FLAGS, 14)

print("\nRead back: ", end='')
if result < 0:
    print("\nFailed to read data from I2C device.")
elif result == 0:
    print("\nNo bytes read. Possible issue with the I2C bus.")
else:
    print("".join(f"{chr(data)}" for data in data_in))

# Close the device
aa_close(aardvark)
