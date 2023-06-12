import struct
from openpyxl import Workbook

# Define the file paths
dat_file_path = 'input.dat'
excel_file_path = 'RESULTS.xlsx'

# Define constants for frame size and parameter sizes
FRAME_SIZE = 1998
HEADER_SIZE = 22
PACKET1_SIZE = 24
PACKET2_SIZE = 842
PARAMETER_SIZES = [8, 8, 8, ...]  # Replace ... with the sizes for the remaining 98 parameters

# Create a workbook and get the active sheet
wb = Workbook()
sheet = wb.active

# Write the heading row
heading_row = ['SI.NO', 'Frame ID', 'Packet validity byte', 'Parameter 1', 'Parameter 2', ...]  # Replace ... with the headings for the remaining 98 parameters
sheet.append(heading_row)

# Read the .dat file as binary
with open(dat_file_path, 'rb') as dat_file:
    frame_id = 1
    byte_index = 0

    while True:
        # Read the frame data
        frame_data = dat_file.read(FRAME_SIZE)
        if not frame_data:
            break

        # Extract the header data
        header_data = frame_data[:HEADER_SIZE]
        header = struct.unpack('!' + 'c' * HEADER_SIZE, header_data)
        header = [h.decode() for h in header]

        # Extract the packet 2 data
        packet2_data = frame_data[HEADER_SIZE + PACKET1_SIZE:HEADER_SIZE + PACKET1_SIZE + PACKET2_SIZE]

        # Extract the packet validity byte
        packet_validity_byte = struct.unpack('!b', packet2_data[0:1])[0]

        # Extract the parameters from packet 2
        parameters = []
        param_start_index = 1  # Starting index of the first parameter

        for param_size in PARAMETER_SIZES:
            parameter_data = packet2_data[param_start_index:param_start_index + param_size]
            if param_size == 8:
                parameter = struct.unpack('!d', parameter_data)[0]  # Convert to double precision floating point
            elif param_size == 1:
                parameter = struct.unpack('!?', parameter_data)[0]  # Convert to binary
            elif param_size == 2:
                parameter = struct.unpack('!H', parameter_data)[0]  # Convert to unsigned integer
            else:
                # Handle other parameter sizes if needed
                parameter = None

            parameters.append(parameter)
            param_start_index += param_size

        # Write the data to the Excel sheet
        row_data = [frame_id, frame_id, packet_validity_byte] + parameters
        sheet.append(row_data)

        frame_id += 1
        byte_index += FRAME_SIZE

# Save the Excel file
wb.save(excel_file_path)
