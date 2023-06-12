import openpyxl

# Function to extract data from a frame
def extract_frame_data(frame):
    # Extract header data
    header_data = frame[:22].decode('utf-8')

    # Extract packet 2 validity byte
    packet_2_validity_byte = frame[47]

    # Extract parameters from packet 2
    parameters = []
    byte_index = 52

    # Parameters with 8 bytes each
    for _ in range(16):
        parameter = int.from_bytes(frame[byte_index:byte_index + 8], 'big')
        parameters.append(parameter)
        byte_index += 8

    # Skip 3 bytes
    byte_index += 3

    # Parameters with 8 bytes each
    for _ in range(16):
        parameter = int.from_bytes(frame[byte_index:byte_index + 8], 'big')
        parameters.append(parameter)
        byte_index += 8

    # Skip 3 bytes
    byte_index += 3

    # Parameters with 8 bytes each
    for _ in range(9):
        parameter = int.from_bytes(frame[byte_index:byte_index + 8], 'big')
        parameters.append(parameter)
        byte_index += 8

    # Skip 9 bytes
    byte_index += 9

    # Parameters with 8 bytes each
    for _ in range(31):
        parameter = int.from_bytes(frame[byte_index:byte_index + 8], 'big')
        parameters.append(parameter)
        byte_index += 8

    # Skip 11 bytes
    byte_index += 11

    # Parameters with 8 bytes each
    for _ in range(23):
        parameter = int.from_bytes(frame[byte_index:byte_index + 8], 'big')
        parameters.append(parameter)
        byte_index += 8

    return header_data, packet_2_validity_byte, parameters

# Main program
def main():
    # Open the .dat file as a binary file
    with open('input.dat', 'rb') as file:
        # Read the entire content of the file
        data = file.read()

    # Create a new Excel workbook
    workbook = openpyxl.Workbook()

    # Select the active sheet
    sheet = workbook.active

    # Write the heading for the PACKET 2 DATA row
    sheet.append(["PACKET 2 DATA"])

    # Write the headings for the parameters
    headings = ["SI.NO", "Frame ID", "Packet validity byte"]
    headings.extend([f"Parameter {i}" for i in range(1, 102)])
    sheet.append(headings)

    # Iterate over the frames
    for frame_index in range(0, len(data), 1998):
        frame = data[frame_index:frame_index + 1998]

        # Extract data from the frame
        header, validity_byte, parameters = extract_frame_data(frame)

        # Write the frame data to the Excel file
        row_data = [frame_index // 1998 + 1, frame_index // 1998, validity_byte]
        row_data.extend(parameters)
        sheet.append(row_data)

    # Save the workbook to a file
    workbook.save("RESULTS.xlsx")

if __name__ == "__main__":
    main()
