from cryptography.fernet import Fernet
from openpyxl import load_workbook

# Generate a key for encryption
key = Fernet.generate_key()
cipher_suite = Fernet(key)

# Input filename
file_path = input("Enter the name of original Excel file: ")

try:
    # Load the Excel file
    workbook = load_workbook(filename=file_path)

    # Encrypt each cell in all sheets
    for sheet in workbook.sheetnames:
        ws = workbook[sheet]
        for row in ws.iter_rows():
            for cell in row:
                cell.value = cipher_suite.encrypt(str(cell.value).encode())

    # Save the encrypted Excel file
    encrypted_file_path = input("Enter the name of the encrypted file (needs to end with .xlsx): ")
    workbook.save(encrypted_file_path)

    # Store the encryption key securely
    with open('encryption_key.txt', 'wb') as key_file:
        key_file.write(key)

except FileNotFoundError:
    print("File not found.")
