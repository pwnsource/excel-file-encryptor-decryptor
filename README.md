**Excel File Encryptor/Decryptor**
This Python script allows you to encrypt and decrypt Excel files using the Fernet symmetric encryption algorithm provided by the cryptography library. You can use it to protect sensitive data in your Excel files by encrypting them before sharing or storing them securely.

**Features**
Encrypt Excel files to protect sensitive data.
Decrypt encrypted Excel files to access the original data.
Simple command-line interface for ease of use.
Prerequisites
Before running the script, make sure you have the following installed:

Python 3.x
Required Python libraries:
cryptography
openpyxl


You can install the required libraries using pip:
pip install cryptography openpyxl


**Usage**
Encrypting Excel Files
Run the **encrypt.py** script.
Enter the path of the Excel file you want to encrypt when prompted.
Enter the path where you want to save the encrypted file.
Decrypting Excel Files
Run the **decrypt.py** script.
Enter the path of the encrypted Excel file you want to decrypt when prompted.
Enter the path where you want to save the decrypted file.
Note
Remember to securely store the encryption key (encryption_key.txt) generated during the encryption process. Losing the encryption key will result in permanent data loss.
