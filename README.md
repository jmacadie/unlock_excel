# Unlock Excel

A small command-line utility to read or remove the VBA protection on Excel files.

It works on all of:
- xls:  Excel 97-2003 Workbooks
- xlsm: Excel Macro-Enabled Workbooks
- xlsb: Excel Binary Workbooks

It will not work with xlsx files since, by definition, they do not include any VBA.

This is pretty rough and ready, so feel free to report any issues.

## Usage

To read the protection on a file:

`$ ./unlock_excel read FILENAME`

Usually the password, if there is one, will be an SHA1 hash of the password plus a
random salt. Both the hash result and the salt will be printed out. These can be
input into password recovery tools such as [hashcat](https://hashcat.net/hashcat/)
or [John the Ripper](https://www.openwall.com/john/) to decrypt the password

Optionally you can pass the `-d` flag to get the application to try to decrypt
against a list of 1.7 million common passwords:
`$ ./unlock_excel read -d FILENAME`

To remove protection on a file:

`$ ./unlock_excel remove FILENAME`

By default, this will be saved to a copy of the original file with '_unlocked'
appended to the name. If you wish you update the file in place, pass the `-i`
flag:
`$ ./unlock_excel remove -i FILENAME`

## WARNING

This utility is designed only to give the user access to files that they already
have the rights to read and edit. For example, gaining access to an old file at
work for which the password has been lost.

USE OF THIS UTILITY TO BREAK ANY LAWS IS NOT CONDONED
