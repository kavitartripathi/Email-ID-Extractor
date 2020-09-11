# Email-ID-Extractor
This tool scans the Inbox of the provided email account and extracts all the email addresses into a CSV file.
Note:
1) Copy the connection.config.sample file to connection.config and enter your credentials in this new file. This file is ignored by source control to protect leakage of credentials.
2) The output file is located in the Outputs folder next to the compiled exe.
3) The program picks up from where it left off (even in case of an error or crash) by storing a count of emails processed as part of the output file name and updating the output file periodically. By keeping this file around, you can run the tool again at any time and it will only process the new emails since the last time it was run.
4) When the program ends successfully, it removes duplicates from the output file.
