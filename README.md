# Email-ID-Extractor
This tool scans the Inbox of the provided email account and extracts all the email addresses into a CSV file.
Note:
1) Email credentials need to be provided in the connection.config file.
2) The temporary output will be provided in the output\temp folder.
3) If the program crashes in between execution, the program picks up from where it left off and at the end merges the email_ids from the temp file.
4) The final output file with all unique email ids will be generated at the end of the program execution.
