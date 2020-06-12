"# Email-Attachment-Dw-and-Analysis---Python" 

This is a simple python script which uses email and imaplib libraries to connect to mail servers.
The Problem Statement is quite simple. A user wants to download data of shipments that come daily
with the same subject name. First task is to download the attachment and extract the zipped file.
After extracting the file, the data is a txt file and in the following format.

SPNA	PRIME	ADMIN DESK	B01HBFQW6M	1165 
SPNA	CUSTOMER_RETURN	CRET1	B00TS8ABHC	1	
SPNA	CUSTOMER_RETURN	CRET1	B01CHUPK8W	1
SPNA	CUSTOMER_RETURN	CRET1	B01CHUYFTW	1
SPNA	CUSTOMER_RETURN	CRET1	B01CHV6F1W	1
SPNA	CUSTOMER_RETURN	CRET1	B01CHVA9KK	1
SPNA	CUSTOMER_RETURN	CRET1	B01CJVVGH8	1

Using pandas, we analyze the data and store it in normal excel spreadsheets.