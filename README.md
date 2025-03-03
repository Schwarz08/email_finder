# email_finder
## Input:
### All input variables can be found under main.
* email_list: List of email addresses to search emails
* folder_list: List of folders to search emails. The tool will search these folders per email address.
* kw_list.csv: csv file containing information on the emails to be searched
### kw_list
|Parameter|Description|
|---|---|
|MSG Identifier 1|Not used when searching emails, only for sorting them into folders on export together with MSG Identifier 2.|
|MSG Identifier 2|Not used when searching emails, only for sorting them into folders on export together with MSG Identifier 1.|
|Subject Keyword|Each email search query should have a Initial Subject Keyword to lessen the emails to loop through.|
|Subject Keyword (Add)|Additional optional keywords to search for in the email subject apart from the Preliminary Subject Keyword. Additional keywords are separated by ","|
|Body Keyword|Additional optional keywords to search for in the email body apart from the Preliminary Subject Keyword, separated by ","|
## Output:
* kw_list_updated.csv: Updated kw_list. Summarizes which emails were found.
* filtered_emails: Folder where the emails found will be exported to
