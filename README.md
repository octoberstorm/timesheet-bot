## Timesheet Bot
Log check-in/checkout data from Google Hangouts via Chatbot and sync daily data to remote/official Spreadsheet

### Versions
#### V1 
- Uses Appscript only
- Hard-coded official monthly sheet name and timesheet columns

#### V2 
- Uses CloudFunction for adding check-in/checkout records
- Automatically lookup for official monthly sheet name and timesheet columns

### Sheet formats

#### Local sheet: sheet names
- MemberMapping: storing member mapping gsuitename-officialname
- ProjectList: storing list of projects to manage
- Project timesheet sheet: store check-in/out data for each project

#### Local sheet: project-list sheet
```
ProjectCool |
_____________
ProjectHot  |
```

#### Local sheet: each-project sheeet
```
DATE        | NAME      | TIME IN   |	TIME OUT    | PROJECT
_________________________________________________________________
20-08-2021  | Doe John  | 7:33      | 17:20         | ProjectCool
```

#### Local sheet: MemberMapping
```
GsuiteName      |	OfficialName
________________________________
Doe John        |   John Doe
```

#### Remote/official sheet
```
NAME        |              20-08-2021       |           21-08-2021
_____________________________________________________________________________
-           | Check in | Check out | Trans  | Check in | Check out | Trans  |
_____________________________________________________________________________
John Doe    | 7:33     | 17:20     |        |          |           |        |
```

### Deployment (v2)
- Update related spreadsheet URLs (local, remote sheets)
- Create local sheet with sheets: MemberMapping, project-name sheets
- Create Hangouts chat box using project name, the project name is used to match local sheet name later for storing check-in/out data
- Deploy the Google Cloud Function
- Add permission for the function to edit local spreadsheet
- Deploy Apps Script, get deployment ID and put to Hangouts bot API
- Add Chatbot to each Hangouts project box
- Create Triggers for syncing data (which call sync function)

### References
<https://cloud.google.com/blog/products/g-suite/building-a-google-hangouts-chatbot-using-apps-script>