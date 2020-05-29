# jira-to-goole-spreadsheet
Export Jira task to Google Spreadsheet script

# Install "Jira Cloud for Sheets" plugin
Install "Jira Cloud for Sheets" plugin from https://gsuite.google.com/marketplace .It will add `=Jira()` formula

# Go to Jira search and make this set of fields as default:
There are some column names hardcoded in this script it is why columns order is important.
```
Issue Type	Key	Summary	Assignee	Reporter	Priority	Status	Resolution	Created	Updated	Due date	Story Points	Epic Name	Epic Link	Sprint	Development	Resolved
```

# Crete sheets in you Google Spreadsheet document
## Crete sheet wuth the name "Epics"
There is bug in Jira for about 3-4 years. It can't export Epics Names. This sheet is needed to replace Epic ID to epic Name.
Add formnula to A1 cell (replace XXXXX):

```
=JIRA("issuetype = Epic AND project = XXXXX order by lastViewed DESC","issuekey,summary")
```

## Crete sheet wuth the name "Sprints"
Manually or with formula add duration of your sprints in this sheet:

| Start	| Finish |
| --- | --- |
| 12 May 2020	| 19 May 2020 |
| 19 May 2020	| 26 May 2020 |
| 26 May 2020	| 2 Jun 2020 |

Tcript will load tasks sprint by sprint and add extra fields to be able to group by sprint in pivot table.

# Create Script
 - Go to "Tools" > "Script" Editor and copy there `Code.gs` file from this repository
 - Try to run importJira `importJira` function and add needed permissions to the script

# Import Data
There will be new menu "Jira" > "Import Data"
