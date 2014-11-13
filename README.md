DynForms
========

Custom and dynamicly generated Google Spreadsheets linked Forms for all your sheets in the spreadsheet


Setup
=====

- In the desired spreadsheet, open the script editor
- Make a new webapp script
- Go copying every file in the github project to new files in the script editor
- Change your spreadsheetid
- Publish as webapp and save your link


Usage
=====

- The sheets of the spreadsheet must have a first row with headers, and every field you want to include in the form should have an attached note with this json configuration contents
```
        {"Type":"", 
         "Title":"",
         "Description":"",
         "DefValue":"",
         "Mandatory":"",
         "Items":[ "", ""]
        }
```

  Type: One of [autoincrement|autodate|autotime|autodatetime|number|text|textarea|checkbox|date|time|select]
  Title: The title you want to show
  Description currently not used
  DefValue: The default value for the field
  Mandatory: If this field have to be provided
  Items: Only if the type is select, the selectable items

- Send the link to your friends!


TODO
====

- Create more user-friendly CSS
- Create mobile friendly version (jquery?)
- Fix mandatory in checkboxes
- Fix default in checkboxes
- Format the cell on the sheet, depending on the given type
- Show only sheets with at least one field settedup
- Add google auth to control one answer per user
- Add new type as autoformula to autogenerate results based on user inputs 

Extra
=====

If you like my work, buy me a drink https://www.paypal.com/cgi-bin/webscr?cmd=_s-xclick&hosted_button_id=ZWE8DKMCP6ENJ


License
=======
```
                    GNU GENERAL PUBLIC LICENSE
                       Version 3, 29 June 2007
```
