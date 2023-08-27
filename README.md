# VBA-challenge
# Note to Graders:
- 2 scripts (btInit_Click() and GetTopTickers()) were created to loop through all the worksheets without to create individual script per worksheet or having to run on each active sheet.
- A Controller sheet was added for sake of fun for code practicing and monitoring so I don't have to look at the spinning wheel only while waiting for 5 minutes or more to the script to complete [see snippet "Controller.png"]). The main script btInit_Click() contained lines related to the Controller and have been noted to indicate which section is related to this worksheet so they can easily be commented out when testing.
- The second script, GetTopTickers() loops through all worksheets and checks for the top tickers based on the output from the first script.
- Both scripts use a button located on the Controller sheet to execute.