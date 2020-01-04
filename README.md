# questrade-interface
Interface to Questrade using their API in order to extract account data

This code can be run in Excel using a VBA macro.

In order to connect, you must log into Questrade and copy a token that is generated in the API center. See Questrade for details, the API center for account holders can be found here: https://login.questrade.com/APIAccess/UserApps.aspx

The token should be pasted into cell "B2" of an Excel spreadsheet. Once run, the code will connect to the account, extract details including the account number and type, then pull down the account balance and security positions.
