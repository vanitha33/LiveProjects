'Description: 			It reads valid data from External Excel Sheet and  enter data in Login Application'
' Inputs: 					 Data in External Excel Sheet must be valid Login data
'Outputs:				 It Logs into the Login Application.

'Load Login and Flight Reservation Objects
RepositoriesCollection.Add "../../Repository/LoginRepository.tsr"
RepositoriesCollection.Add "../../Repository/FlightReservation.tsr"

'Load Environment Details
Environment.LoadFromFile "../../../Environment.xml"

'Define Sheet Name Imported from External  Excel Sheet  into QTP
Const QtpSheetName= "Qtp_Login_Success"

'Imports Data From External Excel Sheet into QTP
DataTable.AddSheet QtpSheetName
DataTable.ImportSheet  "../../TestData/Login.xls","Success",QtpSheetName

'Get Row Count from  Imported  Excel Sheet
rowcount=DataTable.GetSheet(QtpSheetName).GetRowCount

'Read one Row at a time from Excel Sheet and Execute Login and close Flight Reservation at end
For i= 1 to rowcount

SystemUtil.Run Environment("FlightExecutable")
Dialog("Login").Activate
Dialog("Login").WinEdit("Agent Name:").Set DataTable(1,QtpSheetName)
Dialog("Login").WinEdit("Password:").Set DataTable(2,QtpSheetName)
Dialog("Login").WinButton("OK").Click

DataTable.GetSheet(QtpSheetName).SetNextRow

'If Flight Reservation Exist it closes the window
If  Window("Flight Reservation").Exist(10) Then
	Window("Flight Reservation").WinMenu("Menu").Select "File;Exit"
End If
next












