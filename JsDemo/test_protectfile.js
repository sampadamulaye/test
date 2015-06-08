var xfo = WScript.CreateObject( "olexlsf.XLSFileObject" );
var XFile = xfo.XLSFile;

XFile.Workbook.Sheets(1).Cells.Cell(1,1).Value = "test";
XFile.Workbook.Sheets(1).Cells.Cell(2,1).Value = "test2";
XFile.Workbook.Sheets(1).Cells.Cell(3,1).Value = "test3";

// save to password-protected file
// password is 123
XFile.SaveAsProtected ("out.xls", "123");
WScript.Quit(1);

