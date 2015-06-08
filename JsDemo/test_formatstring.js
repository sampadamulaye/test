var xfo = WScript.CreateObject( "olexlsf.XLSFileObject" );
var XFile = xfo.XLSFile;

XFile.Workbook.Sheets(1).Cells.Cell(1,1).Value = 1;
XFile.Workbook.Sheets(1).Cells.Cell(1,1).FormatStringIndex = 3;
XFile.Workbook.Sheets(1).Columns.AutoFit_Columns(1,1,100);

XFile.SaveAs ("out.xls");
WScript.Quit(1);
