var xfo = WScript.CreateObject( "olexlsf.XLSFileObject" );
var XFile = xfo.XLSFile;

XFile.Workbook.Sheets(1).Cells.Cell(1,1).Value = 1;
XFile.Workbook.Sheets(1).Cells.Cell(2,1).Value = 2;
XFile.Workbook.Sheets(1).Cells.Cell(3,1).Value = 3;
XFile.Workbook.Sheets(1).Cells.Cell(4,4).Formula = "A1-A2-A3";

XFile.SaveAs ("out.xls");
WScript.Quit(1);
