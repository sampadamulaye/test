var xfo = WScript.CreateObject( "olexlsf.XLSFileObject" );
var XFile = xfo.XLSFile;

XFile.Workbook.Sheets(1).Cells.Cell(1,1).Value = "test autofit";
XFile.Workbook.Sheets(1).Cells.Cell(1,1).FontBold = 1;
XFile.Workbook.Sheets(1).Cells.Cell(1,1).FontHeight = 12;
XFile.Workbook.Sheets(1).Cells.Cell(1,1).Rotation = 45;

XFile.Workbook.Sheets(1).Cells.Cell(1,2).Value = "test autofit";
XFile.Workbook.Sheets(1).Cells.Cell(1,2).FontBold = 1;
XFile.Workbook.Sheets(1).Cells.Cell(1,2).FontHeight = 24;


XFile.Workbook.Sheets(1).Columns.AutoFit();

XFile.SaveAs ("out.xls");
WScript.Quit(1);
