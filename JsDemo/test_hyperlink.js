var xfo = WScript.CreateObject( "olexlsf.XLSFileObject" );

var XFile = xfo.XLSFile;

// Add new sheet
XFile.Workbook.Sheets.Add("Sheet name with spaces");

// Add link to 2nd sheet
XFile.Workbook.Sheets(1).Cells.Cell(1,1).Hyperlink = "'Sheet name with spaces'!C5";
XFile.Workbook.Sheets(1).Cells.Cell(1,1).HyperlinkType =  4; // hltCurrentWorkbook 

// Add URL link
XFile.Workbook.Sheets(1).Cells.Cell(2,1).Hyperlink = "http://google.com";
XFile.Workbook.Sheets(1).Cells.Cell(2,1).HyperlinkType =  1; // hltURL

// Add UNC link
XFile.Workbook.Sheets(1).Cells.Cell(3,1).Hyperlink = "\\\\server\\share\\myfile.txt";
XFile.Workbook.Sheets(1).Cells.Cell(3,1).HyperlinkType =  3; // hltUNC 

// Add file link
XFile.Workbook.Sheets(1).Cells.Cell(4,1).Hyperlink = "c:\\games\\pool.swf";
XFile.Workbook.Sheets(1).Cells.Cell(4,1).HyperlinkType =  2; // hltFile 


// Save data to output XLS file
XFile.SaveAs ("out.xls");

WScript.Quit(1);
