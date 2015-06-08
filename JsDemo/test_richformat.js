var xfo = WScript.CreateObject( "olexlsf.XLSFileObject" );

var XFile = xfo.XLSFile;

// Set cells' value and rich format
XFile.Workbook.Sheets(1).Cells.Cell(1,1).Value = "aaabbbcccdddaaa";
XFile.Workbook.Sheets(1).Cells.Cell(1,1).RichFormat = "1-4(style:biu;size:20)";
XFile.Workbook.Sheets(1).Cells.Cell(2,1).Value = "aaabbbcccdddaaa";
XFile.Workbook.Sheets(1).Cells.Cell(2,1).RichFormat = "3-4(style:b)";
XFile.Workbook.Sheets(1).Cells.Cell(3,1).Value = "aaabbbcccdddaaa";
XFile.Workbook.Sheets(1).Cells.Cell(3,1).RichFormat = "1-3(style:b;color:$0000ff);5-10(size:20;font:Courier New;)";
XFile.Workbook.Sheets(1).Cells.Cell(4,1).Value = "aaabbbcccdddaaa";
XFile.Workbook.Sheets(1).Cells.Cell(4,1).RichFormat = "5-6(size:20;color:$FF0000;style:b)";
XFile.Workbook.Sheets(1).Cells.Cell(5,1).Value = "aaabbbcccdddaaa";
XFile.Workbook.Sheets(1).Cells.Cell(5,1).RichFormat = "5-6(script:super;style:b);8-9((script:sub;style:iu))";


XFile.SaveAs ("out.xls");

WScript.Quit(1);
