var xfo = WScript.CreateObject( "olexlsf.XLSFileObject" );

var XFile = xfo.XLSFile;

// read data
XFile.OpenFile("in.xls");

// write data to output file
XFile.SaveAs ("out.xls");

WScript.Quit(1);
