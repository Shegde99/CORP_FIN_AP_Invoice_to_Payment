function readFinData(tc,sheet)
{
  var objExcel = Excel.Open("C:\\Files_for_automation\\CORP_FLOW.xlsx")
  var objSheet = objExcel.SheetByTitle(sheet)
  var rowCnt = objSheet.RowCount
  var colCnt = objSheet.ColumnCount
  var objExlDict = getActiveXObject("Scripting.Dictionary");
  var key, item
  for (let i = 2; i < rowCnt + 1; i++)//row by row
 {
    if (objSheet.Cell(2, i).Value == tc) // 2 is column no
    {
      for (let j = 2; j < colCnt + 1; j++)
      {
        key = objSheet.Cell(j, 1).Value
        item = objSheet.Cell(j, i).Value
        objExlDict.Add(key, item)
      }
      break
    }
  }
  return objExlDict
}

function UniqueBatchName()
{
  var BatchName=aqDateTime.Now()
  BatchName=aqConvert.DateTimeToFormatStr(BatchName,"%m%d%y%H%M%S")
  BatchName="Test"+BatchName
  return BatchName
}

function InvoiceNum()
{
  var InvoiceNum=aqDateTime.Now()
  InvoiceNum=aqConvert.DateTimeToFormatStr(InvoiceNum,"%m%d%y%H%M%S")
  InvoiceNum="Invoice"+InvoiceNum
  return InvoiceNum
}

function AddressName()
{
  var AddressName=aqDateTime.Now()
  AddressName=aqConvert.DateTimeToFormatStr(AddressName,"%m%d%y%H%M%S")
  AddressName="Invoice"+AddressName
  return AddressName
}

function GetText(requestid)
{
  var value;
  value=aqString.SubString(requestid,35,7);
  return value;
}

function GetPdfBatchName(pdfData)
{
  var value;
  value=aqString.SubString(pdfData,1671,36);
  return value;
}

function ValidateOutput()
{
  var page = Sys.Browser("iexplore").Page("http://dsvldfdfw0002.dover-global.net:8000/OA_CGI/FNDWRR.exe?temp_id=*");
  var picture = page.PagePicture();
  Log.Message("Verify the xml document for the output pattern!!");
  Log.Picture(picture);
}
