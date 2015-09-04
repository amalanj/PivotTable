package csc.infochimps.delivery.es_batch_reporter;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;

import org.apache.commons.lang.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.AreaReference;
import org.apache.poi.hssf.util.CellReference;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataConsolidateFunction;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFPivotTable;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openxml4j.exceptions.InvalidFormatException;

import au.com.bytecode.opencsv.CSVReader;

public class PivotTable {

	public static void main(String[] args) throws FileNotFoundException,
			IOException, InvalidFormatException {
		long startTime = System.currentTimeMillis();
		XSSFWorkbook wb = new XSSFWorkbook();
		XSSFSheet sheet = (XSSFSheet) wb.createSheet("data");

		//Create some data to build the pivot table on
		//setCellData(sheet);
		loadCellData(sheet);

		XSSFSheet pivotSheet = (XSSFSheet) wb.createSheet("pivot");
		XSSFPivotTable pivotTable = pivotSheet.createPivotTable(new AreaReference(
				"A:P"), new CellReference("A3"), sheet);
		//XSSFPivotTable pivotTable = pivotSheet.createPivotTable(new AreaReference(
				//"A:P"), new CellReference("A1"), sheet);
		//XSSFPivotTable pivotTable = sheet.createPivotTable(new AreaReference(
				//"A1:D4"), new CellReference("R5"));
		// Configure the pivot table
		// Use first column as row label
		pivotTable.addRowLabel(1);
		pivotTable.addRowLabel(2);
		pivotTable.addRowLabel(3);
		// Sum up the second column
		pivotTable.addColumnLabel(DataConsolidateFunction.SUM, 5, "Sum of received");
		pivotTable.addColumnLabel(DataConsolidateFunction.SUM, 6, "Sum of ignored");
		pivotTable.addColumnLabel(DataConsolidateFunction.SUM, 7, "Sum of preprocessed");
		pivotTable.addColumnLabel(DataConsolidateFunction.SUM, 8, "Sum of processed");
		pivotTable.addColumnLabel(DataConsolidateFunction.SUM, 9, "Sum of process-failed");
		pivotTable.addColumnLabel(DataConsolidateFunction.SUM, 10, "Sum of ready-to-deliver");
		pivotTable.addColumnLabel(DataConsolidateFunction.SUM, 11, "Sum of delivering");
		pivotTable.addColumnLabel(DataConsolidateFunction.SUM, 12, "Sum of delivered");
		pivotTable.addColumnLabel(DataConsolidateFunction.SUM, 13, "Sum of delivery-failed");
		pivotTable.addColumnLabel(DataConsolidateFunction.SUM, 14, "Sum of UnrecognizedFormat");
		pivotTable.addColumnLabel(DataConsolidateFunction.SUM, 15, "Sum of other");
		// Set the third column as filter
		//pivotTable.addColumnLabel(DataConsolidateFunction.AVERAGE, 2);
		// Add filter on forth column
		pivotTable.addReportFilter(4);
		
		createSummary(wb);

		FileOutputStream fileOut = new FileOutputStream("ooxml-pivottable.xlsx");
		wb.write(fileOut);
		wb.close();
		fileOut.close();
		long endTime = System.currentTimeMillis();
		System.out.println("It took " + (endTime - startTime) / 1000 + " seconds to load and pivot");
		
	}

	private static void createSummary(XSSFWorkbook wb) {
		
		int x,NewCust,SaveRef,Despatch,GR,ASN,Orders,POD,CustCount,NewLoc,InMessages,InOrder,InASN,InGR,InPoD,InDesp,OrderErrored,ASNErrored,GRErrored,DespErrored,PODErrored,GoodCount,BadCount,FixCount;
		long Received=0,Ignored=0,PreProc=0,Processed=0,ProcFailed=0,ReadyToDeliver=0,Delivering=0,Delivered=0,DeliveryFail=0,UnrecFailed=0,Other=0;
		String y,Cust = null,Location = null,SaveCust,Stuff1,Stuff2;
		double Score_i;
		int AtLeastOneKpi_i, HeaderCount;
		
    SaveRef = Despatch = GR = ASN = Orders = InMessages = HeaderCount = 0;
    NewCust = NewLoc = GoodCount = BadCount = FixCount = 1;
    CustCount = 2;
    
    XSSFSheet pivotSheet = wb.getSheet("pivot");
    XSSFSheet summarySheet = (XSSFSheet) wb.createSheet("summary");
    
    SaveRef = 3;
    int rowIndex = 0;
    
    List<XSSFPivotTable> pivotTables = pivotSheet.getPivotTables();
    for(XSSFPivotTable pivotTable : pivotTables){
    	System.out.println(pivotTable);
    }
    
    for (Row row : pivotSheet) {
    	/*while(rowIndex == SaveRef){
    		System.out.println("row Index: "+rowIndex++);
    		continue;
    	}*/
      for (Cell cell : row) {
      	if(cell.getCellType() == Cell.CELL_TYPE_STRING){
      		//System.out.println(cell.getRichStringCellValue().getString());
      		if(cell.getRichStringCellValue().getString().equals("Despatch")){
            Despatch = InDesp = InMessages = 1;
            InOrder = InASN = InGR = InPoD = 0;
      		} else if(cell.getRichStringCellValue().getString().equals("GoodsRec")){
            GR = InGR = InMessages = 1;
            InOrder = InASN = InDesp = InPoD = 0;
      		} else if(cell.getRichStringCellValue().getString().equals("Order")){
            Orders = InOrder = InMessages = 1;
            InGR = InASN = InDesp = InPoD = 0;
      		} else if(cell.getRichStringCellValue().getString().equals("AdvanceShippingNotice")){
            ASN = InASN = InMessages = 1;
            InGR = InOrder = InDesp = InPoD = 0;
      		} else if(cell.getRichStringCellValue().getString().equals("POD")){
            POD = InPoD = InMessages = 1;
            InGR = InOrder = InDesp = InASN = 0;
      		} else{
      			Despatch = GR = Orders = ASN = POD = 0;
      			InMessages = InGR = InOrder = InDesp = InASN = InPoD = 0;
      			SaveRef = row.getRowNum() + 1;
      		}
      		
      		if (InMessages == 0){
      			HeaderCount = HeaderCount + 1;
      			if (HeaderCount == 1){
      				Location = cell.getRichStringCellValue().getString();
      			} else{
      				Cust = Location;
      				Location = cell.getRichStringCellValue().getString();
      			}
      			OrderErrored = ASNErrored = DespErrored = GRErrored = PODErrored = 0;
      		} else{
      			if (HeaderCount > 0){
      				HeaderCount = 0;
      				CustCount = CustCount + 1;
      				Row summaryRow = summarySheet.createRow(CustCount);
      				Cell custCell = summaryRow.createCell(1);
      				custCell.setCellValue(Cust);
      				Cell locCell = summaryRow.createCell(2);
      				locCell.setCellValue(Location);
      			}
      		}
      		if (InMessages == 1){
      			// Received = (long) row.getCell(0).getNumericCellValue();
      		}
      	}
      }
    }
    
		
	}

	public static void loadCellData(XSSFSheet sheet) throws IOException {
		/* Step -1 : Read input CSV file in Java */
		String inputCSVFile = "DataSheet.csv";
		CSVReader reader = new CSVReader(new FileReader(inputCSVFile));
		/* Variables to loop through the CSV File */
		String[] nextLine; /* for every line in the file */
		int lnNum = 0; /* line number */
		/* Step -2 : Define POI Spreadsheet objects */
		HSSFWorkbook new_workbook = new HSSFWorkbook(); // create a blank workbook
																										// object
		//HSSFSheet sheet = new_workbook.createSheet("CSV2XLS"); // create a worksheet
																														// with caption
																														// score_details
		/* Step -3: Define logical Map to consume CSV file data into excel */
		Map<Integer, Object[]> excel_data = new HashMap<Integer, Object[]>(); // create
																																				// a map
																																				// and
																																				// define
																																				// data
		/* Step -4: Populate data into logical Map */
		while ((nextLine = reader.readNext()) != null) {
			excel_data.put(lnNum++, new Object[]{nextLine[0],
					nextLine[1], nextLine[2], nextLine[3], nextLine[4],
					nextLine[5], nextLine[6], nextLine[7], nextLine[8],
					nextLine[9], nextLine[10], nextLine[11], nextLine[12],
					nextLine[13], nextLine[14], nextLine[15]});
		}
		/* Step -5: Create Excel Data from the map using POI */
		Set<Integer> keyset = excel_data.keySet();
		int rownum = 0;
		for (Integer key : keyset) { // loop through the data and add them to the
																// cell
			Row row = sheet.createRow(rownum++);
			Object[] objArr = excel_data.get(key);
			int cellnum = 0;
			for (Object obj : objArr) {
				Cell cell = row.createCell(cellnum++);
				//if (obj instanceof Double)
				if(StringUtils.isNumeric((String) obj))
					cell.setCellValue(Double.parseDouble((String) obj));
				else
					cell.setCellValue((String) obj);
			}
		}
	}
}
