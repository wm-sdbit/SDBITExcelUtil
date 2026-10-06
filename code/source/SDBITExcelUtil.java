

// -----( IS Java Code Template v1.2

import com.wm.data.*;
import com.wm.util.Values;
import com.wm.app.b2b.server.Service;
import com.wm.app.b2b.server.ServiceException;
// --- <<IS-START-IMPORTS>> ---
import com.softwareag.util.IDataMap;
import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.lang.invoke.MethodHandles;
import java.text.SimpleDateFormat;
import java.util.Base64;
import java.util.Optional;
import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVPrinter;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import com.wm.data.IData;
import com.wm.data.IDataCursor;
import com.wm.data.IDataUtil;
// --- <<IS-END-IMPORTS>> ---

public final class SDBITExcelUtil

{
	// ---( internal utility methods )---

	final static SDBITExcelUtil _instance = new SDBITExcelUtil();

	static SDBITExcelUtil _newInstance() { return new SDBITExcelUtil(); }

	static SDBITExcelUtil _cast(Object o) { return (SDBITExcelUtil)o; }

	// ---( server methods )---




	public static final void convertExcelToCsv (IData pipeline)
        throws ServiceException
	{
		// --- <<IS-START(convertExcelToCsv)>> ---
		// @sigtype java 3.5
		// [i] object:0:required fileExcel
		// [i] field:0:required enableLogging
		// [o] field:0:required fileCSV
		IDataMap dataMap = new IDataMap(pipeline);
		
		Object fileXLSX = dataMap.get("fileExcel");
		boolean enableLogging = Boolean.parseBoolean(dataMap.getAsString("enableLogging"));
		
		addLogMessage("Start converting Excel to CSV", enableLogging);
		
		String csv = "";
		
		try {
			InputStream is = new ByteArrayInputStream((byte[]) fileXLSX);
		
			csv = convert(is, Optional.empty(), Optional.empty(), Optional.empty(), pipeline);
		} catch (IOException e) {
			addLogMessage(e.getMessage(), true);
			
			throw new ServiceException(e);
		}
		
		addLogMessage("Excel file converted to CSV", enableLogging);	
		
		dataMap.put("fileCSV", csv);
		// --- <<IS-END>> ---

                
	}



	public static final void convertExcelToCsv_b64 (IData pipeline)
        throws ServiceException
	{
		// --- <<IS-START(convertExcelToCsv_b64)>> ---
		// @sigtype java 3.5
		// [i] field:0:required fileExcel_b64
		// [i] field:0:required enableLogging
		// [o] field:0:required fileCSV
		IDataMap dataMap = new IDataMap(pipeline);
		
		String fileXLSX_b64 = dataMap.getAsString("fileExcel_b64");
		boolean enableLogging = Boolean.parseBoolean(dataMap.getAsString("enableLogging"));
		
		addLogMessage("Start converting Excel to CSV", enableLogging);
		
		String csv = "";
		
		try {
			byte[] file = Base64.getDecoder().decode(fileXLSX_b64);
			
			InputStream is = new ByteArrayInputStream((byte[]) file);
		
			csv = convert(is, Optional.empty(), Optional.empty(), Optional.empty(), pipeline);
		} catch (IOException e) {
			addLogMessage(e.getMessage(), true);
			
			throw new ServiceException(e);
		}
		
		addLogMessage("Excel file converted to CSV", enableLogging);	
		
		dataMap.put("fileCSV", csv);
		// --- <<IS-END>> ---

                
	}

	// --- <<IS-START-SHARED>> ---
		
	private static String convert(InputStream is,  Optional<Integer> firstDataRowOpt, Optional<Integer> lastDataRowOpt,
	        Optional<Integer> stopAfterEmptyLinesOpt, IData pipeline) throws IOException {
	    try (Workbook workbook = WorkbookFactory.create(is)) {
	        Sheet sheet = workbook.getSheetAt(0);
	
	        int firstDataRow = firstDataRowOpt.isPresent() ? firstDataRowOpt.get() : 0;
	        int lastRowNum = sheet.getLastRowNum() + 1;
	        int end = lastDataRowOpt.isPresent() ? lastDataRowOpt.get() + 1 : lastRowNum;
	        int stopAfterEmptyLines = stopAfterEmptyLinesOpt.isPresent() ? stopAfterEmptyLinesOpt.get() : 0;
	
	        StringBuilder csvSb = new StringBuilder();
	        CSVPrinter csvPrinter = new CSVPrinter(csvSb, CSVFormat.RFC4180);
	        SimpleDateFormat sdf = new SimpleDateFormat("yyyyMMdd_HHmmss");
	
	        StopConversionManager skipEmptyRow = new StopConversionManager(stopAfterEmptyLines);
	        int x = firstDataRow;
	
	        while (x < end && !skipEmptyRow.hasToStopConversion(sheet.getRow(x))) {
	            Row excelRow = sheet.getRow(x);
	
	            int columns = excelRow.getLastCellNum();
	            boolean notEmptyRow = false;
	
	            for (int y = 0; y < columns; y++) {
	                Cell cell = excelRow.getCell(y);
	                String cellValue = getCellValue(cell, sdf);
	                
	                notEmptyRow = notEmptyRow || cellValue != null && cellValue.trim() != "";
	            }
	            
	            if (notEmptyRow) {
		            for (int y = 0; y < columns; y++) {
		                Cell cell = excelRow.getCell(y);
		                String cellValue = getCellValue(cell, sdf);
		
		                if (cellValue == null) {
		                    cellValue = "";
		                } else {
		                    cellValue = cellValue.trim();
		                }
		
		                csvPrinter.print(cellValue);
		            }
		
		            csvPrinter.println();
	            }
	
	            x++;
	        }
	        
	        String strRtn = csvSb.toString();
	        
	        csvPrinter.close();
	
	        return strRtn;
	    } catch (EncryptedDocumentException e1) {
			addLogMessage("The document is encrypted with password", true);
	
	        throw e1;
	    }
	}
	
	private static String getCellValue(Cell cell, SimpleDateFormat sdf) {
	    String result = "";
	
	    if (cell != null) {
	        DataFormatter df = new DataFormatter();
	
	        Workbook workbook = cell.getSheet().getWorkbook();
	
	        FormulaEvaluator fe = workbook.getCreationHelper().createFormulaEvaluator();
	
	        result = df.formatCellValue(cell, fe);
	
	        if (cell.getCellType() == CellType.NUMERIC && DateUtil.isCellDateFormatted(cell)) {
	            result = sdf.format(cell.getDateCellValue());
	        }
	    }
	
	    return result;
	}
	
	private static void addLogMessage(String message, boolean enableLogging) {
		if (!enableLogging) {
			return;
		}
		
		String function = MethodHandles.lookup().lookupClass().getName();
		
		try {			
			StackTraceElement[] elements = Thread.currentThread().getStackTrace();
			
			if (elements.length >= 3) {
				StackTraceElement element = elements[2];
				
				function = function + "." + element.getMethodName();
			}
		} catch (Exception e) {
		}
		
		try {
			Values values = new Values();
			
			values.put("message", message);
												
			values.put("function", function);
						
			Service.doInvoke("pub.flow", "debugLog", (IData) values);
			
			//Service.doInvoke("pub.flow", "debugLog", pipeline);
		} catch (Exception e) {
		}
	}
	
	private static class StopConversionManager {
	    private int numberOfEmptyLines = 0;
	
	    private final int stopAfterEmptyLines;
	
	    public StopConversionManager(int stopAfterEmptyLines) {
	        this.stopAfterEmptyLines = stopAfterEmptyLines;
	    }
	
	    public boolean hasToStopConversion(Row row) {
	        if (stopAfterEmptyLines == 0) {
	            return false;
	        }
	
	        if (row == null) {
	            numberOfEmptyLines++;
	        } else {
	            numberOfEmptyLines = 0;
	        }
	
	        return numberOfEmptyLines >= stopAfterEmptyLines;
	    }
	}
	// --- <<IS-END-SHARED>> ---
}

