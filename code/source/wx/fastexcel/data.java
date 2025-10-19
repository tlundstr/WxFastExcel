package wx.fastexcel;

// -----( IS Java Code Template v1.2

import com.wm.data.*;
import com.wm.util.Values;
import com.wm.app.b2b.server.Service;
import com.wm.app.b2b.server.ServiceException;
// --- <<IS-START-IMPORTS>> ---
import com.softwareag.is.util.DateUtil;
import org.dhatim.fastexcel.reader.Sheet;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Optional;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.stream.Stream;
import org.dhatim.fastexcel.reader.Cell;
import org.dhatim.fastexcel.reader.CellType;
import org.dhatim.fastexcel.reader.Row;
// --- <<IS-END-IMPORTS>> ---

public final class data

{
	// ---( internal utility methods )---

	final static data _instance = new data();

	static data _newInstance() { return new data(); }

	static data _cast(Object o) { return (data)o; }

	// ---( server methods )---




	public static final void dataCellsToDocumentList (IData pipeline)
        throws ServiceException
	{
		// --- <<IS-START(dataCellsToDocumentList)>> ---
		// @sigtype java 3.5
		// [i] object:0:required FESheet
		// [i] field:0:optional firstRowAsHeader {"true","false"}
		// [i] field:0:optional colStart
		// [i] field:0:optional rowStart
		// [i] field:0:optional colEnd
		// [i] field:0:optional rowEnd
		// [i] field:0:optional datePattern {"yyyyDDdd","yyyyMMddHHmmss","dd.MM.yyyy","dd.MM.yyyy HH:mm:ss","dd.MM.yyyy HH:mm:ss.sss"}
		// [o] object:0:required FESheet
		// [o] record:1:required DocumentList
		// [o] field:0:required colStart
		// [o] field:0:required rowStart
		// [o] field:0:required colEnd
		// [o] field:0:required rowEnd
		//pipeline
		IDataCursor pipelineCursor = pipeline.getCursor();
		Object	FESheet = IDataUtil.get( pipelineCursor, "FESheet" );
		int	colStart = getIntegerFromString(IDataUtil.getString( pipelineCursor, "colStart" ),-1);
		int	rowStart = getIntegerFromString(IDataUtil.getString( pipelineCursor, "rowStart"),-1 );
		int	colEnd = getIntegerFromString(IDataUtil.getString( pipelineCursor, "colEnd"),-1 );
		int	rowEnd = getIntegerFromString(IDataUtil.getString( pipelineCursor, "rowEnd"),-1 );
		Boolean	firstRowAsHeader = IDataUtil.getBoolean( pipelineCursor, "firstRowAsHeader" , true);
		String datePattern = IDataUtil.getString(pipelineCursor, "datePattern" );
		pipelineCursor.destroy();
		// pipeline 
		
		Sheet sheet = (Sheet) FESheet;
		ArrayList<IData> DocList = new ArrayList<IData>();
		
		if(sheet == null) throw new ServiceException("Sheet is missing or null");
		
		try{
			
			if(rowStart == -1) rowStart = 0; //If Nothing given, we start at first row..
			if(colStart == -1) colStart = 0; //f Nothing given, we start at first col..
			if(datePattern==null) datePattern = "dd.MM.yyyy HH:mm:ss";
			
			HashMap<Integer, String> headerMap = new HashMap<Integer, String>();
			int currentRowNo = 0;
			
			Stream<Row> rowStream = sheet.openStream();
			Iterator<Row> rowIterator = rowStream.iterator();
			while (rowIterator.hasNext()) {
				Row lrow = (Row) rowIterator.next();
				if(firstRowAsHeader && currentRowNo == 0){
					//Create the headermap
					getHeaderFields(lrow,headerMap);
				} else {
					//if(currentRowNo.get() >= rowStart && currentRowNo.get() < rowEnd){
					if(currentRowNo >= rowStart && (rowEnd == -1 || currentRowNo < rowEnd)){
						DocList.add(getRowAsIData(lrow,colStart,colEnd, datePattern,headerMap ));
					}
				}
				currentRowNo++;
			}
		
		
		}
		catch(Exception e){
			throw new ServiceException(e);
		}
		//Create the IData List
		IData[] DocumentList = new IData[DocList.size()];
		DocumentList = DocList.toArray(DocumentList);
		
		// pipeline
		IDataCursor pipelineCursor_1 = pipeline.getCursor();
		IDataUtil.put( pipelineCursor_1, "FESheet", sheet );
		IDataUtil.put( pipelineCursor_1, "DocumentList", DocumentList );
		IDataUtil.put( pipelineCursor_1, "colStart", Integer.toString(colStart ));
		IDataUtil.put( pipelineCursor_1, "rowStart", Integer.toString(rowStart ));
		IDataUtil.put( pipelineCursor_1, "colEnd", Integer.toString(colEnd ));
		IDataUtil.put( pipelineCursor_1, "rowEnd", Integer.toString(rowEnd ));
		pipelineCursor_1.destroy();
		// pipeline
		// --- <<IS-END>> ---

                
	}

	// --- <<IS-START-SHARED>> ---
	public static int getIntegerFromString(String value, int defaultValue){
		int i = defaultValue;
		
		try {
			i = Integer.parseInt(value);
		} catch (NumberFormatException e) {
			// TODO Auto-generated catch block
			
		}
		
		return i;
	}
	
	public static HashMap<Integer, String> getHeaderFields (Row pHeaderRow, HashMap<Integer, String> pHeaderMap){
		//Function to get all the cell for this row and populate a headermap for it
		
		Stream<Cell> cellStream = pHeaderRow.stream();
		Iterator<Cell> cellIterator = cellStream.iterator();
		int currentColumnNo=0;
		while (cellIterator.hasNext()) {
		    Cell currentCell = (Cell) cellIterator.next();
		    
		    if (currentCell == null){
		    	pHeaderMap.put(currentColumnNo, "Field"+currentColumnNo);
		    } else{
				//Check for duplicates
				String tmp = getCellValueAsString(currentCell,""); //Ignore the datepattern
				
				Boolean vCheck = true;
				int mSuffix = 0;
				while (vCheck == true){
					mSuffix++;
					vCheck = false;
					Iterator<Integer> itr2 = pHeaderMap.keySet().iterator();
					while (itr2.hasNext()) {
					    Integer key = itr2.next();
					    String cellvalue = pHeaderMap.get(key);
					    if (cellvalue.equals(tmp)){
					    	tmp = tmp.concat("_"+Integer.toString(mSuffix));
							vCheck = true;
					    }
					}
				}
				pHeaderMap.put(currentColumnNo, tmp);
			}//else 
			currentColumnNo++;
		}//while
		return pHeaderMap;
	}//getHeaderFields
	
	public static IData getRowAsIData (Row pDataRow,int	pColStart,int	pColEnd, String pPattern, HashMap<Integer, String> pHeaderMap){
		//Function to get all the cell for this row and populate a headermap for it
		IData lreturnIData = IDataFactory.create();
		IDataCursor retCursor = lreturnIData.getCursor();
	
		if (pDataRow != null){
			//We are getting cell values... 
	       	Stream<Cell> cellStream = pDataRow.stream();
			Iterator<Cell> cellIterator = cellStream.iterator();
			int currentColumnNo=0;
			while (cellIterator.hasNext()) {
				Cell currentCell = (Cell) cellIterator.next();
				if(currentColumnNo >= pColStart && (pColEnd == -1 || currentColumnNo < pColEnd)){
					String cellValue = getCellValueAsString(currentCell, pPattern);
					IDataUtil.put(retCursor, pHeaderMap.getOrDefault(currentColumnNo, "Field"+currentColumnNo).toString(),cellValue);
				}//if
	       		currentColumnNo++;
			}//while	       	
		}//if
		
		return lreturnIData;
	}//getRowAsIData
	
	
	public static String getCellValueAsString (Cell cell, String pattern)
	{
		String value = "";
		value = cell.getRawValue();
		/*
		LocalDateTime date = null;
	
		switch (cell.getType()) {
	        case STRING:
	        	value = cell.getText();
	            break;
	        case NUMBER:
	            if (DateUtil.isCellDateFormatted(cell)) {
	            	value = cell.getText();
					DateFormat formatter = new SimpleDateFormat(pattern);
					value = formatter.format(value);
	            } else {
	            	value = cell.getText();
	            }
	            break;
	        case BOOLEAN:
	        	value = cell.getText();
	            break;
		    case FORMULA:
		    	
		    try {
				FormulaEvaluator evaluator = cell.getSheet().getWorkbook().getCreationHelper().createFormulaEvaluator();
				DataFormatter df = new DataFormatter();
				value = df.formatCellValue(cell, evaluator).toString();
			} catch (Exception e) {
				try {
					if (DateUtil.isCellDateFormatted(cell)) {
							date = cell.getDateCellValue();
							DateFormat formatter = new SimpleDateFormat(pattern);
							value = formatter.format(date);
					}
				} catch (Exception e1) {
	
				} 
			
			}
	
			    break;
	        
	        default:
	        	value = "";
		}
		*/
		return value;
	}
	// --- <<IS-END-SHARED>> ---
}

