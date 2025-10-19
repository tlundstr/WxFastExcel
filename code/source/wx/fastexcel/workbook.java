package wx.fastexcel;

// -----( IS Java Code Template v1.2

import com.wm.data.*;
import com.wm.util.Values;
import com.wm.app.b2b.server.Service;
import com.wm.app.b2b.server.ServiceException;
// --- <<IS-START-IMPORTS>> ---
import java.io.FileInputStream;
import java.io.IOException;
import org.dhatim.fastexcel.reader.ReadableWorkbook;
import org.dhatim.fastexcel.reader.Sheet;
// --- <<IS-END-IMPORTS>> ---

public final class workbook

{
	// ---( internal utility methods )---

	final static workbook _instance = new workbook();

	static workbook _newInstance() { return new workbook(); }

	static workbook _cast(Object o) { return (workbook)o; }

	// ---( server methods )---




	public static final void workbookOpen (IData pipeline)
        throws ServiceException
	{
		// --- <<IS-START(workbookOpen)>> ---
		// @sigtype java 3.5
		// [i] field:0:required filePath
		// [o] object:0:required FEWorkbook
		// [o] object:0:required FEFirstSheet
		// [o] field:0:required numberOfRowsFirstSheet
		// [o] field:0:required numberOfSheets
		// pipeline
		IDataCursor pipelineCursor = pipeline.getCursor();
		String	filePath = IDataUtil.getString( pipelineCursor, "filePath" );
		pipelineCursor.destroy();
		
		ReadableWorkbook wb = null;
		Sheet sheet = null;
		long numberOfSheets=0;
		long numberOfRowsFirstSheet=0;
		
		try {
			FileInputStream file = new FileInputStream(filePath);
			wb = new ReadableWorkbook(file);
			sheet = wb.getFirstSheet();
			numberOfSheets = wb.getSheets().count();
			numberOfRowsFirstSheet = sheet.openStream().count();
		}
		catch(IOException e) {
				e.printStackTrace();
		}		
		
		// pipeline
		IDataCursor pipelineCursor_1 = pipeline.getCursor();
		Object MSWorkbook = new Object();
		IDataUtil.put( pipelineCursor_1, "FEWorkbook", wb );
		IDataUtil.put( pipelineCursor_1, "FEFirstSheet", sheet );
		IDataUtil.put( pipelineCursor_1, "numberOfRowsFirstSheet", String.valueOf(numberOfRowsFirstSheet));
		IDataUtil.put( pipelineCursor_1, "numberOfSheets", String.valueOf(numberOfSheets));
		
		pipelineCursor_1.destroy();
		// --- <<IS-END>> ---

                
	}
}

