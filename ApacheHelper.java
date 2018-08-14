
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FilenameFilter;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Constructor;
import java.math.BigDecimal;
import java.net.URL;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.attribute.BasicFileAttributeView;
import java.nio.file.attribute.BasicFileAttributes;
import java.security.NoSuchAlgorithmException;
import java.sql.Timestamp;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Collection;
import java.util.Comparator;
import java.util.Date;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.StringTokenizer;
import java.util.TreeMap;
import java.util.regex.Pattern;

import javax.net.ssl.HttpsURLConnection;
import javax.net.ssl.SSLContext;

import org.apache.commons.io.FileUtils;
import org.apache.poi.hssf.util.AreaReference;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.struts.util.LabelValueBean;



public class ApacheHelper  {

	final static int MIN_CENSUS_AGE = 17;
	final static int MAX_CENSUS_AGE = 85;
	

	
	
	
	public static synchronized Workbook getWorkbook(String pType,String pRaterTemplateName, String pPlanVersion) throws Exception {
			
			Workbook wks = null;
			FileInputStream fstream = null;
			boolean remoteFetch =false;
			File directory = null;
			try {
				Path p = Paths.get(TEMP_LOCATION + pRaterTemplateName);
				BasicFileAttributes view = null;
				try {
					view = Files.getFileAttributeView(p, BasicFileAttributeView.class).readAttributes();
				} catch (Exception e) {
					remoteFetch =true;
				}
			    
			   Calendar cal =Calendar.getInstance();
			   if(view != null) 
				   cal.setTimeInMillis(view.lastModifiedTime().toMillis());
			   SimpleDateFormat sdf = new SimpleDateFormat("MMM dd");	
	
			   if(remoteFetch || !sdf.format(cal.getTime()).equals(sdf.format(Calendar.getInstance().getTime()))){
				   
					String url = getApacheURL() + pRaterTemplateName;
					url = url.replaceAll(" ", "%20");
					URL urlid = new URL(url);
					
					SSLContext sc = SSLContext.getInstance("TLSv1.2"); //$NON-NLS-1$
					sc.init(null, null, new java.security.SecureRandom());
					
					HttpsURLConnection conn = (HttpsURLConnection) urlid.openConnection();
					
					conn.setSSLSocketFactory(sc.getSocketFactory());
					
					conn.setRequestMethod("GET");
					conn.setDoOutput(false);
					
					directory = new File(TEMP_LOCATION);
					
					if(!directory.exists()){
						directory.mkdir();
					}else{
						File[] files = null;
						try {
							// Assuming directory is a directory!!
							FilenameFilter textFilter = new FilenameFilter() {
								public boolean accept(File dir, String name) {
									String lowercaseName = name.toLowerCase();
									if (lowercaseName.endsWith(".tmp")) {
										return true;
									} else {
										return false;
									}
								}
							};
					        files = directory.listFiles(textFilter);
					        if (files == null) {  // null if security restricted
					        	Logger.log("Failed to list contents of " + directory);
					        }
					        IOException exception = null;
					        for (File file : files) {
					            try {
					                FileUtils.forceDelete(file);
					            } catch (IOException ioe) {
					                exception = ioe;
					            }
					        }
					        if (null != exception) {
					            throw exception;
					        }
						} catch (IOException e) {
							Logger.log("Error occurred while cleaning the tmp files.");
						}finally{
							files =null;
						}
					}
					
					InputStream inputStream = conn.getInputStream();
					FileOutputStream outputStream = new FileOutputStream(TEMP_LOCATION + pRaterTemplateName);
					int bytesRead = -1;
					byte[] buffer = new byte[4096];
					while ((bytesRead = inputStream.read(buffer)) != -1) {
						outputStream.write(buffer, 0, bytesRead);
					}
					
					outputStream.flush();
					outputStream.close();
					inputStream.close();
					conn.disconnect();
				   }
					fstream = new FileInputStream(TEMP_LOCATION + pRaterTemplateName);
					
					wks = new XSSFWorkbook(fstream);
				
			} catch (IOException e) {
				
				throw new ManualRateException(RaterHelper.getRaterSycodeDescription( "RaterError", "ErrorMsg4",new String[]{pPlanVersion}, ""));
			}
			catch (Exception e) {
				throw new ManualRateException(RaterHelper.getRaterSycodeDescription( "RaterError", "ErrorMsg5", "")); 
			}finally{
				if (fstream != null)
					fstream.close();
				fstream = null;
				directory = null;
			}
			return wks;
		}
	

	/**
	 * Set cell values to the Cells
	 * @param pWorkSheet  WorkSheet 
	 * @param pCellValue Value of the Cell
	 * @param pTble_Valu Name of the cell
	 */
	public static void setCellValue (String pRaterType, Sheet pWorkSheet, Object pCellValue, String pTble_Valu){
		CellReference cellReference = new CellReference(JSPUtils.getSysCodeDescription(pRaterType, pTble_Valu));
	    Row row = pWorkSheet.getRow(cellReference.getRow());
	    Cell cell = row.getCell(cellReference.getCol());
	    if(pCellValue instanceof Double){
	    	cell.setCellValue((double)pCellValue);
	    }
	    else {
	    	cell.setCellValue(String.valueOf(pCellValue));
	    }
	}
	
	/**
	 * Set cell values to the Cells with reference
	 * @param pWorkSheet
	 * @param pCellValue
	 * @param pCellReference
	 */
	public static void setCellValue (Sheet pWorkSheet, Object pCellValue, String pCellReference){
		CellReference cellReference = new CellReference(pCellReference);
	    Row row = pWorkSheet.getRow(cellReference.getRow());
	    Cell cell = row.getCell(cellReference.getCol());
	    if(pCellValue instanceof Double){
	    	cell.setCellValue((double)pCellValue);
	    }
	    else {
	    	cell.setCellValue(String.valueOf(pCellValue));
	    }
	}
	
	public static String getCellValue (Sheet worksheet, String cellName) throws ManualRateException{
		Object retVal = null;
		CellReference cellReference = new CellReference(cellName);
	    Row row = worksheet.getRow(cellReference.getRow());
	    Cell cell = row.getCell(cellReference.getCol());

	    if (cell!=null) {
	    	int type = cell.getCellType();
	        switch (type) {
	            case Cell.CELL_TYPE_BOOLEAN:
	                retVal = cell.getBooleanCellValue();
	                break;
	            case Cell.CELL_TYPE_NUMERIC:
	            	retVal = cell.getNumericCellValue();
	                break;
	            case Cell.CELL_TYPE_STRING:
	            	retVal = cell.getStringCellValue();
	                break;
	            case Cell.CELL_TYPE_FORMULA:
	            	switch(cell.getCachedFormulaResultType()) {
	                	case Cell.CELL_TYPE_NUMERIC:
	                		retVal = cell.getNumericCellValue();
	                		break;
	                	case Cell.CELL_TYPE_STRING:
	                		retVal = cell.getStringCellValue();
	                		break;
	                	case Cell.CELL_TYPE_BLANK:
	    	                break;
	    	            case Cell.CELL_TYPE_ERROR:
	    	            	retVal = cell.getErrorCellValue();
	            	}
	                break;
	            case Cell.CELL_TYPE_BLANK:
	                break;
	            case Cell.CELL_TYPE_ERROR:
	            	retVal = cell.getErrorCellValue();
	                break;
	        }
	    }
		if(retVal.equals("#VALUE!")|| retVal.equals("#N/A"))
			throw new ManualRateException(getRaterSycodeDescription("RaterError", "ErrorMsg3", "")); 
	    return String.valueOf(retVal);
	}
	
	public static String getIntCellValue(Sheet worksheet, String cellName) throws ManualRateException {
		String rtnValue = getCellValue(worksheet, cellName);
		if (JSPUtils.convertToDouble(rtnValue) > 0) {
			return JSPUtils.removeDecimal(rtnValue, true);
		}
		return rtnValue;
	}
	
	public static void removeCellFormula(String pRaterType, Sheet pWorkSheet, String pTble_Valu) {

		CellReference cellReference = new CellReference(JSPUtils.getSysCodeDescription(pRaterType, pTble_Valu));
		Row row = pWorkSheet.getRow(cellReference.getRow());
		Cell cell = row.getCell(cellReference.getCol());
		cell.setCellFormula(null);
	}
	
	public static String getTmpFileName() {
		String temp = Math.random() + "";
		return temp.replace(".", "") + ".tmp";
	}
	
		
	/**
	 * Set cell Style to the Cells
	 * @param pWorkSheet  WorkSheet 
	 * @param pTble_Valu Name of the cell
	 * @param pStyle style for the cell
	 */
	public static void setCellStyle (String pRaterType, Sheet pWorkSheet, String pTble_Valu, String pStyle){
		CellReference cellReference = new CellReference(JSPUtils.getSysCodeDescription(pRaterType, pTble_Valu));
	    Row row = pWorkSheet.getRow(cellReference.getRow());
	    Cell cell = row.getCell(cellReference.getCol());
	    CellStyle style = pWorkSheet.getWorkbook().createCellStyle();
		style.setDataFormat(pWorkSheet.getWorkbook().createDataFormat().getFormat(pStyle));
		cell.setCellStyle(style);
	}
	
	public static void setCellStyle(String pRaterType, Sheet pWorkSheet, String pTble_Valu, short pAlignment, int pCellType) {
		CellReference cellReference = new CellReference(JSPUtils.getSysCodeDescription(pRaterType, pTble_Valu));
		Row row = pWorkSheet.getRow(cellReference.getRow());
		Cell cell = row.getCell(cellReference.getCol());
		CellStyle style = pWorkSheet.getWorkbook().createCellStyle();
		style.setAlignment(pAlignment);
//		style.setDataFormat(pWorkSheet.getWorkbook().createDataFormat().getFormat("");
		cell.setCellStyle(style);
		cell.setCellType(pCellType);
	}
	
	public static Cell getCell(String pRaterType, Sheet pWorkSheet, String pTble_Valu) {
		CellReference cellReference = new CellReference(JSPUtils.getSysCodeDescription(pRaterType, pTble_Valu));
		Row row = pWorkSheet.getRow(cellReference.getRow());
		return row.getCell(cellReference.getCol());
	}

	/**
	 * Replace text in the passed String using the passes indices and replacement text.
	 * 
	 * @param originalText the original string
	 * @param replacementStart the starting position for the replacement
	 * @param replacementEnd the ending position for the replacement
	 * @param replacementText the new text to add to the original string
	 * @return
	 */
	private static String replaceSubstring(String originalText, int replacementStart,
			int replacementEnd, String replacementText)
	{
		StringBuffer finalText =
			new StringBuffer(originalText.substring(0, replacementStart));
		finalText.append(replacementText);
		finalText.append(originalText.substring(replacementEnd+1));
		
		return finalText.toString();
	}
		
	public static List<String> getStringAsList(String pString) {

		if (pString != null)
			pString = pString.replace("[", "").replace("]", "");
		List<String> list = new ArrayList<String>();

		StringTokenizer clsToken = new StringTokenizer(pString, ",");
		while (clsToken.hasMoreTokens()) {
			list.add(new String(clsToken.nextToken().trim()));
		}
		return list;
	}
	
	
 
	
	public static String removeInValidFileNameChar(String pFileName){
		String newFileName = pFileName.replaceAll(INVALID_FILE_CHAR_REG, EMPTY_STRING);
		if(newFileName.length()==0)
		   return "iRapRater";
		
		return newFileName;
	}
	
	public static String getData(Map<String, String> pData, int pProductID, int pProductElementID){
		
		String theData = pData.get(pProductID + PIPE + pProductElementID);
		
		return theData == null ? EMPTY_STRING : theData;
	}
	
		
	public static void evaluateFormulasOnSheet (Workbook workbook, Set<Integer> pSheetSet){
		FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
		Sheet worksheet = null;
		for (Integer sheetIndex : pSheetSet) {
			worksheet = workbook.getSheetAt(sheetIndex.intValue());
			for(Row row : worksheet) {
		        for(Cell cell : row) {
		            if(cell.getCellType() == Cell.CELL_TYPE_FORMULA) {
						evaluator.evaluateFormulaCell(cell);
		            }
		        }
		    }
		}
		evaluator =null;//for GC
	}	
	

	  
		
	/**
	 * Refresh work book from apache server to server temp location.
	 * @param pType
	 * @param pRaterVersion
	 * @param pTemplateName
	 */
	public synchronized static void resetWorkbookCache(String pType,String pRaterVersion, String pTemplateName){
		
		File tmpFile = null;
		try {
		List<LabelValueBean> versionlist= StrutsUtils.loadCodeTable(pRaterVersion);
		String pPlanVersion =null;
		for (LabelValueBean labelValueBean : versionlist) {
			pPlanVersion = labelValueBean.getValue();
			if (null == pPlanVersion || pPlanVersion.trim().equals("-"))
				continue;
			
			String url = getApacheURL() + getRaterSycodeDescription(pType, pTemplateName, new String[]{pPlanVersion}, "");
			url = url.replaceAll(SPACE, "%20");
			URL urlid = new URL(url);
			
			SSLContext sc = SSLContext.getInstance("TLSv1.2"); //$NON-NLS-1$
			sc.init(null, null, new java.security.SecureRandom());
			
			HttpsURLConnection conn = (HttpsURLConnection) urlid.openConnection();
			
			conn.setSSLSocketFactory(sc.getSocketFactory());
			
			conn.setRequestMethod("GET");			
			conn.setDoOutput(false);
			
			tmpFile = new File(TEMP_LOCATION);
			
			if(!tmpFile.exists()){
				tmpFile.mkdir();//create new DIR if not exists
			}
			
			InputStream inputStream = conn.getInputStream();
			FileOutputStream outputStream = new FileOutputStream(TEMP_LOCATION + getRaterSycodeDescription(pType, pTemplateName, new String[]{pPlanVersion}, ""));
			int bytesRead = -1;
			byte[] buffer = new byte[4096];
			while ((bytesRead = inputStream.read(buffer)) != -1) {
				outputStream.write(buffer, 0, bytesRead);
			}
			outputStream.flush();
			outputStream.close();
			inputStream.close();
			conn.disconnect();
		}
		} catch (Exception e) {
			Logger.log(e);
		}finally{
			tmpFile = null;
		}
	}
	/**
	 * Refresh work book from apache server to server temp location.
	 * @param pType
	 * @param pRaterVersion
	 * @param pTemplateName
	 * @throws NoSuchAlgorithmException 
	 */
	public synchronized static void resetWorkbookCache(String pTemplateName) {

		File tmpFile = null;
		try {

			String url = getApacheURL() + pTemplateName;
			url = url.replaceAll(SPACE, "%20");
			URL urlid = new URL(url);
			
			SSLContext sc = SSLContext.getInstance("TLSv1.2"); //$NON-NLS-1$
			sc.init(null, null, new java.security.SecureRandom());
			
			HttpsURLConnection conn = (HttpsURLConnection) urlid.openConnection();
			
			conn.setSSLSocketFactory(sc.getSocketFactory());

			conn.setRequestMethod("GET");
			conn.setDoOutput(false);

			tmpFile = new File(TEMP_LOCATION);

			if (!tmpFile.exists()) {
				tmpFile.mkdir();// create new DIR if not exists
			}

			InputStream inputStream = conn.getInputStream();
			FileOutputStream outputStream = new FileOutputStream(TEMP_LOCATION + pTemplateName);
			int bytesRead = -1;
			byte[] buffer = new byte[4096];
			while ((bytesRead = inputStream.read(buffer)) != -1) {
				outputStream.write(buffer, 0, bytesRead);
			}
			outputStream.flush();
			outputStream.close();
			inputStream.close();
			conn.disconnect();
		} catch (Exception e) {
			Logger.log(e);
		} finally {
			tmpFile = null;
		}
	}
}
