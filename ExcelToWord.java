import java.io.File;
import java.io.FileOutputStream;
import java.io.FileInputStream;

import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.xwpf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

public class WordMaker 
{
	//private static int pCount = 0;
	private static ArrayList<XWPFParagraph> content = new ArrayList<XWPFParagraph>();
	private static XWPFDocument doc = new XWPFDocument();
	private static XWPFParagraph activePara = null;
	private static XWPFRun activeRun = null;
	
	public static void main(String[] args) throws Exception
	{
		System.out.println("Start");
		
		FileInputStream ExcelFileToRead = new FileInputStream("Z:\\Pending Preliminary Listings\\FADI Excel Listings/1769_CaseiNedda_20140827 copy.xlsx"); //selects target Excel file to be read
        XSSFWorkbook wb = new XSSFWorkbook(ExcelFileToRead);
        XSSFSheet sheet = wb.getSheetAt(0);
        XSSFRow row; 
        XSSFCell cell;
		
        int collectionName = -1, collectionId = -1, accessionDate = -1, cont1 = -1, cont1Start = -1, cont1End = -1, 
        	cont2 = -1, cont2Start = -1, cont2End = -1, series = -1, subseries = -1, subsubseries = -1, heading = -1, 
        	description = -1, medium = -1, form = -1, dateExpression = -1, namedEntities = -1, beginDate = -1, endDate = -1;
        
        //Iterate through the first row and assign indexes of columns to variables
        //Might want to convert to all caps later for quality control
        
        XSSFRow firstRow = sheet.getRow(0);
        Iterator<Cell> firstCells = firstRow.cellIterator();
        
        while (firstCells.hasNext())
        {
            cell = (XSSFCell) firstCells.next();   
            if (cell.getStringCellValue().equals("Collection Name"))
                collectionName = cell.getColumnIndex();
            else if(cell.getStringCellValue().equals("Collection ID"))
                collectionId = cell.getColumnIndex();
            else if(cell.getStringCellValue().equals("Accession Date"))
	            accessionDate = cell.getColumnIndex();
            else if(cell.getStringCellValue().equals("Cont 1"))
                cont1 = cell.getColumnIndex();  	
            else if(cell.getStringCellValue().equals("Cont 1 Start"))
                cont1Start = cell.getColumnIndex();            	
            else if(cell.getStringCellValue().equals("Cont 1 End"))
                cont1End = cell.getColumnIndex();            	
            else if(cell.getStringCellValue().equals("Cont 2"))
                cont2 = cell.getColumnIndex();           	
            else if(cell.getStringCellValue().equals("Cont 2 Start"))
                cont2Start = cell.getColumnIndex();            	
            else if(cell.getStringCellValue().equals("Cont 2 End"))
                cont2End = cell.getColumnIndex();            	
            else if(cell.getStringCellValue().equals("Series"))
                series = cell.getColumnIndex();            	
            else if(cell.getStringCellValue().equals("Subseries"))
                subseries = cell.getColumnIndex();            	
                //Make sure this is how they would write sub-sub-series on the form
            else if(cell.getStringCellValue().equals("Subsubseries"))
                subsubseries = cell.getColumnIndex();            	
            else if(cell.getStringCellValue().equals("Heading"))
                heading = cell.getColumnIndex();           	
            else if(cell.getStringCellValue().equals("Description"))
                description = cell.getColumnIndex();           	
            else if(cell.getStringCellValue().equals("Medium"))
                medium = cell.getColumnIndex();            	
            else if(cell.getStringCellValue().equals("Form"))
                form = cell.getColumnIndex();
            else if(cell.getStringCellValue().equals("Date Expression"))
                dateExpression = cell.getColumnIndex();            	
            else if(cell.getStringCellValue().equals("Named Entities"))
                namedEntities = cell.getColumnIndex();
        }
        //Now need to iterate through all the rows while adding to the word doc
        int boxTracker;
        String seriesTracker;
        //there might not be a sub-series
        String subseriesTracker;
        //Add the collectionName, collectionId, accessionDate, and "Preliminary Listing" to the doc
		makeNewRun("C", 0, 0);
		activeRun.setText(sheet.getRow(1).getCell(collectionName).getStringCellValue());
		activeRun.addBreak();
		activeRun.setText("" + (int) sheet.getRow(1).getCell(collectionId).getNumericCellValue());
		activeRun.addBreak();
		
		//checking if there are different accession dates
		Iterator<Row> accessionRows = sheet.rowIterator();
		String currentAccessionDate = sheet.getRow(1).getCell(accessionDate).getStringCellValue();
		String accessionExpression = currentAccessionDate;
		while(!row.getCell(accessionDate).getStringCellValue().isEmpty()) 
		{
			if(!(currentAccessionDate.equals(row.getCell(accessionDate).getStringCellValue()))) {
				accessionExpression =  accessionExpression + ", " + row.getCell(accessionDate).getStringCellValue();
				currentAccessionDate = row.getCell(accessionDate).getStringCellValue();
			}
			row = (XSSFRow) accessionRow.next();
			
		}
		activeRun.setText(accessionExpression);	
		
		activeRun.addBreak();
		activeRun.setText("Preliminary Listing");	
		
		int romanNum, itemNum;
		char subLetter;
		
		Iterator<Row> rows = sheet.rowIterator();
		row = (XSSFRow) rows.next();
    	row = (XSSFRow) rows.next();
        
		System.out.println("Looping...");
		int loopCount = 0;
		
        while(!row.getCell(collectionName).getStringCellValue().isEmpty())
        {	
        	boxTracker = (int) row.getCell(cont1Start).getNumericCellValue();
        	
        	makeNewRun("L", 0, 0);
        	
        	if(row.getCell(cont1Start).getNumericCellValue() == row.getCell(cont1End).getNumericCellValue())
				activeRun.setText(row.getCell(cont1).getStringCellValue() + " " + (int) row.getCell(cont1Start).getNumericCellValue());
        	else 
				activeRun.setText(row.getCell(cont1).getStringCellValue() + " " + (int) row.getCell(cont1Start).getNumericCellValue() 
									+ "-" + (int) row.getCell(cont1End).getNumericCellValue() +"]");
			
        	romanNum = 1;
            while(row.getCell(cont1Start).getNumericCellValue() == boxTracker) 
            {
            	seriesTracker = row.getCell(series).getStringCellValue();
            	
            	makeNewRun("L", 1, 0);
            	//Print out series with Roman numerals
        		activeRun.setText(toRomanNum(romanNum++) + ". " + row.getCell(series).getStringCellValue());

            	subLetter = 65;
        		while(row.getCell(series).getStringCellValue().equals(seriesTracker)) 
            	{
            		subseriesTracker = row.getCell(subseries).getStringCellValue();
        			
            		//Print out sub-series with capital letters
        			if(!row.getCell(subseries).getStringCellValue().isEmpty())
        			{
        				makeNewRun("L", 2, 0);
        				activeRun.setText(subLetter++ + ". " + row.getCell(subseries).getStringCellValue());
        			}

            		itemNum = 1;
            		while(subseries == -1 || row.getCell(subseries).getStringCellValue().equals(subseriesTracker)) 
            		{
            			//System.out.println("Trapped!");
            			System.out.println(++loopCount + " items added...");
            			
            			makeNewRun("L", 3, itemNum);
            			//Print out all the heading + description + media + form + dateExpression + namedEntities
            			//still need to check to see if things are empty or not
            			//still need to check to see if things are empty or not
            			String headerAndDetails = "";
            			if(!row.getCell(heading).getStringCellValue().isEmpty())
            			{
            				String headerString = row.getCell(heading).getStringCellValue();
            				if(headerString.charAt(0) == '"') 
            				{
            					headerString = headerString.substring(0,headerString.length()-1);
	            				
	            				if(!row.getCell(description).getStringCellValue().isEmpty())
	            					headerAndDetails = headerString + ",\" ";
	            				else
	            					headerAndDetails = headerString + "\" ";
            				}else if(!row.getCell(description).getStringCellValue().isEmpty())
	            					headerAndDetails = headerString + ", ";
            				
            			}
            			if(!row.getCell(description).getStringCellValue().isEmpty())
            				headerAndDetails = headerAndDetails + row.getCell(description).getStringCellValue();
            			if(!row.getCell(medium).getStringCellValue().isEmpty())
            				headerAndDetails = headerAndDetails + ", " + row.getCell(medium).getStringCellValue();
            			if(!row.getCell(form).getStringCellValue().isEmpty())
            				headerAndDetails = headerAndDetails + ", " + row.getCell(form).getStringCellValue();
            			if(!row.getCell(namedEntities).getStringCellValue().isEmpty())
            				headerAndDetails = headerAndDetails + "; " + row.getCell(namedEntities).getStringCellValue();
            			if(!row.getCell(beginDate).getStringCellValue().isEmpty())
            			{	
            				headerAndDetails = headerAndDetails + row.getCell(beginDate).getStringCellValue();
            				if(row.getCell(beginDate).getStringCellValue().equals(row.getCell(endDate).getStringCellValue()))
            					headerAndDetails = headerAndDetails + "-" + row.getCell(endDate).getStringCellValue() + ".";
            			}else
            				headerAndDetails = headerAndDetails + " N.D.";
            			//if(headerAndDetails.charAt(headerAndDetails.length() - 1) != '.')
            				//headerAndDetails = headerAndDetails + ".";
            			activeRun.setText(itemNum++ + ". " + headerAndDetails);
            			//Print out and align right "[F. " + cont2Start + "]" if the start and end are the same 
            			//Print out and align right "[F. " + cont2Start +"-" + cont2End +"]" if the start and end are different
            			makeNewRun("R", 0, 0);
            			char contentTwo = '?';
            			if(!row.getCell(cont2).getStringCellValue().isEmpty())
            				contentTwo = row.getCell(cont2).getStringCellValue().charAt(0);
            			
            			if(row.getCell(cont2Start).getNumericCellValue() == row.getCell(cont2End).getNumericCellValue())
            				activeRun.setText("[" + contentTwo + ". " + (int) row.getCell(cont2Start).getNumericCellValue() + "]");
            			else 
            				activeRun.setText("[" + contentTwo + ". " + (int) row.getCell(cont2Start).getNumericCellValue() + "-" 
            									+ (int) row.getCell(cont2End).getNumericCellValue() +"]");
            			row = (XSSFRow) rows.next();
            			
            			if(row.getCell(collectionName).getStringCellValue().isEmpty())
            				break;
            		}		      		
            		if(row.getCell(collectionName).getStringCellValue().isEmpty())
            			break;
            	}
        		if(row.getCell(collectionName).getStringCellValue().isEmpty())
        			break;
            }
        }
		wb.close();
		
        System.out.println("Scanning Complete.");
        
		try
		{
			FileOutputStream out = new FileOutputStream(new File("C:\\Users\\student\\Desktop\\Hatchet/target2.docx"));
			doc.write(out);
			out.close();
		}
		catch(Exception e)
		{
			System.out.println(e);
		}
		System.out.println("Mission accomplished!");
	}
	
	private static void makeNewRun(String pAlign, int indentFactor, int bulletNumber)
	{
		content.add(doc.createParagraph());
    	activePara = content.get(content.size() - 1);
    	if(pAlign.toUpperCase().equals("L"))
    	{
    		activePara.setAlignment(ParagraphAlignment.LEFT);
    		activePara.setIndentationLeft((indentFactor * 360) + ((digits(bulletNumber) + 2) * 90)); //720 unit = 0.5 inch
        	activePara.setIndentationHanging((digits(bulletNumber) + 2) * 90);
    	}
    	else if(pAlign.toUpperCase().equals("R"))
    		activePara.setAlignment(ParagraphAlignment.RIGHT);
    	else if(pAlign.toUpperCase().equals("C"))
    		activePara.setAlignment(ParagraphAlignment.CENTER);
    	else
    		System.out.println("Improper Format! Defaulting to LEFT Alignment.");
    	activePara.setSpacingAfter(80);
    	activePara.createRun();
    	activeRun = activePara.getRuns().get(0);
    	activeRun.setFontFamily("Times New Roman");
	}
	
	private static String toRomanNum(int i) //supports all integers from [1, 50)
	{
		String ret = "";
		if(i / 10 == 4)
			ret = ret + "XL";
		else
			for(int j = 0; j < (i / 10); j++)
				ret = ret + "X";
		if(i % 5 == 4)
		{
			if(i % 10 == 9)
				ret = ret + "IX";
			else
				ret = ret + "IV";
		}
		else
		{
			if((i / 5) % 2 == 1)
				ret = ret + "V";
			for(int j = 0; j < (i % 5); j++)
				ret = ret + "I";
		}
		return ret;
	}
	
	private static int digits(int i) //supports all integers from [0, 1000)
	{
		int ret = 0;
		if(i / 1 > 0)
			ret++;
		if(i / 10 > 0)
			ret++;
		if(i / 100 > 0)
			ret++;
		return ret;
	}
}