package evaluate;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.simple.JSONArray;
import org.json.simple.JSONObject;
import org.json.simple.parser.JSONParser;
import org.json.simple.parser.ParseException;

import com.pwnetics.metric.WordSequenceAligner;
import com.pwnetics.metric.WordSequenceAligner.Alignment;





public class Evaluate {
	final static File folder = new File("/Users/apple/Documents/data");
	final static File answerfolder = new File("/Users/apple/Documents/official_answer");
	final static String[] title = new String[]{"_id","name","initial_time","initial_answer","initial_quality","final_time","final_answer","final_quality","improve"};
	
	public static void main(String[] args) throws IOException{
		
		
		writeXLSXFile(folder);
	}
	
	
	public static void loadfolder(){
		
	}

	public static void writeXLSXFile(final File folder) throws IOException {
		
		String excelFileName = "/Users/apple/Documents/evaluate_results.xlsx";//name of excel file

		XSSFWorkbook wb = new XSSFWorkbook();
		
		for (final File fileEntity: folder.listFiles()){
			if(fileEntity.isDirectory()){
				
			}else{
				json2xlsx(wb,fileEntity);
			}
		}
		
		FileOutputStream fileOut = new FileOutputStream(excelFileName);

		//write this workbook to an Outputstream.
		wb.write(fileOut);
		fileOut.flush();
		fileOut.close();
	}
	
	public static void json2xlsx(XSSFWorkbook wb,File jsonFile){
		XSSFSheet page = wb.createSheet(jsonFile.getName()) ;
		
		JSONParser parser = new JSONParser();

		try {
			JSONArray jsonArray = (JSONArray) parser.parse(new FileReader(folder+"/"+jsonFile.getName()));
			
			//Header
			XSSFRow title_row = page.createRow(0);
			for(int i = 0 ; i < title.length; i++){
				XSSFCell cell = title_row.createCell(i);
				cell.setCellValue(title[i]);
			}
			
			System.out.println(jsonFile.getName());
			int rowIndex = 1;
			for(Object obj : jsonArray){
				XSSFRow row = page.createRow(rowIndex);
				
				JSONObject tasks = (JSONObject) obj;
//				System.out.println(jsonFile.getName());
				
				for(int i = 0; i < title.length; i++){
					if(title[i].contains("answer")){
						String answer = (String) tasks.get(title[i]);
						XSSFCell cell = row.createCell(i);
						cell.setCellValue(answer);
						
					}else if(title[i].contains("time")){
						Long time = (Long) tasks.get(title[i]);
						XSSFCell cell = row.createCell(i);
						cell.setCellValue(time);
						
					}else if(title[i].equals("_id")){
						String id = (String) tasks.get(title[i]);
						XSSFCell cell = row.createCell(i);
						cell.setCellValue(id);
						
					}else if(title[i].equals("name")){
						String name = (String) tasks.get(title[i]);
						XSSFCell cell = row.createCell(i);
						cell.setCellValue(name);
					}else if(title[i].contains("quality")){
						WordSequenceAligner werEval = new WordSequenceAligner();
						String[] ref;
						JSONArray answerArray = (JSONArray) parser.parse(new FileReader(answerfolder +"/"+jsonFile.getName()));
						JSONObject answers = (JSONObject) answerArray.get(0);
						ref = ((String) answers.get(title[i-1])).split(" ");
						String[] hyp = ((String) tasks.get(title[i-1])).split(" ");
						Alignment a = werEval.align(ref, hyp);
						float acc = a.getNumCorrect()/(float)a.getReferenceLength();
						
						XSSFCell cell = row.createCell(i);
						cell.setCellValue(acc);
						CellStyle style = wb.createCellStyle();
						style.setDataFormat(wb.createDataFormat().getFormat("0.00%"));
						cell.setCellStyle(style);
						
					}else if(title[i].equals("improve")){
						Long initialTime = (Long) tasks.get("initial_time");
						Long finalTime = (Long) tasks.get("final_time");
						
						float improve = (initialTime - finalTime)/ (float)initialTime;
//						System.out.println(improve);
						XSSFCell cell = row.createCell(i);
						cell.setCellValue(improve);
						CellStyle style = wb.createCellStyle();
						style.setDataFormat(wb.createDataFormat().getFormat("0.00%"));
						cell.setCellStyle(style);
					}
					page.setColumnWidth( i +1, 30*256);
				}
				rowIndex++;
				
			}
		
		
		} catch (FileNotFoundException e) {
			System.err.println("FileNotFoundException");
			e.printStackTrace();
		} catch (IOException e) {
			System.err.println("IOException");
			e.printStackTrace();
		} catch (ParseException e) {
			System.err.println("ParseException");
			e.printStackTrace();
		}
	}
	
}
