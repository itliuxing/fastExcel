/**
 * 
 */
package com.xing.common.server.impl;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import com.xing.common.exception.ReflectParamterException;
import com.xing.common.server.CommonServer;
import com.xing.common.server.ImportExcelServer;
import com.xing.common.util.StringUtils;

/**
 * 类描述：   		
 * 创建时间：	2013-9-27 上午10:05:42   
 * 创建人：		LiuXing   
 * 修改备注：   	
 * @version		V1.0 
 */
public class ImportExcelServerImpl implements ImportExcelServer {
	
	private Sheet sheet = null;
	private InputStream inStream = null;
	private Workbook book = null ;
	private Row row = null ;
	private Cell cell = null;

	@Override
	public <T> List<T> importExcelServer(File excelFile, Object reflectEntity,Class<T> dataEntity, int rowStartNum)throws Exception {
		CommonServer commonServer = new CommonServerImpl();
		//Map<String,Integer> excelMap = commonServer.headHandlerServer(excelFile) ;
		Map<String,Integer> excelMap = this.headHandlerServer(excelFile) ;
		Map<String, Object> reflectMap =commonServer.reflectEntityHandlerServer(reflectEntity) ;
		Map<String, Integer> twoLevelDataMap = commonServer.twoLevelHandlerServer(excelMap, reflectMap) ;
		return dataPaddingServer(excelFile, dataEntity, (rowStartNum<1 ? 1 : rowStartNum) , twoLevelDataMap);
	}
	
	private <T> List<T> dataPaddingServer( File excelFile, Class<T> dataEntity , int rowStartNum , Map<String,Integer> twoLevelDataMap )throws Exception{
		List<T> dataList = new ArrayList<T>() ;
		/*try {
			inStream = new FileInputStream( excelFile );
			book = new HSSFWorkbook(inStream);
		} catch (Exception e) {
			throw new ReflectParamterException( e.getMessage() + "文件不存在，或格式不正确。" ) ;
		}
		sheet = book.getSheetAt( 0 );*/
		int lastRowNum = sheet.getLastRowNum() ;
		//dataEntity dataEntitys = new dataEntity() ;
		
		for( int i=rowStartNum ; i<=lastRowNum ; i++ ){
			Object dataEntitys = dataEntity.newInstance() ;
			row = sheet.getRow(i) ;
			if( row != null ){		//只有当这一行不为空时，才对每个数据列去检查已有的数据
				Iterator<String> iter = twoLevelDataMap.keySet().iterator() ; 
				while( iter.hasNext() ){
					String obj = iter.next() ;
					int cellRum = twoLevelDataMap.get(obj) ;
					cell = row.getCell( cellRum ) ;
					for ( Field field : dataEntity.getDeclaredFields() ) {
						if ( field.getName().equals(obj) ) {
							field.setAccessible( true ) ;
							try {
								Object objFiled = importDataToFiled(cell) ;
								if( objFiled != null && objFiled.toString().length() > 0 ){
									field.set( dataEntitys, objFiled ) ;
								}
							} catch (Exception e) {
								// 给字段赋值的时候出现了，数据格式不匹配，哥们不处理了，为什么呢。
								//都说了Excel单元格的格式必须跟映射的java数据类的格式一样的撒,你做不到，难倒还怪别人呀，是吧.....
								throw new ReflectParamterException( 
										new StringBuffer().
										append( e.getMessage() ).
										append(":java对象的").
										append( field.getName() ).
										append("属性格式与Excel格式不同....").toString() 
								) ;
							}
						}
					}
				}
			}
			dataList.add( (T)dataEntitys ) ;
		}

		try {
			inStream.close() ;
		} catch (Exception e) {}
		
		return dataList ;
	}
	
	/***
	 * 获取格式化后的数据值
	 * @param cell
	 * @return
	 */
	private Object importDataToFiled( Cell cell ){
		int type = cell.getCellType() ;
		switch (type) {
			case 0:
				if(DateUtil.isCellDateFormatted(cell)){
					return cell.getDateCellValue();
				}else {
					//type_cn = "NUMBER";
					return cell.getNumericCellValue();
				}
			case 1:
				//type_cn = "STRING";
				return cell.getStringCellValue();
			case 2:
				//type_cn = "FORMULA";
				return cell.getCellFormula();
			case 3:
				//type_cn = "BLANK";
				return cell.getStringCellValue();
			case 4:
				//type_cn = "BOOLEAN";
				return cell.getBooleanCellValue();
			case 5:
				//type_cn = "ERROR";
				return cell.getErrorCellValue();
			default:
				return "" ;
		}
	}
	
	private Map<String, Integer> headHandlerServer(File excelFile) throws Exception {
		try {
			inStream = new FileInputStream( excelFile );
			book = new HSSFWorkbook(inStream);
		} catch (Exception e) {
			throw new ReflectParamterException( e.getMessage() + "文件不存在，或格式不正确。" ) ;
		}
		sheet = book.getSheetAt( 0 );
		row = sheet.getRow( 0 );

		Map<String,Integer> marksMapq = new HashMap<String,Integer>() ;
		if( row != null ){
			int remark = row.getLastCellNum() ;
			for( int i=0 ; i<remark ; i++ ){
				cell = row.getCell(i) ;
				if( cell != null ){
					String paramter = cell.getStringCellValue() ;
					if( StringUtils.isNoEmpty(paramter) ){
						marksMapq.put( paramter, i ) ;
					}
				}
			}
		}
		
		return marksMapq ;
		
	}

}
