/**
 * 
 */
package com.xing.common.server.impl;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import com.xing.common.exception.ReflectParamterException;
import com.xing.common.server.CommonServer;
import com.xing.common.server.ExportExcelServer;
import com.xing.common.util.StringUtils;
import com.xing.config.StaticProperties;

/**
 * 类描述：   		
 * 创建时间：	2013-9-27 下午04:55:48   
 * 创建人：		LiuXing   
 * 修改备注：   	
 * @version		V1.0 
 */
public class ExportExcelServerImpl implements ExportExcelServer {

	private Sheet sheet = null;
	private InputStream inStream = null;
	private FileOutputStream ouputStream  = null ;
	private Workbook book = null ;
	private Row row = null ;
	private Cell cell = null;
	
	@Override
	public <T> String exportExcelServer(List<T> dataList, Object reflectEntity, File excelFile ,int startRomNum ) throws Exception {
		CommonServer commonServer = new CommonServerImpl();
		Map<String,Integer> excelMap = this.headHandlerServer(excelFile) ;
		Map<String, Object> reflectMap =commonServer.reflectEntityHandlerServer(reflectEntity) ;
		Map<String, Integer> twoLevelDataMap = commonServer.twoLevelHandlerServer(excelMap, reflectMap) ;
		return exportExcel(dataList, twoLevelDataMap , startRomNum);
	}
	
	private  <T> String exportExcel( List<T> dataList , Map<String,Integer> twoLevelDataMap ,int startRomNum)throws Exception{
		//int lastRowNum = sheet.getLastRowNum() ;
		Map<Integer, Cell> map = new HashMap<Integer, Cell>() ;
		int rowStart = startRomNum ;
		for( T t : dataList ){
			row = sheet.getRow( rowStart ) ;
			if( row==null ){
				row = sheet.createRow( rowStart ) ;
			}
			Iterator<String> iter = twoLevelDataMap.keySet().iterator() ; 
			while( iter.hasNext() ){
				String obj = iter.next() ;
				int cellRum = twoLevelDataMap.get(obj) ;
				cell = row.getCell( cellRum ) ;
				if( cell == null ){
					cell = row.createCell( cellRum ) ;
					if( map.get(cellRum) != null ){
						cell.setCellStyle( map.get(cellRum).getCellStyle() ) ;
					}
				}else{
					map.put(cellRum, cell) ;
				}
				for ( Field field : t.getClass().getDeclaredFields() ) {
					if ( field.getName().equals(obj) ) {
						seCellStyle(cell, field) ;
						field.setAccessible(true) ;
						if( field.get(t) != null ){		//防止当数据列为空时，放入excel单元格的时候报错
							cellSetValue( cell, field.get(t) ) ;
						}
						break ;
					}
				}
			}
			rowStart ++ ;
		}
		StringBuffer path ;
		if( StringUtils.isNoEmpty( StaticProperties.FILE_SAVE_PATH ) ){
			path = new StringBuffer( StaticProperties.FILE_SAVE_PATH ) ;	//外部配置路径(提供给外部修改的路径)
		}else{
			path = new StringBuffer( System.getProperty("user.dir") ) ; 	//模块自动配置路径
		}
		if( StringUtils.isNoEmpty( StaticProperties.FILE_SAVE_NAME ) ){
			path.append( StaticProperties.FILE_SAVE_NAME ) ;	//外部配置文件名称(提供给外部修改的路径)
		}else{
			path.append( System.currentTimeMillis() ) ; 		//模块自动生成文件名称
		}
		ouputStream = new FileOutputStream( path.append(".xls").toString() ) ;
		book.write(ouputStream);
		try {
			if (ouputStream != null) {
				ouputStream.close();
				ouputStream = null ;
			}
			if (inStream != null) {
				inStream.close();
				inStream = null ;
			}
			System.out.println("文件写入完成！");
		} catch (Exception e) {}
		return path.toString() ;
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
	
	/***
	 * 设置value
	 * @param cell
	 * @param value
	 */
	private void cellSetValue(Cell cell,Object value ){
		int type = cell.getCellType();
		Double cell_value = null ;
		if( value != null ){
			try {
				cell_value = Double.parseDouble(value.toString()) ;
				cell.setCellValue( cell_value ) ;
			} catch (NumberFormatException e) {
				cell.setCellValue( value.toString() ) ;
			}
		}
		switch (type) {
			case 0:
				cell.setCellValue( Double.parseDouble( value.toString() ) ) ;
				break;
			case 1:
				//文本型
				cell.setCellValue( value.toString() ) ;
				break;
			case 2:
				//公式型
				cell.setCellValue( Double.parseDouble(value.toString()) ) ;
				break;
			case 3:
				//空白型
				cell.setCellValue( value.toString() ) ;
				break;
			case 4:
				//boolean型
				cell.setCellValue( Boolean.parseBoolean( value.toString() ) ) ;
				break;
			case 5:
				//错误
				cell.setCellValue(  value.toString()  ) ;
				break;
		}
	}
	
	private void seCellStyle( Cell cell,Field field ){
		//System.out.println( field.getType().getName() );
		///cell.setCellStyle(1) ;
		//http://www.cnblogs.com/pengfeisun/archive/2011/07/10/2102221.html
	}
}
