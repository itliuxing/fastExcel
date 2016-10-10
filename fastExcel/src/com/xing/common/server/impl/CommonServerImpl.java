/**
 * 
 */
package com.xing.common.server.impl;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import com.xing.common.exception.ReflectParamterException;
import com.xing.common.server.CommonServer;
import com.xing.common.util.BeanParameterReflect;
import com.xing.common.util.StringUtils;

/**
 * 类描述：   		
 * 创建时间：	2013-9-27 上午10:19:17   
 * 创建人：		LiuXing   
 * 修改备注：   	
 * @version		V1.0 
 */
public class CommonServerImpl implements CommonServer {
	
	private Sheet sheet = null;
	private InputStream inStream = null;
	private Workbook book = null ;
	private Row row = null ;
	private Cell cell = null;

	
	@Override
	public Map<String, Integer> headHandlerServer(File excelFile) throws Exception {
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
		
		try {
			inStream.close() ;
		} catch (Exception e) {}
		
		return marksMapq ;
		
	}
	
	@Override
	public Map<String, Object> reflectEntityHandlerServer(Object reflectEntity)throws Exception{
		Map<String, Object> reflectMap = null ;
		try {
			reflectMap = BeanParameterReflect.reflectPrarameter(reflectEntity);
		} catch (Exception e) {
			throw new ReflectParamterException( e.getMessage() + "映射对象解析映射出现异常..." ) ;
		}
		return reflectMap  ;
	}

	@Override
	public Map<String, Integer> twoLevelHandlerServer(Map<String, Integer> excelMap, Map<String, Object> reflectMap) {
		Map<String,Integer> levelServerMap = new HashMap<String,Integer>() ;
		if( excelMap != null && reflectMap != null ){
			Iterator iterator = reflectMap.keySet().iterator() ;
			while( iterator.hasNext() ){
				String reflectName = iterator.next().toString() ;
				if( excelMap.get( reflectMap.get( reflectName ) ) != null ){
					levelServerMap.put( reflectName , excelMap.get( reflectMap.get( reflectName ) ) ) ;
				}
			}
		}
		return levelServerMap ;
	}

}
