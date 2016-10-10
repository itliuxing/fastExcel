/**
 * 
 */
package com.xing.common.interfacePort;

import java.io.File;
import java.util.List;

import com.xing.common.server.ExportExcelServer;
import com.xing.common.server.ImportExcelServer;
import com.xing.common.server.impl.ExportExcelServerImpl;
import com.xing.common.server.impl.ImportExcelServerImpl;


/**
 * 类描述：   		
 * 创建时间：	2013-9-27 上午09:50:28   
 * 创建人：		LiuXing   
 * 修改备注：   	
 * @version		V1.0 
 */
public class OperateExcelServer {

	private static OperateExcelServer operateExcelServer ;
	
	private OperateExcelServer(){}
	
	public static synchronized OperateExcelServer getInstance(){
		if( operateExcelServer == null ){
			operateExcelServer = new OperateExcelServer() ;
		}
		return operateExcelServer ;
	}
	
	/***
	 * 导入数据
	 * @param <T>
	 * @param excelFile
	 * @param reflectEntity
	 * @param dataEntity
	 * @param ronStartNum
	 * @return
	 * @throws Exception
	 */
	public <T> List<T>  improtExcel(File excelFile, Object reflectEntity, Class<T> dataEntity, Integer ronStartNum)throws Exception{
		ImportExcelServer importExcelServer = new ImportExcelServerImpl();
		return importExcelServer.importExcelServer(excelFile, reflectEntity, dataEntity, ronStartNum);
	}
	
	/***
	 * 导出数据
	 * @param <T>
	 * @param dataList
	 * @param reflectEntity
	 * @param excelFile
	 * @param startRomNum
	 * @return
	 * @throws Exception
	 */
	public <T> String exportExcelServer( List<T> dataList , Object reflectEntity, File excelFile, int startRomNum )throws Exception{
		ExportExcelServer exportExcelServer = new ExportExcelServerImpl() ;
		return exportExcelServer.exportExcelServer(dataList, reflectEntity, excelFile, startRomNum) ;
	}
}
