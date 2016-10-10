/**
 * 
 */
package com.xing.test;

import java.io.File;
import java.util.ArrayList;
import java.util.List;

import com.xing.common.interfacePort.OperateExcelServer;
import com.xing.test.bean.DataEntity;
import com.xing.test.bean.data;

/**
 * 类描述：   		
 * 创建时间：	2013-9-27 下午02:28:26   
 * 创建人：		LiuXing   
 * 修改备注：   	
 * @version		V1.0 
 */
public class TestOperateExcelServer  {

	public static void improtExcel(){
		File excelFile = new File( "D://a.xls" ) ;
		try {
			List<DataEntity> dataEntityList = OperateExcelServer.getInstance().improtExcel( excelFile, new data(), DataEntity.class , 1 ) ;
			for( DataEntity data : dataEntityList ){
				System.out.println( "name:" + data.getName() + "====age:" + data.getAge() );
			}
			excelFile = new File( "D://c.xls" ) ;
			String path =  OperateExcelServer.getInstance().exportExcelServer(dataEntityList, new data(), excelFile, 0) ;
			//String path =  OperateExcelServer.getInstance().exportExcelServer( new ArrayList<DataEntity>() , new data(), excelFile, 1) ;
			System.out.println( "导出的位置是:" + path );
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
	
	public static void main(String[] args) {
		long start = System.currentTimeMillis() ;
		System.out.println( start );
		improtExcel() ;
		long stop = System.currentTimeMillis() ;
		System.out.println( stop );
		System.out.println( "服务调用花费时间" + (stop - start) );
	}
	
}
