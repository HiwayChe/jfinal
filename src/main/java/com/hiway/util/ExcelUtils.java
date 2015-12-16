package com.hiway.util;

import java.beans.BeanDescriptor;
import java.beans.BeanInfo;
import java.beans.IntrospectionException;
import java.beans.Introspector;
import java.beans.PropertyDescriptor;
import java.beans.SimpleBeanInfo;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.OutputStreamWriter;
import java.io.PrintWriter;
import java.lang.reflect.InvocationTargetException;
import java.nio.charset.Charset;
import java.text.DecimalFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import javax.servlet.http.HttpServletResponse;

import org.apache.commons.beanutils.PropertyUtils;
import org.apache.commons.collections.CollectionUtils;
import org.apache.commons.io.IOUtils;
import org.apache.commons.lang3.ArrayUtils;
import org.apache.commons.lang3.BooleanUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;

/**
 * 读写Excel的工具类，包括csv和txt。可直接下载Excel
 * Excel单元格只支持字符串和数字类型。
 * @author guangai.che
 * @date Aug 14, 2014
 * 
 */
@SuppressWarnings({"unchecked","rawtypes"})
public class ExcelUtils {

	private static final int DEFAULT_WIDTH = -1;

	private Map<String, ValueConverter> converterMap = new HashMap<String, ValueConverter>(6);
	
	private BeanInfo beanInfo;
	/**
	 * if <b>isMap</b> is true, it means you want to hold data in a Map, <b>keys</b> && <b>values</b> are both STRING type.
	 */
	private boolean isMap = false;
	private Map<String, PropertyDescriptor> propertyDescriptorMap = new HashMap<String, PropertyDescriptor>();
	
	public ExcelUtils(final Class clazz){
		if(Map.class.isAssignableFrom(clazz)){
			this.isMap = true;
			this.beanInfo = new SimpleBeanInfo(){
				@Override
				public BeanDescriptor getBeanDescriptor(){
					return new BeanDescriptor(clazz);
				}
			};
		}else{
			try {
				this.beanInfo = Introspector.getBeanInfo(clazz);
			} catch (IntrospectionException e) {
				throw new IllegalArgumentException(e);
			}
		}
		PropertyDescriptor[] propertyArray = this.beanInfo.getPropertyDescriptors();
		if(ArrayUtils.isNotEmpty(propertyArray)){
			for(PropertyDescriptor desp : propertyArray){
				this.propertyDescriptorMap.put(desp.getName(), desp);
			}
		}
	}
	
	/**
	 * 输出excel到客户端下载
	 * @param response  结束后流自动 关闭
	 * @param fileName可以为空
	 * @param dataList
	 * @param columnList
	 */
	public void downloadExcel(HttpServletResponse response, String fileName, List dataList, List<String> fieldList, List<String> titleList, boolean isCsv){
		if(CollectionUtils.isEmpty(fieldList)){
			return;
		}
		if(CollectionUtils.isEmpty(titleList)){
			titleList = fieldList;
		}
		if(fieldList.size() != titleList.size()){
			return;
		}
		List<Column> columnList = new ArrayList<Column>();
		for(int i=0; i<fieldList.size(); i++){
			columnList.add(new Column(titleList.get(i), fieldList.get(i)));
		}
		this.downloadExcel(response, fileName, dataList, columnList, isCsv);
	}
	
	/**
	 * 输出excel
	 * @param response 自动关闭
	 * @param fileName 可以为空
	 * @param dataList
	 * @param columnList
	 * @param isCsv
	 */
	public void downloadExcel(HttpServletResponse response, String fileName, List dataList, List<Column> columnList, boolean isCsv){
		if(StringUtils.isBlank(fileName)){
			fileName = "download" + System.currentTimeMillis() + ".xls";
		}
		response.setContentType("application/xls");
		response.addHeader("Content-Disposition", "attachment;filename=" + fileName);
		try{
			if(isCsv){
				this.writeLargeData(response.getOutputStream(), dataList, columnList);
			}else{
				this.write(response.getOutputStream(), dataList, columnList);
			}
		}catch(IOException e){
			//ignore
		}finally{
			try{
				response.getOutputStream().close();
			}catch(IOException e){
				//ignore
			}
		}
	}
	
	/**
	 * 写出数据到流，用POI实现，数据量大的时候不要用这个方法
	 * @param stream
	 * @param dataList
	 * @param columnList
	 * @throws IOException
	 */
	public void write(OutputStream stream, List dataList, List<Column> columnList) throws IOException {
		if (stream == null || CollectionUtils.isEmpty(dataList) || CollectionUtils.isEmpty(columnList)) {
			return;
		}
		this.checkNecessaryValueConverter(columnList);
		HSSFWorkbook workbook = new HSSFWorkbook();
		HSSFSheet sheet = workbook.createSheet("sheet-1");
		HSSFRow head = sheet.createRow(0);
		for (int i = 0; i < columnList.size(); i++) {
			HSSFCell cell = head.createCell(i);
			cell.setCellType(HSSFCell.CELL_TYPE_STRING);
			cell.setCellValue(StringUtils.trimToEmpty(columnList.get(i).getName()));
		}
		for (int i = 0; i < dataList.size(); i++) {
			HSSFRow row = sheet.createRow(i + 1);
			this.writeToRow(row, dataList.get(i), columnList);
		}
		for (int i = 0; i < columnList.size(); i++) {
			if (columnList.get(i).getWidth() > 0) {
				sheet.setColumnWidth(i, columnList.get(i).getWidth() * 256);
			} else {
				sheet.autoSizeColumn(i);
			}
		}

		workbook.write(stream);
		stream.flush();
	}
	
	/**
	 * 输出大量的数据，输出格式是csv,纯文本，字符编码gbk.
	 * @param stream
	 * @param dataList
	 * @param columnList
	 * @throws IOException 
	 */
	public void writeLargeData(OutputStream stream, List dataList, List<Column> columnList) throws IOException{
		if (stream == null || CollectionUtils.isEmpty(dataList) || CollectionUtils.isEmpty(columnList)) {
			return;
		}
		this.checkNecessaryValueConverter(columnList);
		PrintWriter pw = new PrintWriter(new OutputStreamWriter(stream, Charset.forName("gbk")));
		for(int i=0; i<columnList.size(); i++){
			this.printCSVCell(pw, columnList.get(i).getName());
			if(i != columnList.size() - 1){
				pw.print("\t");
			}
		}
		pw.println();
		for (int i = 0; i < dataList.size(); i++) {
			for(int j=0; j<columnList.size(); j++){
				Object data = dataList.get(i);
				String property = columnList.get(j).getProperty();
				Object value = this.convertValue(this.getPropertyValue(data, property), this.converterMap.get(property), true);
				this.printCSVCell(pw, value);
				if(j != columnList.size() - 1){
					pw.print("\t");
				}
			}
			pw.println();
		}
		if(pw.checkError()){
			throw new IOException("error occured when write to " + stream.toString()); 
		}
		pw.flush();
	}
	
	/**
	 * 
	 * @param stream
	 * @param columnList
	 * @return
	 * @throws IOException
	 */
	public List read(InputStream stream, List<Column> columnList) throws IOException {
		return this.read(stream, columnList, 1);
	}

	/**
	 * 
	 * @param stream
	 * @param clazz
	 * @param headRows
	 *            标题所占的行数
	 * @return
	 * @throws IOException
	 */
	public List read(InputStream stream, List<Column> columnList, int headRows) throws IOException {
		if (stream == null || CollectionUtils.isEmpty(columnList)) {
			return Collections.EMPTY_LIST;
		}
		this.checkNecessaryValueConverter(columnList);
		HSSFWorkbook workbook = new HSSFWorkbook(stream);
		HSSFSheet sheet = workbook.getSheetAt(0);
		int rowCount = sheet.getPhysicalNumberOfRows();
		if (rowCount <= headRows) {
			return Collections.EMPTY_LIST;
		}
		List dataList = new ArrayList();
		for (int i = headRows; i < rowCount; i++) {
			HSSFRow row = sheet.getRow(i);
			Object t = this.readFromRow(row, columnList);
			dataList.add(t);
		}
		return dataList;
	}
	
	/**
	 * 读取txt中的数据返回List<Map>
	 * @param stream  读取完后不关闭
	 * @param columnName  列名
	 * @param columnType  列的类型 支持String.class, Long.class, Boolean.class
	 * @param delimiter  默认制表符\t
	 */
	public List<Map> readTxt(InputStream stream, List<String> columnName, List<Class> columnType){
		return this.readTxt(stream, columnName, columnType, "\t");
	}
	
	public List<Map> readTxt(InputStream stream, List<String> columnName, List<Class> columnType, String delimiter){
		if(stream == null|| CollectionUtils.isEmpty(columnName) || CollectionUtils.isEmpty(columnType) 
				|| columnName.size() != columnType.size()){
			return Collections.emptyList();
		}
		try{
			List<String> lines = IOUtils.readLines(stream,"UTF-8");
			if(CollectionUtils.isEmpty(lines)){
				return Collections.emptyList();
			}
			List<Map> ret = new ArrayList<Map>(lines.size());
			for(String line : lines){
				if(StringUtils.isBlank(line)){
					continue;
				}
				String[] data = StringUtils.trimToEmpty(line).split(delimiter);
				if(data.length < columnName.size()){//数据不全的忽略
					continue;
				}
				Map params = new HashMap();
				for(int i=0; i<columnName.size(); i++){
					Object convertedData = null;
					if(columnType.get(i) == Long.class){
						convertedData = Long.valueOf(StringUtils.trim(data[i]));
					}else if(columnType.get(i) == Boolean.class){
						convertedData = Boolean.valueOf(StringUtils.trim(data[i]));
					}else{
						convertedData = StringUtils.trim(data[i]);
					}
					params.put(columnName.get(i), convertedData);
				}
				ret.add(params);
			}
			return ret;
		}catch(IOException e){
			//ignore
		}
		return Collections.emptyList();
	}
	
	/**
	 * @param property
	 * @param converter
	 */
	public void registerValueConverter(String property, ValueConverter converter) {
		this.converterMap.put(property, converter);
	}

	private Object readFromRow(HSSFRow row, List<Column> columnList) {
		Object bean = null;
		try{
			bean = this.beanInfo.getBeanDescriptor().getBeanClass().newInstance();
			for (int i = 0; i < columnList.size(); i++) {
				HSSFCell cell = row.getCell(i);
				Object value = this.getCellValue(cell, columnList.get(i).getProperty());
				String property  = columnList.get(i).getProperty();
				this.setPropertyValue(bean, property, this.convertValue(value, this.converterMap.get(property), false));
			}
		} catch (InstantiationException e) {
		} catch (IllegalAccessException e) {
		} catch (InvocationTargetException e) {
		} catch (NoSuchMethodException e) {
		}
		return bean;
	}
	
	private void writeToRow(HSSFRow row, Object data, List<Column> columnList){
		for (int j = 0; j < columnList.size(); j++) {
			HSSFCell cell = row.createCell(j);
			String property = columnList.get(j).getProperty();
			cell.setCellType(this.guessCellType(property));
			Object value = this.getPropertyValue(data, property);
			this.setCellValue(cell, this.convertValue(value, this.converterMap.get(property), true));
		}
	}


	private Object getCellValue(HSSFCell cell, String property) {
		if(StringUtils.isBlank(property)){
			return null;
		}
		int cellType = this.guessCellType(property);
		if(cellType == Cell.CELL_TYPE_NUMERIC){
			return cell.getNumericCellValue();
		}
		return cell.getStringCellValue();
	}
	
	private void setCellValue(Cell cell, Object value) {
		if (value == null) {
			cell.setCellValue("");
			return;
		}
		if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
			cell.setCellValue(Double.valueOf(value.toString()));
		}else{
			cell.setCellValue(value.toString());
		}
	}
	
	private int guessCellType(String property){
		if(this.isMap){
			return Cell.CELL_TYPE_STRING;
		}
		PropertyDescriptor desp = this.propertyDescriptorMap.get(property);
		if(desp == null){
			return Cell.CELL_TYPE_STRING;
		}
		Class type = this.getBoxingType(desp.getPropertyType());
		if(Number.class.isAssignableFrom(type)){
			return Cell.CELL_TYPE_NUMERIC;
		}
		return Cell.CELL_TYPE_STRING;
	}
	
	/**
	 * boolean, byte, char, short, int, long, float, and double 
	 * @param type
	 * @return
	 */
	private Class getBoxingType(Class type){
		if(type.isPrimitive()){
			if(type == boolean.class){
				return Boolean.class;
			}
			if(type == byte.class){
				return Byte.class;
			}
			if(type == char.class){
				return Character.class;
			}
			if(type == short.class){
				return Short.class;
			}
			if(type == int.class){
				return Integer.class;
			}
			if(type == long.class){
				return Long.class;
			}
			if(type == float.class){
				return Float.class;
			}
			if(type == double.class){
				return Double.class;
			}
		}
		return type;
	}
	
	private void printCSVCell(PrintWriter writer, Object value){
		if(value == null){
			writer.print("\"\"");
			return;
		}
		char[] chArray = value.toString().toCharArray();
		writer.print("\"");
		for(char ch : chArray){
			if(ch == '\r' || ch == '\n'){
				continue;
			}else if(ch == '\"'){
				writer.print("\"");
				writer.print("\"");
			}else{
				writer.print(ch);
			}
		}
		writer.print("\"");
	}
	
	private void checkNecessaryValueConverter(List<Column> columnList){
		if(this.isMap){
			return ;
		}
		for(Column column : columnList){
			String property = column.getProperty();
			PropertyDescriptor desp = this.propertyDescriptorMap.get(property);
			Class type = this.getBoxingType(desp.getPropertyType());
			if(type == Date.class && this.converterMap.get(property) == null){
				this.converterMap.put(property, ExcelUtils.dateConverter);
			}
			if(Number.class.isAssignableFrom(type) && this.converterMap.get(property) == null){
				this.converterMap.put(property, ExcelUtils.numberConverter);
			}
			if(type == Boolean.class && this.converterMap.get(property) == null){
				this.converterMap.put(property, ExcelUtils.booleanConverter);
			}
		}
	}
	
	private Object getPropertyValue(Object data, String property) {
		if (StringUtils.isBlank(property)) {
			return "";
		}
		if(this.isMap){
			return ((Map)data).get(property);
		}
		Object value = null;
		try {
			value = PropertyUtils.getSimpleProperty(data, property);
		} catch (IllegalAccessException e) {
		} catch (InvocationTargetException e) {
		} catch (NoSuchMethodException e) {
		}
		return value;
	}
	
	/**
	 * @param bean
	 * @param property
	 * @param value
	 * @throws IllegalAccessException
	 * @throws InvocationTargetException
	 * @throws NoSuchMethodException
	 */
	private void setPropertyValue(Object bean, String property, Object value) throws IllegalAccessException, InvocationTargetException, NoSuchMethodException{
		if(this.isMap){
			((Map)bean).put(property, value);
			return;
		}
		PropertyDescriptor desp = this.propertyDescriptorMap.get(property);
		Class targetType = this.getBoxingType(desp.getPropertyType());
		Object result = value;
		if(value instanceof Long){
			if(targetType == Integer.class){
				result = ((Long)value).intValue();
			}else if(targetType == Short.class){
				result = ((Long)value).shortValue();
			}else if(targetType == Byte.class){
				result = ((Long)value).byteValue();
			}else if(targetType == Float.class){
				result = ((Long)value).floatValue();
			}else if(targetType == Double.class){
				result = ((Long)value).doubleValue();
			}
		}else if(value instanceof Double){
			if(targetType == Float.class){
				result = ((Double)value).floatValue();
			}else if(targetType == Long.class){
				result = ((Double)value).longValue();
			}else if(targetType == Integer.class){
				result = ((Double)value).intValue();
			}else if(targetType == Short.class){
				result = ((Double)value).shortValue();
			}else if(targetType == Byte.class){
				result = ((Double)value).byteValue();
			}
		}
		PropertyUtils.setProperty(bean, property, result);
	}

	/**
	 * convert <b>value</b>
	 * <br/><br/>
	 * ATTENTION: <b>direction</b> is very important
	 * @param value
	 * @return
	 */
	private Object convertValue(Object value, ValueConverter converter, boolean beanToExcel) {
		if (converter == null || value == null || "".equals(value)) {
			return value;
		}
		if(beanToExcel){
			return converter.beanToExcel(value);
		}else{
			return converter.excelToBean(value.toString());
		}
	}

	public static class Column {
		/**
		 * 显示的列名
		 */
		private String name;
		/**
		 * 属性名，用于反射取数据
		 */
		private String property;
		/**
		 * 宽度
		 */
		private int width;

		public Column(String name, String property, int width) {
			this.name = name;
			this.property = property;
			this.width = width;
		}

		public Column(String name, String property) {
			this(name, property, DEFAULT_WIDTH);
		}

		public Column() {

		}

		public String getName() {
			return name;
		}

		public void setName(String name) {
			this.name = name;
		}

		public String getProperty() {
			return property;
		}

		public void setProperty(String property) {
			this.property = property;
		}

		public int getWidth() {
			return width;
		}

		public void setWidth(int width) {
			this.width = width;
		}
	}

	/**
	 * 值转换器<br/>
	 * <b>CELL_TYPE_NUMERIC</b> only supports <b>LONG</b> and <b>DOUBLE</b>, so if you write your 
	 * own <b>NUMERIC</b> ValueConverter, you must return <b>LONG</b> or <b>DOUBLE</b>
	 * @author user
	 * @date   Aug 14, 2014
	 *
	 */
	public static class ValueConverter {
		
		public String beanToExcel(Object value){
			return value.toString();
		}

		public Object excelToBean(String value){
			return value;
		}
	}

	public static ValueConverter dateConverter = new ValueConverter() {
		SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
		
		@Override
		public String beanToExcel(Object value) {
			return sdf.format(value);
		}

		@Override
		public Object excelToBean(String value) {
			try {
				return sdf.parse(value);
			} catch (ParseException e) {
				return null;
			}
		}

	};

	public static ValueConverter numberConverter = new ValueConverter() {
		DecimalFormat df = new DecimalFormat("#.##");
		
		@Override
		public String beanToExcel(Object value) {
			return df.format(value);
		}

		@Override
		public Object excelToBean(String value) {
			try {
				return df.parse(value);
			} catch (ParseException e) {
				return null;
			}
		}

	};
	
	public static ValueConverter booleanConverter = new ValueConverter(){

		@Override
		public String beanToExcel(Object value) {
			return BooleanUtils.toString((Boolean)value, "是", "否", "否");
		}

		@Override
		public Object excelToBean(String value) {
			return BooleanUtils.toBooleanObject(value, "是", "否", "否");
		}

	};
}
