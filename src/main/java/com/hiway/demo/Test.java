package com.hiway.demo;

import java.beans.PropertyVetoException;

import com.jfinal.kit.PropKit;
import com.jfinal.plugin.activerecord.dialect.OracleDialect;
import com.jfinal.plugin.activerecord.generator.Generator;
import com.mchange.v2.c3p0.ComboPooledDataSource;

public class Test {

	public static void main(String...args){
		PropKit.use("jdbc.properties");
		ComboPooledDataSource dataSource = new ComboPooledDataSource();
		dataSource.setJdbcUrl(PropKit.get("jdbc.url"));
		dataSource.setUser(PropKit.get("jdbc.username"));
		dataSource.setPassword(PropKit.get("jdbc.password"));
		try {dataSource.setDriverClass(PropKit.get("jdbc.driverClass"));}
		catch (PropertyVetoException e) {throw new RuntimeException("no driver class found");} 
		dataSource.setMaxPoolSize(2);
		dataSource.setMinPoolSize(1);
		dataSource.setInitialPoolSize(1);
		dataSource.setMaxIdleTime(1);
		dataSource.setAcquireIncrement(0);

		String dir = "D:\\workspace\\jfinal\\src\\main\\java\\com\\hiway\\model";
		Generator generator = new Generator(dataSource, "com.hiway.model", dir, "com.hiway.model", dir);
		generator.setDialect(new OracleDialect());
		generator.generate(); 
	}
}

