package com.hiway.demo;

import java.sql.Connection;

import com.hiway.model._MappingKit;
import com.jfinal.config.Constants;
import com.jfinal.config.Handlers;
import com.jfinal.config.Interceptors;
import com.jfinal.config.JFinalConfig;
import com.jfinal.config.Plugins;
import com.jfinal.config.Routes;
import com.jfinal.plugin.activerecord.ActiveRecordPlugin;
import com.jfinal.plugin.activerecord.CaseInsensitiveContainerFactory;
import com.jfinal.plugin.activerecord.dialect.OracleDialect;
import com.jfinal.plugin.activerecord.tx.TxByMethods;
import com.jfinal.plugin.c3p0.C3p0Plugin;
import com.jfinal.render.ViewType;

public class DemoConfig extends JFinalConfig {

	public void configConstant(Constants me) {
		this.loadPropertyFile("jdbc.properties");
		me.setDevMode(true);
		me.setViewType(ViewType.FREE_MARKER);
	}

	@Override
	public void configRoute(Routes me) {
		me.add("/", IndexController.class);
	}

	@Override
	public void configPlugin(Plugins me) {
		C3p0Plugin cp = new C3p0Plugin(this.getProperty("jdbc.url"), this.getProperty("jdbc.username"), this.getProperty("jdbc.password"));
		// 配置Oracle驱动
		cp.setDriverClass(this.getProperty("jdbc.driverClass"));
		me.add(cp);
		ActiveRecordPlugin arp = new ActiveRecordPlugin(cp);
		arp.setTransactionLevel(Connection.TRANSACTION_READ_COMMITTED);
		me.add(arp);
		// 配置Oracle方言
		arp.setDialect(new OracleDialect());
		arp.setShowSql(true);
		// 配置属性名(字段名)大小写不敏感容器工厂
		arp.setContainerFactory(new CaseInsensitiveContainerFactory());
		_MappingKit.mapping(arp);
	}

	@Override
	public void configInterceptor(Interceptors me) {
		me.addGlobalServiceInterceptor(new TxByMethods("save", "update", "delete"));
	}

	@Override
	public void configHandler(Handlers me) {

	}

}
