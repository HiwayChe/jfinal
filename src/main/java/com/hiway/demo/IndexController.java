package com.hiway.demo;

import com.jfinal.core.Controller;

public class IndexController extends Controller{

	public void index(){
		this.renderText("Hello Jfinal");
	}
}
