package com.hiway.util;

public class CacheData<V> {
	private V data;
	private Long expireTime;
	
	public V getData() {
		return data;
	}
	public void setData(V data) {
		this.data = data;
	}
	
	public Long getExpireTime() {
		return expireTime;
	}
	public void setExpireTime(Long expireTime) {
		this.expireTime = expireTime;
	}
	public boolean isExpired(){
		return System.currentTimeMillis() > expireTime;
	}
}
