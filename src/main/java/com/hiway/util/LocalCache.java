package com.hiway.util;

import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;
import java.util.concurrent.Callable;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.ScheduledExecutorService;
import java.util.concurrent.TimeUnit;

/**
 * 
 * @author chega 允许脏数据
 * @param <K>
 * @param <V>
 */
public class LocalCache<K, V> {
	private final Map<K, CacheData<V>> cache = new HashMap<K, CacheData<V>>(500);
	ScheduledExecutorService cleanExecutor = Executors.newScheduledThreadPool(1); // 清除过期数据
	ExecutorService queryExecutor = Executors.newFixedThreadPool(1);//查询新数据

	{
		cleanExecutor.scheduleWithFixedDelay(new Runnable() {
			@Override
			public void run() {
				Iterator<K> it = cache.keySet().iterator();
				while (it.hasNext()) {
					K key = it.next();
					CacheData<V> d = cache.get(key);
					if (d != null && d.isExpired()) {
						it.remove();
					}
				}
			}
		}, 0, 3, TimeUnit.SECONDS);
	}

	public void put(K key, V value, Long seconds) {
		CacheData<V> d = new CacheData<V>();
		d.setData(value);
		d.setExpireTime(System.currentTimeMillis() + seconds * 1000);
		this.cache.put(key, d);
	}

	public V get(final K key, final Callable<V> task, final Long seconds) {
		V result = null;
		CacheData<V> d = this.cache.get(key);
		if (d == null) {
			try {
				V v = task.call();
				this.put(key, v, seconds);
				result = v;
			} catch (Exception e) {
				e.printStackTrace();
			}
		} else {
			if (d.isExpired()) {
				//如果数据已经过期，先返回过期数据，再异步取新数据
				queryExecutor.submit(new Runnable() {
					@Override
					public void run() {
						try {
							V v = task.call();
							put(key, v, seconds);
						} catch (Exception e) {
							e.printStackTrace();
						}
					}
				});
				
			}
			result = d.getData();
		}
		return result;
	}

	public V get(K key) {
		CacheData<V> d = this.cache.get(key);
		if (d != null) {
			return d.getData();
		}
		return null;
	}

	public int size() {
		return this.cache.size();
	}
}
