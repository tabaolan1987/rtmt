/**
 * Copyright (c) CMG Ltd All rights reserved.
 *
 * This software is the confidential and proprietary information of CMG
 * ("Confidential Information"). You shall not disclose such Confidential
 * Information and shall use it only in accordance with the terms of the
 * license agreement you entered into with CMG.
 */

package com.cmg.hipspot.data.dao;

import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import javax.jdo.JDOObjectNotFoundException;
import javax.jdo.PersistenceManager;
import javax.jdo.Query;
import javax.jdo.Transaction;

import com.cmg.hipspot.data.jdo.FeedbackModel;
import com.cmg.hipspot.data.jdo.FeedbackModelJDO;
import org.codehaus.jackson.JsonGenerationException;
import org.codehaus.jackson.JsonParseException;
import org.codehaus.jackson.annotate.JsonAutoDetect.Visibility;
import org.codehaus.jackson.annotate.JsonMethod;
import org.codehaus.jackson.map.JsonMappingException;
import org.codehaus.jackson.map.ObjectMapper;

import com.cmg.hipspot.data.Mirrorable;
import com.cmg.hipspot.util.PersistenceManagerHelper;
import com.cmg.hipspot.util.UUIDGenerator;


/**
 *  T is JDO class E is mirror class
 *  
 *  Query with format
        [WHERE <filter>]
        [VARIABLES <variable declarations>]
        [PARAMETERS <parameter declarations>]
        [<import declarations>]
        [GROUP BY <grouping>]
        [ORDER BY <ordering>]
        [RANGE <start>, <end>]
 * 
 * @Creator Hai Lu
 * @author $Author$
 * @version $Revision$
 * @Last changed: $LastChangedDate$
 */
public class DataAccess<T, E> implements InDataAccess<T, E> {
	private final Class<T> clazzT;
	private final Class<E> clazzE;

	/**
	 *  @param clazzT
	 * @param clazzE
     */
	public DataAccess(Class<T> clazzT, Class<E> clazzE) {
		this.clazzT = clazzT;
		this.clazzE = clazzE;
	}
	/**
	 * 
	 * @param obj
	 * @return
	 * @throws JsonParseException
	 * @throws JsonMappingException
	 * @throws JsonGenerationException
	 * @throws IOException
	 * @throws DataAccessException
	 */
	protected T from(final E obj) throws JsonParseException,
			JsonMappingException, JsonGenerationException, IOException,
			DataAccessException {
		verifyObject(obj);
		ObjectMapper om = new ObjectMapper().setVisibility(JsonMethod.FIELD, Visibility.ANY);
		om.configure(org.codehaus.jackson.map.DeserializationConfig.Feature.FAIL_ON_UNKNOWN_PROPERTIES, false);
		T out = om.readValue(om.writeValueAsString(obj), clazzT);
		verifyObject(out);
		
		return out;
	}
	/**
	 * 
	 * @param obj
	 * @return
	 * @throws JsonParseException
	 * @throws JsonMappingException
	 * @throws JsonGenerationException
	 * @throws IOException
	 * @throws DataAccessException
	 */
	protected E to(final T obj) throws JsonParseException, JsonMappingException,
			JsonGenerationException, IOException, DataAccessException {
		//verifyObject(obj);
		ObjectMapper om = new ObjectMapper().setVisibility(JsonMethod.FIELD, Visibility.ANY);
		om.configure(org.codehaus.jackson.map.DeserializationConfig.Feature.FAIL_ON_UNKNOWN_PROPERTIES, false);
		return om.readValue(om.writeValueAsString(obj), clazzE);
	}
	/**
	 * 
	 * @param obj
	 * @throws DataAccessException
	 */
	protected void verifyObject(Object obj) throws DataAccessException {
		if (obj instanceof Mirrorable) {
			if (((Mirrorable) obj).getId() == null || ((Mirrorable) obj).getId().length() == 0) 
				((Mirrorable) obj).setId(UUIDGenerator.generateUUID());			
		} else {
			throw new DataAccessException(
					"The object must implement interface Mirrorable");
		}
	}
	
	/**
	 * 
	 * @param obj
	 * @return
	 * @throws Exception
	 */
	public boolean put(E obj) throws Exception {
		verifyObject(obj);
		String id = ((Mirrorable) obj).getId();
		if (checkExistence(id))
			delete(id);
		return create(obj);
	}
	/**
	 * 
	 * @param obj
	 * @return
	 * @throws Exception
	 */
	public boolean create(E obj) throws Exception {
		verifyObject(obj);
		PersistenceManager pm = PersistenceManagerHelper.get();
		Transaction tx = pm.currentTransaction();
		try {
			tx.begin();
			T jdo = from(obj);
			verifyObject(jdo);
			pm.makePersistent(jdo);
			tx.commit();
			return true;
		} catch (Exception e) {
			throw e;
		} finally {
			if (tx.isActive()) {
				tx.rollback();
			}
			pm.close();
		}
	}
	/**
	 * 
	 * @param obj
	 * @return
	 * @throws Exception
	 */
	public boolean delete(E obj) throws Exception {
		verifyObject(obj);
		return delete(((Mirrorable) obj).getId());
	}
	/**
	 * 
	 * @param id
	 * @return
	 * @throws Exception
	 */
	public boolean delete(String id) throws Exception {
		if (!checkExistence(id)) {
			return false;
		}
		PersistenceManager pm = PersistenceManagerHelper.get();
		Transaction tx = pm.currentTransaction();
		try {
			tx.begin();
			T obj = pm.getObjectById(clazzT, id);
			pm.deletePersistent(obj);
			tx.commit();
			return true;
		} catch (Exception e) {
			throw e;
		} finally {
			if (tx.isActive()) {
				tx.rollback();
			}
			pm.close();
		}
	}
	/**
	 * 
	 * @param obj
	 * @return
	 * @throws Exception
	 */
	public boolean update(E obj) throws Exception {
		verifyObject(obj);
		PersistenceManager pm = PersistenceManagerHelper.get();
		Transaction tx = pm.currentTransaction();
		try {
			tx.begin();
			T jdo = from(obj);
			verifyObject(jdo);
			pm.makePersistent(jdo);
			tx.commit();
			return true;
		} catch (Exception e) {
			throw e;
		} finally {
			if (tx.isActive()) {
				tx.rollback();
			}
			pm.close();
		}
	}
	/**
	 * 
	 * @param id
	 * @return
	 * @throws Exception
	 */
	public E getById(String id) throws Exception {
		T tmp = getJDOById(id);		
		return to(tmp);
	}
	
	public List<E> listAll() throws Exception {
		PersistenceManager pm = PersistenceManagerHelper.get();
		Transaction tx = pm.currentTransaction();
		List<E> list = new ArrayList<E>();
		Query q = pm.newQuery(clazzT);
		try {
			tx.begin();
			List<T> tmp = (List<T>) q.execute();
			Iterator<T> iter = tmp.iterator();
			while (iter.hasNext()) {
				list.add(to(iter.next()));				
			}
			tx.commit();
			return list;
		} catch (Exception e) {
			throw e;
		} finally {
			if (tx.isActive()) {
				tx.rollback();
			}
			q.closeAll();
			pm.close();
		}
	}
	/**
	 * 
	 * @return
	 * @throws Exception
	 */
	public boolean deleteAll() throws Exception {
		PersistenceManager pm = PersistenceManagerHelper.get();
		Transaction tx = pm.currentTransaction();
		Query query = pm.newQuery(clazzT);
		if (query == null)
			return false;
		try {
			tx.begin();
			Object obj = query.execute();
			if (obj != null) {
				List list = (List) obj;
				if (list != null && list.size() > 0) {
					pm.deletePersistentAll(list);
				}
			}
			tx.commit();
			return true;
		} catch (Exception e) {
			throw e;
		} finally {
			if (tx.isActive()) {
				tx.rollback();
			}
			query.closeAll();
			pm.close();
		}
	}
	/**
	 * 
	 * @param id
	 * @return
	 * @throws Exception
	 */
	private T getJDOById(String id) throws Exception {
		PersistenceManager pm = PersistenceManagerHelper.get();
		Transaction tx = pm.currentTransaction();
		try {
			tx.begin();
			try {
				T tmp = pm.getObjectById(clazzT, id);
				if (tmp != null) {
					verifyObject(tmp);
					return tmp;
				}
			} catch (JDOObjectNotFoundException jex) {
			}
			tx.commit();
		} catch (Exception e) {
			throw e;
		} finally {
			if (tx.isActive()) {
				tx.rollback();
			}
			pm.close();
		}
		return null;
	}
	/**
	 * 
	 * @param id
	 * @return
	 * @throws Exception
	 */
	public boolean checkExistence(String id) throws Exception {
		T obj = getJDOById(id);
		return obj != null;
	}
	
	/**
	 * (non-Javadoc)
	 * @see com.cmg.hipspot.data.dao.InDataAccess#list(java.lang.String) 
	 */
	@Override
	public List<E> list(String query, Object... parameters) throws Exception {
		PersistenceManager pm = PersistenceManagerHelper.get();
		Transaction tx = pm.currentTransaction();
		List<E> list = new ArrayList<E>();
		Query q = pm.newQuery("SELECT FROM " + clazzT.getCanonicalName() + " " + query);
		try {
			tx.begin();
			List<T> tmp = (List<T>) q.execute(parameters);
			Iterator<T> iter = tmp.iterator();
			while (iter.hasNext()) {
				list.add(to(iter.next()));				
			}
			tx.commit();
			return list;
		} catch (Exception e) {
			throw e;
		} finally {
			if (tx.isActive()) {
				tx.rollback();
			}
			q.closeAll();
			pm.close();
		}
	}
	
	/**
	 * (non-Javadoc)
	 * @see com.cmg.hipspot.data.dao.InDataAccess#list(java.lang.String) 
	 */
	@Override
	public List<E> list(String query, Object parameter) throws Exception {
		PersistenceManager pm = PersistenceManagerHelper.get();
		Transaction tx = pm.currentTransaction();
		List<E> list = new ArrayList<E>();
		Query q = pm.newQuery("SELECT FROM " + clazzT.getCanonicalName() + " " + query);
		try {
			tx.begin();
			List<T> tmp = (List<T>) q.execute(parameter);
			Iterator<T> iter = tmp.iterator();
			while (iter.hasNext()) {
				list.add(to(iter.next()));				
			}
			tx.commit();
			return list;
		} catch (Exception e) {
			throw e;
		} finally {
			if (tx.isActive()) {
				tx.rollback();
			}
			q.closeAll();
			pm.close();
		}
	}
	
	/**
	 * (non-Javadoc)
	 * @see com.cmg.hipspot.data.dao.InDataAccess#list(java.lang.String) 
	 */
	@Override
	public List<E> list(String query, Object para1, Object para2) throws Exception {
		PersistenceManager pm = PersistenceManagerHelper.get();
		Transaction tx = pm.currentTransaction();
		List<E> list = new ArrayList<E>();
		Query q = pm.newQuery("SELECT FROM " + clazzT.getCanonicalName() + " " + query);
		try {
			tx.begin();
			List<T> tmp = (List<T>) q.execute(para1, para2);
			Iterator<T> iter = tmp.iterator();
			while (iter.hasNext()) {
				list.add(to(iter.next()));				
			}
			tx.commit();
			return list;
		} catch (Exception e) {
			throw e;
		} finally {
			if (tx.isActive()) {
				tx.rollback();
			}
			q.closeAll();
			pm.close();
		}
	}
	
	/**
	 * (non-Javadoc)
	 * @see com.cmg.hipspot.data.dao.InDataAccess#list(java.lang.String) 
	 */
	@Override
	public List<E> list(String query, Object para1, Object para2, Object para3) throws Exception {
		PersistenceManager pm = PersistenceManagerHelper.get();
		Transaction tx = pm.currentTransaction();
		List<E> list = new ArrayList<E>();
		Query q = pm.newQuery("SELECT FROM " + clazzT.getCanonicalName() + " " + query);
		try {
			tx.begin();
			List<T> tmp = (List<T>) q.execute(para1, para2, para3);
			Iterator<T> iter = tmp.iterator();
			while (iter.hasNext()) {
				list.add(to(iter.next()));				
			}
			tx.commit();
			return list;
		} catch (Exception e) {
			throw e;
		} finally {
			if (tx.isActive()) {
				tx.rollback();
			}
			q.closeAll();
			pm.close();
		}
	}
	
	/**
	 * (non-Javadoc)
	 * @see com.cmg.hipspot.data.dao.InDataAccess#list(java.lang.String) 
	 */
	@Override
	public List<E> list(String query) throws Exception {
		PersistenceManager pm = PersistenceManagerHelper.get();
		Transaction tx = pm.currentTransaction();
		List<E> list = new ArrayList<E>();
		Query q = pm.newQuery("SELECT FROM " + clazzT.getCanonicalName() + " " + query);
		try {
			tx.begin();
			List<T> tmp = (List<T>) q.execute();
			Iterator<T> iter = tmp.iterator();
			while (iter.hasNext()) {
				list.add(to(iter.next()));				
			}
			tx.commit();
			return list;
		} catch (Exception e) {
			throw e;
		} finally {
			if (tx.isActive()) {
				tx.rollback();
			}
			q.closeAll();
			pm.close();
		}
	}
}
