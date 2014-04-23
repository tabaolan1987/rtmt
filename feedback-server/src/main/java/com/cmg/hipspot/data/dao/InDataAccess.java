/**
 * Copyright (c) CMG Ltd All rights reserved.
 *
 * This software is the confidential and proprietary information of CMG
 * ("Confidential Information"). You shall not disclose such Confidential
 * Information and shall use it only in accordance with the terms of the
 * license agreement you entered into with CMG.
 */

package com.cmg.hipspot.data.dao;

import java.util.List;

/** 
 * DOCME
 * 
 * @Creator Hai Lu
 * @author $Author$
 * @version $Revision$
 * @Last changed: $LastChangedDate$
 */

public interface InDataAccess<T, E> {
	
	public boolean deleteAll() throws Exception;
	
	public boolean put(E obj) throws Exception;
	
	public boolean create(E obj) throws Exception;
	
	public boolean delete(E obj) throws Exception;
	
	public boolean delete(String id) throws Exception;
	
	public boolean update(E obj) throws Exception;
	
	public E getById(String id) throws Exception;
	
	public List<E> listAll() throws Exception;
	
	public List<E> list(String query, Object parameter) throws Exception;
	
	public List<E> list(String query, Object para1, Object para2) throws Exception;
	
	public List<E> list(String query, Object para1, Object para2, Object para3) throws Exception;
	
	public List<E> list(String query, Object... parameters) throws Exception;
	
	public List<E> list(String query) throws Exception;
	
	public boolean checkExistence(String id) throws Exception;
}
