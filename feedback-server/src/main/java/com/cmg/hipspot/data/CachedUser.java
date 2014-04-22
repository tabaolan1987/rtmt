/**
 * Copyright (c) CMG Ltd All rights reserved.
 *
 * This software is the confidential and proprietary information of CMG
 * ("Confidential Information"). You shall not disclose such Confidential
 * Information and shall use it only in accordance with the terms of the
 * license agreement you entered into with CMG.
 */

package com.cmg.hipspot.data;

import java.util.ArrayList;
import java.util.List;

/** 
 * DOCME
 * 
 * @Creator Hai Lu
 * @author $Author$
 * @version $Revision$
 * @Last changed: $LastChangedDate$
 */

public class CachedUser {
	private String fid;
	private List<String> regIds;
	
	public CachedUser(String fbid, String regid) {
		regIds = new ArrayList<String>();
		fid = fbid;
		regIds.add(regid);
	}
	/** 
	 * @return the fid 
	 */
	public String getFid() {
		return fid;
	}
	/** 
	 * @param fid the fid to set 
	 */
	
	public void setFid(String fid) {
		this.fid = fid;
	}
	/** 
	 * @return the regIds 
	 */
	public List<String> getRegIds() {
		return regIds;
	}
	/** 
	 * @param regIds the regIds to set 
	 */
	
	public void setRegIds(List<String> regIds) {
		this.regIds = regIds;
	}
	
	public boolean addRegId(String regId) {
		if (regIds == null) {
			regIds = new ArrayList<String>();
		}
		boolean isExist = false;
		for (String s: regIds) {
			if (s.equalsIgnoreCase(regId)) {
				isExist = true;
				break;
			}
		}
		if (!isExist) {
			regIds.add(regId);
			return true;
		}
		return false;
	}
	
	public boolean removeRegId(String regId) {
		if (regIds == null) {
			regIds = new ArrayList<String>();
		}
		int index = -1;
		int i;
		for (i = 0; i < regIds.size() - 1; i ++) {
			if (regIds.get(i).equalsIgnoreCase(regId)) {
				index = i;
				break;
			}
		}
		if (index != -1) {
			regIds.remove(index);
			return true;
		}
		return false;
	}
	
	public boolean updateRegId(String oldId, String newId) {
		if (regIds == null) {
			regIds = new ArrayList<String>();
		}
		int index = -1;
		int i;
		for (i = 0; i < regIds.size() - 1; i ++) {
			if (regIds.get(i).equalsIgnoreCase(oldId)) {
				index = i;
				break;
			}
		}
		regIds.add(newId);
		if (index != -1) {
			regIds.remove(index);
			
			return true;
		}
		
		return false;
	}
	
	
}
