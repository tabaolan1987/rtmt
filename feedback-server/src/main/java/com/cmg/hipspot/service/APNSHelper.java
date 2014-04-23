/**
 * Copyright (c) CMG Ltd All rights reserved.
 *
 * This software is the confidential and proprietary information of CMG
 * ("Confidential Information"). You shall not disclose such Confidential
 * Information and shall use it only in accordance with the terms of the
 * license agreement you entered into with CMG.
 */

package com.cmg.hipspot.service;

import java.util.List;

import org.apache.log4j.Logger;

import com.cmg.hipspot.properties.Configuration;

import javapns.Push;
import javapns.communication.exceptions.CommunicationException;
import javapns.communication.exceptions.KeystoreException;

/** 
 * DOCME
 * 
 * @Creator Hai Lu
 * @author $Author$
 * @version $Revision$
 * @Last changed: $LastChangedDate$
 */

public class APNSHelper {
	private static final Logger log = Logger.getLogger(APNSHelper.class);
	public static void postMessage(String message, String token) {
		try {
			
			Push.alert(message, Configuration.getValue(Configuration.ASPN_KEYSTORE), Configuration.getValue(Configuration.ASPN_KEY_PASSWD), false, token);
			log.info("Post message: " + message + " to device token: " + token);
		} catch (CommunicationException e) {
			log.error("Error when post message to iOS Devices", e);
		} catch (KeystoreException e) {
			log.error("Error when post message to iOS Devices", e);
		}
	}
	
	public static void postMessage(String message, List<String> tokens) {
		try {
			Push.alert(message, Configuration.getValue(Configuration.ASPN_KEYSTORE), Configuration.getValue(Configuration.ASPN_KEY_PASSWD), false, tokens);
			log.info("Post message: " + message + " to " + tokens.size() + " devices");
		} catch (CommunicationException e) {
			log.error("Error when post message to iOS Devices", e);
		} catch (KeystoreException e) {
			log.error("Error when post message to iOS Devices", e);
		}
	}
}
