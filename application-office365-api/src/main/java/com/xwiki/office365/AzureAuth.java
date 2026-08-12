/*
 * See the NOTICE file distributed with this work for additional
 * information regarding copyright ownership.
 *
 * This is free software; you can redistribute it and/or modify it
 * under the terms of the GNU Lesser General Public License as
 * published by the Free Software Foundation; either version 2.1 of
 * the License, or (at your option) any later version.
 *
 * This software is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU
 * Lesser General Public License for more details.
 *
 * You should have received a copy of the GNU Lesser General Public
 * License along with this software; if not, write to the Free
 * Software Foundation, Inc., 51 Franklin St, Fifth Floor, Boston, MA
 * 02110-1301 USA, or see the FSF site: http://www.fsf.org.
 */
package com.xwiki.office365;

import java.net.MalformedURLException;

import javax.inject.Inject;
import javax.inject.Named;

import org.apache.commons.lang3.exception.ExceptionUtils;
import org.slf4j.Logger;

import com.microsoft.aad.msal4j.ClientCredentialFactory;
import com.microsoft.aad.msal4j.ConfidentialClientApplication;
import com.microsoft.aad.msal4j.IClientCredential;
import com.xwiki.office365.configuration.AzureConfiguration;
import com.xwiki.office365.configuration.IAzureConfiguration;

/**
 * Log in using msal. FIXME: Placeholder.
 *
 * @since 1.14.0
 * @version $Id$
 */
public class AzureAuth
{
    @Inject
    private Logger logger;

    @Inject
    @Named(AzureConfiguration.HINT)
    private IAzureConfiguration azureConfiguration;

    private ConfidentialClientApplication app;

    /**
     * Initialize the msal app if it doesn't exist, or get the existing instance.
     *
     * @return The ClientApplication if it is possible to get, null if an error happened.
     */
    public ConfidentialClientApplication getApp()
    {
        if (null != this.app) {
            return this.app;
        }
        try {
            IClientCredential credential = ClientCredentialFactory.createFromSecret(azureConfiguration.getSecret());
            this.app =
                ConfidentialClientApplication.builder(azureConfiguration.getClientID(), credential)
                .authority(azureConfiguration.getAuthority()).build();
            return this.app;
        } catch (MalformedURLException e) {
            logger.error("Error while creating MSAL Client: [{}]", ExceptionUtils.getRootCauseMessage(e));
        }
        return null;
    }
}
