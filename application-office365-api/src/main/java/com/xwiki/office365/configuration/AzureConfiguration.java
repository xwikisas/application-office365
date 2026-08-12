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
package com.xwiki.office365.configuration;

import javax.inject.Inject;
import javax.inject.Named;
import javax.inject.Singleton;

import org.xwiki.component.annotation.Component;
import org.xwiki.configuration.ConfigurationSource;
import org.xwiki.stability.Unstable;

/**
 * Old AzureAD configuration properties from Identity OAuth integration.
 *
 * @since 1.14.0
 * @version $Id$
 */
@Component
@Singleton
@Named(AzureConfiguration.HINT)
@Unstable
public class AzureConfiguration implements IAzureConfiguration
{
    /**
     * Component hint.
     */
    public static final String HINT = "OFFICE365_AZURE_CONFIGURATION";

    @Inject
    @Named(AzureConfigurationSource.HINT)
    private ConfigurationSource configurationSource;

    @Override
    public String getTenantID()
    {
        return configurationSource.getProperty("tenant");
    }

    @Override
    public String getClientID()
    {
        return configurationSource.getProperty("clientid");
    }

    @Override
    public String getSecret()
    {
        return configurationSource.getProperty("secret");
    }

    @Override
    public String getAuthority()
    {
        return "https://login.microsoftonline.com";
    }
}
