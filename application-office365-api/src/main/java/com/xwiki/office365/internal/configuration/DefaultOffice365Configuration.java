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
package com.xwiki.office365.internal.configuration;

import java.util.LinkedHashMap;
import java.util.Map;

import javax.inject.Inject;
import javax.inject.Named;
import javax.inject.Singleton;

import org.apache.commons.lang3.StringUtils;
import org.xwiki.component.annotation.Component;
import org.xwiki.configuration.ConfigurationSource;

import com.xwiki.office365.configuration.Office365Configuration;

/**
 * Reads the configuration from the {@code Office365.Office365Config} page of the current wiki, falling back on the
 * main wiki when a value is not set. The tenant, client id and secret can also be set with the
 * {@code office365.tenant}, {@code office365.clientid} and {@code office365.secret} properties of
 * {@code xwiki.properties}, used as a last fallback.
 *
 * @version $Id$
 * @since 1.14.0
 */
@Component
@Singleton
public class DefaultOffice365Configuration implements Office365Configuration
{
    private static final String PROPERTIES_PREFIX = "office365.";

    @Inject
    @Named("office365")
    private ConfigurationSource wikiConfiguration;

    @Inject
    @Named("xwikiproperties")
    private ConfigurationSource xwikiProperties;

    @Override
    public String getTenant()
    {
        return getValue("tenant");
    }

    @Override
    public String getClientId()
    {
        return getValue("clientid");
    }

    @Override
    public String getSecret()
    {
        return getValue("secret");
    }

    @Override
    public Map<String, String> getSites()
    {
        Map<String, String> sites = new LinkedHashMap<>();
        for (String line : StringUtils.split(this.wikiConfiguration.getProperty("sites", ""), '\n')) {
            String[] values = line.split("=");
            if (values.length == 2) {
                sites.put(values[0].trim(), values[1].trim());
            }
        }
        return sites;
    }

    @Override
    public boolean isSearchExtensionActive()
    {
        return this.wikiConfiguration.getProperty("activateSearchExtension", false);
    }

    private String getValue(String name)
    {
        return this.wikiConfiguration.getProperty(name,
            this.xwikiProperties.getProperty(PROPERTIES_PREFIX + name, "")).trim();
    }
}
