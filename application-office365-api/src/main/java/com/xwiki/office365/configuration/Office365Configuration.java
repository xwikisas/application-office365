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

import java.util.Map;

import org.xwiki.component.annotation.Role;
import org.xwiki.stability.Unstable;

/**
 * Configuration of the Office365 integration, stored in the {@code Office365.Office365Config} page. Values defined in
 * the current wiki take precedence; empty values fall back on the configuration of the main wiki.
 *
 * @version $Id$
 * @since 1.14.0
 */
@Role
@Unstable
public interface Office365Configuration
{
    /**
     * @return the Azure AD tenant (directory) identifier or domain
     */
    String getTenant();

    /**
     * @return the client (application) identifier of the Azure AD application
     */
    String getClientId();

    /**
     * @return the client secret of the Azure AD application
     */
    String getSecret();

    /**
     * @return the configured SharePoint sites, mapping a display name to the Graph site identifier
     */
    Map<String, String> getSites();

    /**
     * @return {@code true} if the Office365 results should be displayed in the XWiki search
     */
    boolean isSearchExtensionActive();
}
