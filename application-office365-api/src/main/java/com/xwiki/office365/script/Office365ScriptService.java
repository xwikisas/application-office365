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
package com.xwiki.office365.script;

import java.util.Map;

import javax.inject.Inject;
import javax.inject.Named;
import javax.inject.Singleton;

import org.xwiki.component.annotation.Component;
import org.xwiki.script.service.ScriptService;
import org.xwiki.security.authorization.AccessDeniedException;
import org.xwiki.security.authorization.ContextualAuthorizationManager;
import org.xwiki.security.authorization.Right;
import org.xwiki.stability.Unstable;

import com.xwiki.office365.configuration.Office365Configuration;
import com.xwiki.office365.internal.auth.Office365Authenticator;
import com.xwiki.office365.internal.graph.GraphApiClient;

/**
 * Script service exposing the Office365 integration as {@code $services.office365}. The methods returning Microsoft
 * 365 data require the author of the calling script to have programming rights.
 *
 * @version $Id$
 * @since 1.14.0
 */
@Component
@Singleton
@Named("office365")
@Unstable
public class Office365ScriptService implements ScriptService
{
    @Inject
    private Office365Authenticator authenticator;

    @Inject
    private GraphApiClient graphApiClient;

    @Inject
    private Office365Configuration configuration;

    @Inject
    private ContextualAuthorizationManager authorization;

    /**
     * Run one step of the authentication flow for the current request, which may redirect the user to the Azure AD
     * login page.
     *
     * @return {@code true} if the current user is authenticated on Office365
     * @see Office365Authenticator#authenticate()
     */
    public boolean authenticate()
    {
        return this.authenticator.authenticate();
    }

    /**
     * @return the reason why the Office365 authentication of the current user failed (e.g. the error returned by Azure
     *     AD), or {@code null} if it didn't fail
     * @see Office365Authenticator#getError()
     */
    public String getAuthenticationError()
    {
        return this.authenticator.getError();
    }

    /**
     * @param url the Graph API URL, which must start with {@value GraphApiClient#GRAPH_URL}
     * @return the parsed JSON response, which holds an {@code error.message} entry when the request failed
     * @throws AccessDeniedException if the author of the calling script doesn't have programming rights
     * @see GraphApiClient#get(String)
     */
    public Map<String, Object> getGraphApiData(String url) throws AccessDeniedException
    {
        checkProgrammingRights();
        return this.graphApiClient.get(url);
    }

    /**
     * Search the personal drive of the user and all the configured SharePoint sites.
     *
     * @param query the search query
     * @return the parsed JSON response of each drive, mapped by site name (an empty string for the personal drive,
     *     which comes first); a response with an {@code error.message} entry means the drive could not be searched
     * @throws AccessDeniedException if the author of the calling script doesn't have programming rights
     */
    public Map<String, Map<String, Object>> searchSitesForText(String query) throws AccessDeniedException
    {
        checkProgrammingRights();
        return this.graphApiClient.searchSitesForText(query);
    }

    /**
     * @param site the name of a configured SharePoint site, or an empty string for the personal drive
     * @param itemId the identifier of the drive item
     * @return the parsed JSON response, which holds an {@code error.message} entry when the request failed
     * @throws AccessDeniedException if the author of the calling script doesn't have programming rights
     */
    public Map<String, Object> getDriveItem(String site, String itemId) throws AccessDeniedException
    {
        checkProgrammingRights();
        return this.graphApiClient.getDriveItem(site, itemId);
    }

    /**
     * @return {@code true} if the Office365 results should be displayed in the XWiki search
     */
    public boolean isSearchExtensionActive()
    {
        return this.configuration.isSearchExtensionActive();
    }

    /**
     * The Graph API data is read with the access token of the current user, so only the scripts written by users with
     * programming rights can read it. Otherwise any script author could read the Microsoft 365 data of the users
     * viewing their page.
     */
    private void checkProgrammingRights() throws AccessDeniedException
    {
        this.authorization.checkAccess(Right.PROGRAM);
    }
}
