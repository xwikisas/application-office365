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

package com.xwiki.office365.groovy;

import java.util.Map;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

/**
 * OAuth handler for Office 365 authentication flow. Simplifies integration by providing a high-level API for the
 * authentication process.
 * <p>
 * Usage example (similar to OAuth.xml):
 * <pre>
 * Office365OAuthHandler handler = new Office365OAuthHandler(configProvider);
 * handler.initialize(request, response, currentDocFullName);
 * if (handler.authenticate()) {
 *     // User is authenticated
 * } else {
 *     // Show authentication warning
 * }
 * System.out.println(handler.getDebugInfo(request));
 * </pre>
 *
 * @version $Id$
 * @since 1.14.0
 */
public class Office365OAuthHandler
{
    private AzureAuthClient authClient;

    /**
     * Initialize the handler with request and response objects.
     *
     * @param request HTTP request
     * @param response HTTP response
     * @param currentDocFullName Current document full name
     */
    public void initialize(HttpServletRequest request, HttpServletResponse response, String currentDocFullName)
    {
        authClient.initialize(request, response, currentDocFullName);
    }

    /**
     * Perform authentication.
     *
     * @return true if authentication is successful
     */
    public boolean authenticate()
    {
        return authClient.authenticate();
    }

    /**
     * Perform authentication with specific document name.
     *
     * @param docFullName Document full name
     * @return true if authentication is successful
     */
    public boolean authenticate(String docFullName)
    {
        return authClient.authenticate(docFullName);
    }

    /**
     * Get the authorization URL.
     *
     * @param docFullName Document full name
     * @return Authorization URL
     */
    public String getAuthorizationUrl(String docFullName)
    {
        return authClient.getAuthorizationUrl(docFullName);
    }

    /**
     * Check if current session has valid authentication.
     *
     * @param request HTTP request
     * @return true if authenticated
     */
    public boolean isAuthenticated(HttpServletRequest request)
    {
        return authClient.isAuthenticated(request.getSession());
    }

    /**
     * Get current authentication result.
     *
     * @param request HTTP request
     * @return Authentication result or null
     */
    public Object getAuthenticationResult(HttpServletRequest request)
    {
        return authClient.getAuthentication(request.getSession());
    }

    /**
     * Clear authentication from session.
     *
     * @param request HTTP request
     */
    public void clearAuthentication(HttpServletRequest request)
    {
        authClient.clearAuthentication(request.getSession());
    }

    /**
     * Get data from Graph API.
     *
     * @param url Graph API endpoint URL
     * @return Parsed JSON response as Map or null
     */
    public Map<String, Object> getGraphApiData(String url)
    {
        return authClient.getGraphApiData(url);
    }
}
