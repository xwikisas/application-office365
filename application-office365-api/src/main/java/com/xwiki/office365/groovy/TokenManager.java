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

import javax.servlet.http.HttpSession;

import com.microsoft.aad.msal4j.IAuthenticationResult;

/**
 * Manages authentication state and token storage in session. Handles authentication result caching and failure
 * tracking.
 *
 * @version $Id$
 * @since 1.14.0
 */
public class TokenManager
{
    private static final String AZURE_DATA_FIELD = "azuredata";

    private static final int MAX_FAILURES = 3;

    /**
     * Check if user is authenticated.
     *
     * @param session HTTP session
     * @return true if authenticated
     */
    public boolean isAuthenticated(HttpSession session)
    {
        AzureAuthData authData = getAuthData(session);
        return authData != null && authData.authResult != null && authData.nbfails <= MAX_FAILURES;
    }

    /**
     * Get current authentication result.
     *
     * @param session HTTP session
     * @return Authentication result or null
     */
    public IAuthenticationResult getAuthentication(HttpSession session)
    {
        AzureAuthData authData = getAuthData(session);
        if (authData == null || authData.authResult == null) {
            return null;
        }
        return authData.authResult;
    }

    /**
     * Store successful authentication result.
     *
     * @param session HTTP session
     * @param authResult Authentication result
     */
    public void storeAuthentication(HttpSession session, IAuthenticationResult authResult)
    {
        AzureAuthData authData = new AzureAuthData();
        authData.authResult = authResult;
        authData.nbfails = 0;
        authData.user = "";
        session.setAttribute(AZURE_DATA_FIELD, authData);
    }

    /**
     * Record a failed authentication attempt.
     *
     * @param session HTTP session
     */
    public void recordFailedAttempt(HttpSession session)
    {
        AzureAuthData currentAuthData = getAuthData(session);
        AzureAuthData authData = new AzureAuthData();
        authData.user = "";
        authData.authResult = (currentAuthData == null) ? null : currentAuthData.authResult;
        authData.nbfails = (currentAuthData == null) ? 1 : currentAuthData.nbfails + 1;
        session.setAttribute(AZURE_DATA_FIELD, authData);
    }

    /**
     * Clear authentication from session.
     *
     * @param session HTTP session
     */
    public void clearAuthentication(HttpSession session)
    {
        AzureAuthData authData = getAuthData(session);
        if (authData != null) {
            authData.authResult = null;
            authData.nbfails = 0;
            authData.user = "";
        }
    }

    /**
     * Check if too many failures have occurred.
     *
     * @param session HTTP session
     * @return true if max failures exceeded
     */
    public boolean hasTooManyFailures(HttpSession session)
    {
        AzureAuthData authData = getAuthData(session);
        return authData != null && authData.nbfails > MAX_FAILURES;
    }

    /**
     * Get authentication data from session.
     *
     * @param session HTTP session
     * @return Authentication data or null
     */
    private AzureAuthData getAuthData(HttpSession session)
    {
        return (AzureAuthData) session.getAttribute(AZURE_DATA_FIELD);
    }

    /**
     * Azure authentication data stored in session.
     */
    public static class AzureAuthData
    {
        /**
         * User identifier.
         */
        public String user = "";

        /**
         * Authentication result from MSAL.
         */
        public IAuthenticationResult authResult;

        /**
         * Number of failed authentication attempts.
         */
        public int nbfails;
    }
}
