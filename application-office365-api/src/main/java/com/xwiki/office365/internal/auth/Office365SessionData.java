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
package com.xwiki.office365.internal.auth;

import java.io.Serializable;

/**
 * Office365 authentication state stored in the HTTP session of a user. It holds the serialized MSAL token cache of the
 * user, so that tokens are never shared between sessions, the number of failed authentication attempts and the reason
 * of the last failure.
 *
 * @version $Id$
 * @since 1.14.0
 */
public class Office365SessionData implements Serializable
{
    private static final long serialVersionUID = 1L;

    private String tokenCache;

    private int failedAttempts;

    private String error;

    /**
     * @return the serialized MSAL token cache, or {@code null} if not authenticated
     */
    public String getTokenCache()
    {
        return this.tokenCache;
    }

    /**
     * @param tokenCache the serialized MSAL token cache
     */
    public void setTokenCache(String tokenCache)
    {
        this.tokenCache = tokenCache;
    }

    /**
     * @return the number of failed authentication attempts since the last successful authentication
     */
    public int getFailedAttempts()
    {
        return this.failedAttempts;
    }

    /**
     * @param failedAttempts the number of failed authentication attempts
     */
    public void setFailedAttempts(int failedAttempts)
    {
        this.failedAttempts = failedAttempts;
    }

    /**
     * @return the reason of the last failed authentication attempt, or {@code null}
     */
    public String getError()
    {
        return this.error;
    }

    /**
     * @param error the reason of the last failed authentication attempt
     */
    public void setError(String error)
    {
        this.error = error;
    }

    /**
     * Forget the authentication, reset the failed attempts counter and the last error.
     */
    public void clear()
    {
        this.tokenCache = null;
        this.failedAttempts = 0;
        this.error = null;
    }
}
