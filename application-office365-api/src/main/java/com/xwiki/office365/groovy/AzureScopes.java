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

/**
 * Contains string constants for azure scopes.
 *
 * @version $Id$
 * @since 1.14.0
 */
public final class AzureScopes
{
    static final String FILES_READWRITE = "Files.ReadWrite";

    static final String USER_READ = "User.Read";

    static final String USER_READWRITE = "User.ReadWrite";

    static final String USER_READBASIC_ALL = "User.ReadBasic.All";

    static final String FILES_READWRITE_ALL = "Files.ReadWrite.All";

    static final String FILES_READWRITE_APPFOLDER = "Files.ReadWrite.AppFolder";

    static final String SITES_READWRITE_ALL = "Sites.ReadWrite.All";

    private AzureScopes()
    {
        /* This utility class should not be instantiated */
    }
}
