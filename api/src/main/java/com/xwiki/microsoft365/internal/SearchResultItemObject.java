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
package com.xwiki.microsoft365.internal;

import java.util.Map;

import com.xwiki.microsoft365.Microsoft365Connections;

class SearchResultItemObject implements Microsoft365Connections.SearchResultItem, Microsoft365Constants
{
    private Map res;

    SearchResultItemObject(Map r)
    {
        this.res = r;
    }

    public String getName()
    {
        return (String) res.get(NAME);
    }

    public String getEmbedUrl()
    {
        return ((String) res.get(WEB_URL))
                .replaceAll("action=default", "action=embedview");
    }

    public String getViewUrl()
    {
        return ((String) res.get(WEB_URL));
    }

    public String getId()
    {
        return ((String) res.get("id"));
    }

    public String getVersion()
    {
        return ((String) res.get("version"));
    }

    public String getFilename()
    {
        return ((String) res.get(NAME));
    }
}
