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

import java.io.IOException;
import java.net.URLEncoder;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;
import java.util.TreeMap;

import javax.inject.Inject;
import javax.inject.Named;
import javax.inject.Singleton;

import org.apache.velocity.tools.generic.EscapeTool;
import org.xwiki.component.annotation.Component;
import org.xwiki.model.reference.DocumentReference;

import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.xwiki.azureoauth.AzureADIdentityOAuthProvider;
import com.xwiki.identityoauth.IdentityOAuthException;
import com.xwiki.microsoft365.Microsoft365Connections;

/**
 * Implementation of the methods to talk to the cloud of Microsoft 365.
 *
 * @version $Id$
 * @since 1.0
 */
@Component
@Singleton
@Named("microsoft365")
public class Microsoft365ConnectionsImpl extends AzureADIdentityOAuthProvider
        implements Microsoft365Connections, Microsoft365Constants
{
    private DocumentReference ms365WebPreferences;

    @Inject
    private MacroRunner macroRunner;

    private Map<String, String> sites = new TreeMap<>();

    @Override
    public void initialize(Map<String, String> config)
    {
        // let objects regarding simple AzureAD§ config be read
        super.initialize(config);
        // grab the office and sharepoint specific information
        String sitesConfig = config.get(SITES);
        if (sitesConfig == null) {
            sitesConfig = "";
        }
        sites = parseSites(sitesConfig);
        ms365WebPreferences = documentResolver.resolve("xwiki:Microsoft365.WebPreferences");
    }

    @Override
    public boolean isActive()
    {
        return true;
        // TODO licensorProvider.get().hasLicensure(ms365WebPreferences);
    }

    /**
     * Parses equality pairs.
     *
     * @param sitesLinesConfig a string made of lines key = value
     * @return a map describing these equalities.
     */
    Map<String, String> parseSites(String sitesLinesConfig)
    {
        String[] sitesLines = sitesLinesConfig.split("\\n|\\r");
        Map<String, String> parsed = new TreeMap<>();
        for (String siteLine : sitesLines) {
            if (siteLine == null || siteLine.trim().length() == 0) {
                continue;
            }
            String[] pair = siteLine.split("=|\\s");
            parsed.put(pair[0], pair[1]);
        }
        return parsed;
    }

    /**
     * Checks for existence of a token as a followup of a login-with-microsoft.
     *
     * @return true if a token for running APIs is not available.
     */
    @Override
    public boolean isMissingAuth()
    {
        return identityOAuthManager == null || !identityOAuthManager.get().hasSessionIdentityInfo(PROVIDERNAME_MS365);
    }

    /**
     * @return the URL to request to authorize.
     */
    public String getOAuthStartUrl()
    {
        return identityOAuthManager.get().getOAuthStartUrl(this);
    }

    /**
     * Returns the constant value (needed as otherwise AzureAD is returned and the wrong provider is given the tokens).
     *
     * @return "microsoft365"
     */
    @Override
    public String getProviderHint()
    {
        return PROVIDERNAME_MS365;
    }

    /**
     * @param hint expected to be exactly "microsoft365"
     */
    @Override
    public void setProviderHint(String hint)
    {
        if (!PROVIDERNAME_MS365.equals(hint)) {
            throw new IllegalStateException("Only \"" + PROVIDERNAME_MS365 + "\" is accepted as hint.");
        }
    }

    private void debugMsg(String msg)
    {
        List<String> dbgMsgs = (List<String>) contextProvider.get().get(DEBUG_MESSAGES);
        if (dbgMsgs == null) {
            dbgMsgs = new LinkedList<String>();
            contextProvider.get().put(DEBUG_MESSAGES, dbgMsgs);
        }
        logger.debug(msg);
        dbgMsgs.add(msg);
    }

    @Override
    public String getDebugInfo()
    {
        StringBuilder b = new StringBuilder();
        b.append("Produced by ").append(this.toString()).append(" ");
        List<String> debugMessages = (List<String>) contextProvider.get().get(DEBUG_MESSAGES);
        if (debugMessages != null && debugMessages.size() > 0) {
            for (String msg : debugMessages) {
                b.append("\n<br>").append(msg);
            }
        }
        return b.toString();
    }

    private String readParam(String paramName)
    {
        return contextProvider.get().getRequest().get(paramName);
    }

    @Override
    public SearchResult searchDocuments()
    {
        try {
            String site = readParam(SITE);
            if (site != null && sites.containsKey(site)) {
                throw new RuntimeException("No such site " + QUOTE + site + QUOTE);
            }

            String baseURL = getDriveURL(site);
            EscapeTool escapeTool = new EscapeTool();
            String text = readParam(SEARCH_TEXT);
            if (text == null || text.length() == 0) {
                return new SearchResult(Collections.emptyList(), "", null);
            }
            if (text.contains(APOSTROPHE)) {
                text = text.replaceAll(APOSTROPHE, QUOTE);
            }
            debugMsg("Searching for " + QUOTE + text + QUOTE);

            Map searchResultMap = super.makeApiCall(baseURL + "search(q='" + escapeTool.url(text) + "')");
            List results = (List) searchResultMap.get("value");

            throwOnPossibleError((Map) searchResultMap.get("error"));
            List<SearchResultItem> searchResults = new ArrayList<>(results.size());
            for (Object resO : results) {
                final Map res = (Map) resO;
                SearchResultItem r = new SearchResultItemObject(res);
                searchResults.add(r);
            }
            return new SearchResult(searchResults, text, null);
        } catch (Exception e) {
            logger.warn("Trouble at search document.", e);
            debugMsg(e.toString());
            SearchResult sr = new SearchResult(null, null, e);
            return sr;
        }
    }

    private void throwOnPossibleError(Map errorMap)
    {
        if (errorMap != null) {
            Exception details = null;
            try {
                details = new Exception(new ObjectMapper().writeValueAsString(errorMap));
            } catch (JsonProcessingException e) {
                e.printStackTrace();
            }
            String message = "Connection to Microsoft365 failed: "
                    + errorMap.get("message");
            IdentityOAuthException ex = new IdentityOAuthException(message, details);
            ex.printStackTrace();
            throw ex;
        }
    }

    @Override
    public MacroRun runMacro(Object macroObject)
    {
        Map<String, Object> map = (Map<String, Object>) macroObject;
        return macroRunner.runMacro(map);
    }

    @Override
    public List<String> getSites()
    {
        return Arrays.asList((String[]) sites.keySet().toArray());
    }

    String getDriveURL(String site)
    {
        String baseURL = "https://graph.microsoft.com/v1.0/me/drive/";
        if (site != null && site.length() > 0) {
            String siteId = sites.get(site);
            baseURL = "https://graph.microsoft.com/v1.0/sites/${siteId}/drive/";
        }
        return baseURL;
    }

    Map getGraphApiInfo(String site, String docId) throws IOException
    {
        return super.makeApiCall(getDriveURL(site) + "items/" + URLEncoder.encode(docId, "UTF-8"));
    }

    String requestEmbedUrl(String docId, String site)
    {
        Map json = super.makeApiCall(getDriveURL(site) + docId + PREVIEW);
        return (String) json.get(GETURL);
    }
}
