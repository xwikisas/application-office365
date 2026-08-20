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

import java.net.URI;
import java.util.Collections;
import java.util.Map;

import org.apache.commons.lang3.exception.ExceptionUtils;
import org.slf4j.Logger;

import com.google.gson.Gson;
import com.microsoft.graph.serviceclient.GraphServiceClient;
import com.microsoft.kiota.authentication.AccessTokenProvider;
import com.microsoft.kiota.authentication.AllowedHostsValidator;
import com.microsoft.kiota.authentication.BaseBearerTokenAuthenticationProvider;

/**
 * Client for interacting with Microsoft Graph API using the Graph SDK. Provides a clean, object-oriented interface for
 * Graph API calls.
 *
 * @version $Id$
 * @since 1.14.0
 */
public class GraphApiClient
{
    protected Logger logger;

    private GraphServiceClient graphServiceClient;

    /**
     * Constructor.
     */
    public GraphApiClient()
    {
    }

    /**
     * Initialize the Graph client with an access token.
     *
     * @param accessToken Access token from MSAL authentication
     */
    public void initialize(String accessToken)
    {
        AccessTokenProvider tokenProvider = new AccessTokenProvider()
        {
            @Override
            public String getAuthorizationToken(URI uri, Map<String, Object> additionalAuthenticationContext)
            {
                return accessToken;
            }

            @Override
            public AllowedHostsValidator getAllowedHostsValidator()
            {
                return new AllowedHostsValidator("graph.microsoft.com");
            }
        };
        this.graphServiceClient = new GraphServiceClient(new BaseBearerTokenAuthenticationProvider(tokenProvider));
    }

    /**
     * Get data from Graph API endpoint.
     *
     * @param url Full Graph API URL
     * @return Response as Map or empty map on error
     */
    public Map<String, Object> getGraphData(String url)
    {
        if (graphServiceClient == null) {
            logger.warn("Graph client not initialized");
            return Collections.emptyMap();
        }

        try {
            java.net.URI uri = new java.net.URI(url);
            String resourcePath = extractResourcePath(uri.getPath());

            return makeGraphRequest(resourcePath);
        } catch (Exception e) {
            logger.error("Exception executing Graph API query: [{}]", ExceptionUtils.getRootCauseMessage(e));
            return Collections.emptyMap();
        }
    }

    /**
     * Make a Graph API request using the SDK's fluent API.
     *
     * @param resourcePath Resource path (e.g., "me", "users", "drive/root/children")
     * @return Response as Map or empty map on error
     */
    private Map<String, Object> makeGraphRequest(String resourcePath)
    {
        try {
            Object result = fetchResource(resourcePath);
            return convertToMap(result);
        } catch (Exception e) {
            if (isInvalidTokenError(e)) {
                logger.warn("Invalid authentication token detected");
            }
            logger.error("Exception making Graph request for [{}]: [{}]", resourcePath,
                ExceptionUtils.getRootCauseMessage(e));
            return Collections.emptyMap();
        }
    }

    /**
     * Fetch a resource from Graph API using the appropriate SDK method.
     *
     * @param resourcePath Resource path
     * @return Response object from Graph SDK
     */
    private Object fetchResource(String resourcePath)
    {
        String[] segments = resourcePath.split("/");

        if ("me".equalsIgnoreCase(segments[0])) {
            return graphServiceClient.me().get();
        } else if ("users".equalsIgnoreCase(segments[0]) && segments.length > 1) {
            return graphServiceClient.users().byUserId(segments[1]).get();
        } else if ("drive".equalsIgnoreCase(segments[0])) {
            return graphServiceClient.me().drive().get();
        } else if (resourcePath.contains("sites")) {
            return graphServiceClient.sites().get();
        } else {
            logger.warn("Unsupported Graph API resource path: [{}]", resourcePath);
            return null;
        }
    }

    /**
     * Extract the resource path from a full URL path.
     *
     * @param urlPath Full URL path from URI
     * @return Clean resource path
     */
    private String extractResourcePath(String urlPath)
    {
        return urlPath.replace("/v1.0/", "").replace("/beta/", "");
    }

    /**
     * Check if an exception indicates an invalid token.
     *
     * @param e Exception to check
     * @return true if token is invalid
     */
    private boolean isInvalidTokenError(Exception e)
    {
        String message = e.getMessage();
        return message != null && message.contains("InvalidAuthenticationToken");
    }

    /**
     * Convert a Graph SDK response object to a Map.
     *
     * @param obj Object from Graph API response
     * @return Map representation or empty map
     */
    private Map<String, Object> convertToMap(Object obj)
    {
        if (obj == null) {
            return Collections.emptyMap();
        }

        Gson gson = new Gson();
        try {
            String json = gson.toJson(obj);
            return gson.fromJson(json, Map.class);
        } catch (Exception e) {
            logger.warn("Could not convert response to Map: [{}]", ExceptionUtils.getRootCauseMessage(e));
            return Collections.emptyMap();
        }
    }
}
