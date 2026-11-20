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
package com.xwiki.office365.internal.graph;

import java.net.URI;
import java.net.URLEncoder;
import java.net.http.HttpClient;
import java.net.http.HttpRequest;
import java.net.http.HttpResponse;
import java.nio.charset.StandardCharsets;
import java.time.Duration;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.concurrent.CompletableFuture;
import java.util.stream.Collectors;

import javax.inject.Inject;
import javax.inject.Singleton;

import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.exception.ExceptionUtils;
import org.slf4j.Logger;
import org.xwiki.component.annotation.Component;

import com.fasterxml.jackson.core.type.TypeReference;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.xwiki.office365.configuration.Office365Configuration;
import com.xwiki.office365.internal.auth.Office365Authenticator;

/**
 * Client for the Microsoft Graph API, using the JDK HTTP client and the access token of the current user.
 *
 * @version $Id$
 * @since 1.14.0
 */
@Component(roles = GraphApiClient.class)
@Singleton
public class GraphApiClient
{
    /**
     * The base URL of the Microsoft Graph API. Only URLs starting with it are accepted.
     */
    public static final String GRAPH_URL = "https://graph.microsoft.com/";

    private static final String API_URL = GRAPH_URL + "v1.0/";

    private static final String DRIVE = "drive/";

    private static final String ERROR = "error";

    private static final Duration TIMEOUT = Duration.ofSeconds(30);

    private static final TypeReference<Map<String, Object>> MAP_TYPE = new TypeReference<Map<String, Object>>()
    {
    };

    private final HttpClient httpClient = HttpClient.newBuilder().connectTimeout(TIMEOUT).build();

    private final ObjectMapper objectMapper = new ObjectMapper();

    @Inject
    private Office365Authenticator authenticator;

    @Inject
    private Office365Configuration configuration;

    @Inject
    private Logger logger;

    /**
     * Run a GET request on the Graph API. If the access token is rejected, the authentication is cleared and a new
     * authentication is started.
     *
     * @param url the full Graph API URL, starting with {@link #GRAPH_URL}
     * @return the parsed JSON response; when the request failed or could not be made (e.g. the user is not
     *     authenticated), it contains an {@code error} entry whose {@code message} holds the reason, as returned by
     *     the Graph API
     */
    public Map<String, Object> get(String url)
    {
        return getAll(Collections.singletonList(url)).get(0);
    }

    /**
     * Search, in parallel, the personal drive of the user and the drives of all the configured SharePoint sites.
     *
     * @param query the search query
     * @return the parsed JSON response of each drive, see {@link #get(String)}, mapped by site name (an empty string
     *     for the personal drive, which comes first)
     */
    public Map<String, Map<String, Object>> searchSitesForText(String query)
    {
        Set<String> sites = new LinkedHashSet<>();
        sites.add("");
        sites.addAll(this.configuration.getSites().keySet());
        List<String> urls = sites.stream().map(site -> getSearchURL(site, query)).collect(Collectors.toList());
        Iterator<Map<String, Object>> responses = getAll(urls).iterator();
        Map<String, Map<String, Object>> results = new LinkedHashMap<>();
        for (String site : sites) {
            results.put(site, responses.next());
        }
        return results;
    }

    /**
     * Get the metadata of a drive item, including its {@code @microsoft.graph.downloadUrl}.
     *
     * @param site the name of a configured SharePoint site, or an empty string for the personal drive of the user
     * @param itemId the identifier of the drive item
     * @return the parsed JSON response, see {@link #get(String)}
     */
    public Map<String, Object> getDriveItem(String site, String itemId)
    {
        String driveURL = getDriveURL(site);
        if (driveURL == null || StringUtils.isEmpty(itemId)) {
            return error("Unknown SharePoint site or document");
        }
        return get(driveURL + "items/" + encode(itemId));
    }

    /**
     * Run the GET requests in parallel, using the same access token.
     *
     * @param urls the Graph API URLs; a {@code null} URL leads to an error response
     * @return the parsed JSON responses, in the same order as the URLs
     */
    private List<Map<String, Object>> getAll(List<String> urls)
    {
        String accessToken = getAccessToken();
        if (accessToken == null) {
            String reason = StringUtils.defaultIfBlank(this.authenticator.getError(),
                "Not authenticated on Microsoft 365");
            return new ArrayList<>(Collections.nCopies(urls.size(), error(reason)));
        }

        List<CompletableFuture<Map<String, Object>>> requests = new ArrayList<>(urls.size());
        for (String url : urls) {
            requests.add(send(url, accessToken));
        }
        List<Map<String, Object>> results = new ArrayList<>(urls.size());
        boolean invalidToken = false;
        for (CompletableFuture<Map<String, Object>> request : requests) {
            Map<String, Object> data = request.join();
            invalidToken |= isInvalidTokenError(data);
            results.add(data);
        }
        if (invalidToken) {
            this.logger.debug("The access token was rejected. Authenticating again");
            restartAuthentication();
        }
        return results;
    }

    private String getAccessToken()
    {
        if (!this.authenticator.isAuthenticated()) {
            return null;
        }
        String accessToken = this.authenticator.getAccessToken();
        if (accessToken == null) {
            // The tokens can't be refreshed anymore, so the user needs to log in again.
            restartAuthentication();
        }
        return accessToken;
    }

    /**
     * Send an asynchronous GET request to the Graph API.
     *
     * @param url the Graph API URL; any other URL is refused, so that the access token is never leaked
     * @param accessToken the access token of the current user
     * @return the parsed JSON response, or an error response (see {@link #error(String)}) if the request failed
     */
    private CompletableFuture<Map<String, Object>> send(String url, String accessToken)
    {
        if (url == null) {
            return CompletableFuture.completedFuture(error("Unknown SharePoint site"));
        }
        // Never send the access token of the user anywhere else than to the Graph API.
        if (!url.startsWith(GRAPH_URL)) {
            this.logger.warn("Refusing to query a URL outside of the Graph API: [{}]", url);
            return CompletableFuture.completedFuture(error("Only Microsoft Graph API URLs are allowed"));
        }
        try {
            HttpRequest request = HttpRequest.newBuilder(URI.create(url)).timeout(TIMEOUT)
                .header("Authorization", "Bearer " + accessToken).header("Accept", "application/json").GET().build();
            this.logger.debug("HTTP running query: [{}]", url);
            return this.httpClient.sendAsync(request, HttpResponse.BodyHandlers.ofString())
                .handle((response, e) -> (e != null) ? failure(url, e) : read(url, response));
        } catch (IllegalArgumentException e) {
            return CompletableFuture.completedFuture(failure(url, e));
        }
    }

    /**
     * Parse the JSON body of a Graph API response, including the error responses returned by the Graph API.
     *
     * @param url the requested URL, for logging
     * @param response the HTTP response
     * @return the parsed JSON response, or an error response (see {@link #error(String)}) if the body is not JSON
     */
    private Map<String, Object> read(String url, HttpResponse<String> response)
    {
        try {
            Map<String, Object> data = this.objectMapper.readValue(response.body(), MAP_TYPE);
            if (data != null) {
                return data;
            }
        } catch (Exception e) {
            // Not a JSON response, so not an error returned by the Graph API itself (e.g. a proxy error page).
            this.logger.debug("Invalid response for [{}]: [{}]", url, ExceptionUtils.getRootCauseMessage(e));
        }
        return error("Unexpected response from Microsoft Graph (HTTP " + response.statusCode() + ")");
    }

    private Map<String, Object> failure(String url, Throwable e)
    {
        String reason = ExceptionUtils.getRootCauseMessage(e);
        this.logger.debug("Exception executing query [{}]: [{}]", url, reason);
        return error(reason);
    }

    /**
     * @param message the reason of the failure
     * @return an error response with the same structure as the ones returned by the Graph API, so that all the errors
     *     are handled the same way
     */
    private static Map<String, Object> error(String message)
    {
        return Map.of(ERROR, Map.of("message", message));
    }

    private String getDriveURL(String site)
    {
        if (StringUtils.isEmpty(site)) {
            return API_URL + "me/" + DRIVE;
        }
        String siteId = this.configuration.getSites().get(site);
        if (siteId == null) {
            this.logger.debug("Unknown SharePoint site: [{}]", site);
            return null;
        }
        return API_URL + "sites/" + siteId + '/' + DRIVE;
    }

    private String getSearchURL(String site, String query)
    {
        String driveURL = getDriveURL(site);
        if (driveURL == null) {
            return null;
        }
        // Single quotes are escaped by doubling them in OData string literals.
        return driveURL + "search(q='" + encode(StringUtils.defaultString(query).replace("'", "''")) + "')";
    }

    private void restartAuthentication()
    {
        this.authenticator.clearAuthentication();
        this.authenticator.authenticate();
    }

    private boolean isInvalidTokenError(Map<String, Object> data)
    {
        Object error = data.get(ERROR);
        return error instanceof Map && "InvalidAuthenticationToken".equals(((Map<?, ?>) error).get("code"));
    }

    private String encode(String value)
    {
        return URLEncoder.encode(value, StandardCharsets.UTF_8).replace("+", "%20");
    }
}
