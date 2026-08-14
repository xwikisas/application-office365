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

import java.net.URLEncoder;
import java.nio.charset.StandardCharsets;
import java.util.Collections;
import java.util.HashMap;
import java.util.Map;
import java.util.Set;
import java.util.UUID;

import javax.inject.Inject;
import javax.inject.Named;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import javax.servlet.http.HttpSession;

import org.apache.commons.lang3.exception.ExceptionUtils;
import org.apache.http.HttpResponse;
import org.apache.http.client.methods.HttpGet;
import org.apache.http.impl.client.DefaultHttpClient;
import org.apache.http.util.EntityUtils;
import org.slf4j.Logger;

import com.google.gson.Gson;
import com.microsoft.aad.msal4j.ClientCredentialFactory;
import com.microsoft.aad.msal4j.ConfidentialClientApplication;
import com.microsoft.aad.msal4j.IAuthenticationResult;
import com.xwiki.office365.configuration.AzureConfiguration;
import com.xwiki.office365.configuration.IAzureConfiguration;

/**
 * Azure Active Directory authentication client for Office 365 integration. Handles OAuth2 authentication flow using
 * MSAL4J (Microsoft Authentication Library for Java).
 * <p>
 * MSAL4J provides: - Modern OAuth2/OpenID Connect support - Built-in token caching - Automatic token refresh - Better
 * error handling and diagnostics
 *
 * @version $Id$
 * @since 1.14.0
 */
public class AzureAuthClient
{
    private static final String AUTH_DOMAIN = "login.windows.net";

    private static final String API_VERSION = "1.0";

    private static final String AZURE_DATA_FIELD = "azuredata";

    private static final String AZURE_PARAM_ERROR = "error";

    private static final String AZURE_PARAM_ERROR_DESCRIPTION = "error_description";

    private static final String AZURE_PARAM_CODE = "code";

    private static final String AUTH_PAGE = "Office365.OAuth";

    private static final String AZURE_SESSION_REDIRECT = "azure_redirect";

    private static final String GRAPH_API_URL = "https://graph.microsoft.com";

    private static final String URL_SEPARATOR = "/";

    private static final String CONFIG_SEPARATOR = "=";

    private static final String EXCEPTION_REDIRECT_MSG = "Exception sending redirect: [{}]";

    private static final String[] SCOPES =
        { AzureScopes.FILES_READWRITE, AzureScopes.USER_READ, AzureScopes.USER_READWRITE,
            AzureScopes.USER_READBASIC_ALL, AzureScopes.FILES_READWRITE, AzureScopes.FILES_READWRITE_ALL,
            AzureScopes.FILES_READWRITE_APPFOLDER, AzureScopes.SITES_READWRITE_ALL };

    private final String authDomain = AUTH_DOMAIN;

    private final String apiVersion = API_VERSION;

    @Inject
    private Logger logger;

    @Inject
    @Named(AzureConfiguration.HINT)
    private IAzureConfiguration configProvider;

    private String authority;

    private String redirectUri;

    private HttpServletRequest request;

    private HttpServletResponse response;

    private ConfidentialClientApplication cca;

    /**
     * Initialize the client with request, response and document information.
     *
     * @param request HTTP request
     * @param response HTTP response
     * @param currentDocFullName Current document full name
     */
    public void initialize(HttpServletRequest request, HttpServletResponse response, String currentDocFullName)
    {
        this.request = request;
        this.response = response;

//        this.sites = configProvider.getSites();

        this.authority = "https://" + authDomain;
        this.redirectUri = currentDocFullName;

        // Initialize MSAL token cache aspect
//        this.tokenCacheAspect = new TokenCacheAspect();

        // Initialize MSAL Confidential Client Application
        try {
            this.cca = ConfidentialClientApplication.builder(configProvider.getClientID(),
                    ClientCredentialFactory.createFromSecret(configProvider.getSecret()))
                .authority(authority + URL_SEPARATOR + configProvider.getTenantID()).build();
        } catch (Exception e) {
            logger.error("Error initializing MSAL client: [{}]", ExceptionUtils.getRootCause(e));
        }
    }

    /**
     * Check if user is authenticated.
     *
     * @param session HTTP session
     * @return true if authenticated
     */
    public boolean isAuthenticated(HttpSession session)
    {
        AzureAuthData azureData = (AzureAuthData) session.getAttribute(AZURE_DATA_FIELD);
        return azureData != null && azureData.authResult != null && azureData.nbfails <= 3;
    }

    /**
     * Get current authentication result.
     *
     * @param session HTTP session
     * @return Authentication result or null
     */
    public IAuthenticationResult getAuthentication(HttpSession session)
    {
        AzureAuthData azureData = (AzureAuthData) session.getAttribute(AZURE_DATA_FIELD);
        if (azureData == null || azureData.authResult == null) {
            return null;
        }
        return azureData.authResult;
    }

    /**
     * Clear authentication from session.
     *
     * @param session HTTP session
     */
    public void clearAuthentication(HttpSession session)
    {
        AzureAuthData azureData = (AzureAuthData) session.getAttribute(AZURE_DATA_FIELD);
        if (azureData != null) {
            azureData.authResult = null;
            azureData.nbfails = 0;
            azureData.user = "";
        }
    }

    /**
     * Get the OAuth redirect URL.
     *
     * @param currentDocFullName Current document full name
     * @return Redirect URL for OAuth flow
     */
    public String getRedirectURL(String currentDocFullName)
    {
        if (!currentDocFullName.equals(AUTH_PAGE)) {
            String finalRedirect = request.getRequestURL().toString();
            String qs = request.getQueryString();
            if (qs != null && !qs.isEmpty()) {
                finalRedirect += "?" + qs;
            }
            finalRedirect = finalRedirect.replaceAll("clearAuth=1", "");
            request.getSession().setAttribute(AZURE_SESSION_REDIRECT, finalRedirect);
        }

        StringBuilder redirectUrl = new StringBuilder();
        redirectUrl.append(authority).append(URL_SEPARATOR).append(configProvider.getTenantID()).append(URL_SEPARATOR)
            .append("oauth2/authorize?response_type=code%20id_token");

        Map<String, String> params = new HashMap<>();
        params.put("scope", buildScope());
        params.put("response_mode", "form_post");
        params.put("redirect_uri", redirectUri);
        params.put("client_id", configProvider.getClientID());
        params.put("nonce", UUID.randomUUID().toString());

        for (Map.Entry<String, String> entry : params.entrySet()) {
            try {
                redirectUrl.append("&").append(entry.getKey()).append(CONFIG_SEPARATOR)
                    .append(URLEncoder.encode(entry.getValue(), StandardCharsets.UTF_8));
            } catch (Exception e) {
                logger.error("Error encoding parameter: [{}]", ExceptionUtils.getRootCauseMessage(e));
            }
        }

        return redirectUrl.toString();
    }

    /**
     * Get data from Graph API using current authentication.
     *
     * @param url Graph API endpoint
     * @return Parsed response or null
     */
    public Map<String, Object> getGraphApiData(String url)
    {
        IAuthenticationResult authResult = getAuthentication(request.getSession());
        if (authResult == null) {
            return null;
        }
        return getHTTPData(url, "Bearer " + authResult.accessToken());
    }

    /**
     * Perform authentication flow.
     *
     * @param currentDocFullName Current document full name
     * @return true if authentication successful or already authenticated
     */
    public boolean authenticate(String currentDocFullName)
    {
        HttpSession session = request.getSession();

        // Check for clearAuth parameter
        if (request.getParameter("clearAuth") != null) {
            clearAuthentication(session);
        }

        if (hasTooManyFailedAuthentication(session)) {
            logger.error("Too many auth failures. Stopping");
            return false;
        }

        if (isAuthenticated(session)) {
            return true;
        }

        String requestURL = request.getRequestURL().toString();

        if (request.getParameter(AZURE_PARAM_CODE) != null) {
            return handleAuthCodeFlow(session, requestURL);
        }

        if (request.getParameter(AZURE_PARAM_ERROR) != null) {
            return handleAuthError();
        }

        return handleInitialRedirect(currentDocFullName, session);
    }

    /**
     * Perform authentication flow (convenience method).
     *
     * @return true if authentication successful
     */
    public boolean authenticate()
    {
        return authenticate(AUTH_PAGE);
    }

    /**
     * Get the authorization URL for starting OAuth flow.
     *
     * @param docFullName Document full name for redirect target
     * @return Authorization URL
     */
    public String getAuthorizationUrl(String docFullName)
    {
        return getRedirectURL(docFullName);
    }

    /**
     * Handle authentication code flow (Step 2).
     *
     * @param session HTTP session
     * @param requestURL Current request URL
     * @return true if successful
     */
    private boolean handleAuthCodeFlow(HttpSession session, String requestURL)
    {
        try {
            String authCode = request.getParameter(AZURE_PARAM_CODE);
            IAuthenticationResult authResult = getAccessToken(authCode, requestURL);

            if (authResult == null || authResult.accessToken() == null) {
                storeFailedAuthentication(session);
                return false;
            }

            storeAuthentication(session, authResult);
            return handlePostAuthRedirect(session);
        } catch (Exception e) {
            logger.error("Exception during authentication: [{}]", ExceptionUtils.getRootCauseMessage(e));
            storeFailedAuthentication(session);
            return false;
        }
    }

    /**
     * Handle redirect after successful authentication.
     *
     * @param session HTTP session
     * @return true if redirect performed
     */
    private boolean handlePostAuthRedirect(HttpSession session)
    {
        String redirectUrl = (String) session.getAttribute(AZURE_SESSION_REDIRECT);
        if (redirectUrl != null && redirectUrl.length() > 0) {
            allowRedirect();
            try {
                response.sendRedirect(redirectUrl);
            } catch (Exception e) {
                logger.error(EXCEPTION_REDIRECT_MSG, ExceptionUtils.getRootCauseMessage(e));
            }
            session.removeAttribute(AZURE_SESSION_REDIRECT);
        }
        return true;
    }

    /**
     * Handle authentication error.
     *
     * @return false indicating authentication failed
     */
    private boolean handleAuthError()
    {
        String error = request.getParameter(AZURE_PARAM_ERROR);
        String errorDesc = request.getParameter(AZURE_PARAM_ERROR_DESCRIPTION);
        logger.error("Error: [{}] - [{}]", error, errorDesc);
        storeFailedAuthentication(request.getSession());
        return false;
    }

    /**
     * Handle initial redirect (Step 1).
     *
     * @param currentDocFullName Current document full name
     * @param session HTTP session
     * @return false indicating authentication pending
     */
    private boolean handleInitialRedirect(String currentDocFullName, HttpSession session)
    {
        String redirectUrl = getRedirectURL(currentDocFullName);
        storeFailedAuthentication(session);

        if (request.getParameter("noredirect") == null) {
            allowRedirect();
            try {
                response.sendRedirect(redirectUrl);
            } catch (Exception e) {
                logger.error(EXCEPTION_REDIRECT_MSG, ExceptionUtils.getRootCauseMessage(e));
            }
        }
        return false;
    }

    /**
     * Store successful authentication result in session.
     *
     * @param session HTTP session
     * @param authResult Authentication result
     */
    private void storeAuthentication(HttpSession session, IAuthenticationResult authResult)
    {
        AzureAuthData azureData = new AzureAuthData();
        azureData.authResult = authResult;
        azureData.nbfails = 0;
        azureData.user = "";
        session.setAttribute(AZURE_DATA_FIELD, azureData);
    }

    /**
     * Record failed authentication attempt.
     *
     * @param session HTTP session
     */
    private void storeFailedAuthentication(HttpSession session)
    {
        AzureAuthData currentAuthData = (AzureAuthData) session.getAttribute(AZURE_DATA_FIELD);
        AzureAuthData azureData = new AzureAuthData();
        azureData.user = "";
        azureData.authResult = (currentAuthData == null) ? null : currentAuthData.authResult;
        azureData.nbfails = (currentAuthData == null) ? 1 : currentAuthData.nbfails + 1;
        session.setAttribute(AZURE_DATA_FIELD, azureData);
    }

    /**
     * Check if too many authentication failures have occurred.
     *
     * @param session HTTP session
     * @return true if too many failures
     */
    private boolean hasTooManyFailedAuthentication(HttpSession session)
    {
        AzureAuthData azureData = (AzureAuthData) session.getAttribute(AZURE_DATA_FIELD);
        if (azureData != null && azureData.nbfails > 3) {
            logger.error("Too many fails: [{}]", azureData.nbfails);
            return true;
        }
        return false;
    }

    /**
     * Check if authentication was successful.
     *
     * @param authResult Authentication result
     * @return true if successful
     */
    private boolean isAuthenticationSuccessful(IAuthenticationResult authResult)
    {
        return authResult != null && authResult.accessToken() != null;
    }

    /**
     * Allow redirect by setting bypass flag.
     */
    private void allowRedirect()
    {
        // Placeholder for XWiki-specific bypass logic
    }

    /**
     * Build the OAuth scope string.
     *
     * @return Space-separated scope string
     */
    private String buildScope()
    {
        StringBuilder scope = new StringBuilder();
        for (String s : SCOPES) {
            scope.append(s).append(" ");
        }
        return scope.toString().trim();
    }

    /**
     * Get access token using authorization code.
     *
     * @param authorizationCode Authorization code from OAuth response
     * @param currentUri Current URI
     * @return Authentication result or null
     */
    private IAuthenticationResult getAccessToken(String authorizationCode, String currentUri)
    {
        try {
            Set<String> scopes = new java.util.HashSet<>(java.util.Arrays.asList(SCOPES));

            IAuthenticationResult result = cca.acquireToken(
                com.microsoft.aad.msal4j.AuthorizationCodeParameters.builder(authorizationCode,
                    new java.net.URI(currentUri)).scopes(scopes).build()).get();

            return result;
        } catch (Exception e) {
            logger.error("Exception acquiring token: [{}]", ExceptionUtils.getRootCauseMessage(e));
            return null;
        }
    }

    /**
     * Make HTTP GET request to Graph API.
     *
     * @param url API endpoint URL
     * @param accessToken Bearer token
     * @return Parsed JSON response as Map or null
     */
    private Map<String, Object> getHTTPData(String url, String accessToken)
    {
        String json = "";

        try {
            DefaultHttpClient client = new DefaultHttpClient();
            HttpGet httpRequest = new HttpGet(url);
            httpRequest.setHeader("Authorization", accessToken);
            httpRequest.setHeader("api-version", apiVersion);
            httpRequest.setHeader("Accept", "application/json");

            HttpResponse httpResponse = client.execute(httpRequest);

            json = EntityUtils.toString(httpResponse.getEntity());

            Gson gson = new Gson();
            Map<String, Object> data = gson.fromJson(json, Map.class);

            // Check for invalid token and retry
            if (data != null && data.containsKey(AZURE_PARAM_ERROR)) {
                Map<String, Object> error = (Map<String, Object>) data.get(AZURE_PARAM_ERROR);
                if (error != null && "InvalidAuthenticationToken".equals(error.get(AZURE_PARAM_CODE))) {
                    clearAuthentication(request.getSession());
                    authenticate();
                }
            }

            return data;
        } catch (Exception e) {
            logger.error("Exception executing query: [{}]\nJSON: [{}]", ExceptionUtils.getRootCauseMessage(e), json);
            return Collections.emptyMap();
        }
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
