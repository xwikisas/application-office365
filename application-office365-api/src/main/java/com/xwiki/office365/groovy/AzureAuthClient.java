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
import java.util.Arrays;
import java.util.HashSet;
import java.util.Map;
import java.util.Set;
import java.util.UUID;

import javax.inject.Inject;
import javax.inject.Named;
import javax.inject.Singleton;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import javax.servlet.http.HttpSession;

import org.apache.commons.lang3.exception.ExceptionUtils;
import org.slf4j.Logger;
import org.xwiki.component.annotation.Component;
import org.xwiki.component.phase.Initializable;

import com.microsoft.aad.msal4j.AuthorizationCodeParameters;
import com.microsoft.aad.msal4j.ClientCredentialFactory;
import com.microsoft.aad.msal4j.ConfidentialClientApplication;
import com.microsoft.aad.msal4j.IAuthenticationResult;
import com.xwiki.office365.configuration.AzureConfiguration;
import com.xwiki.office365.configuration.IAzureConfiguration;

/**
 * Azure Active Directory authentication client for Office 365 integration. Handles OAuth2 authentication flow using
 * MSAL4J. Delegates Graph API calls to GraphApiClient and token management to TokenManager.
 *
 * @version $Id$
 * @since 1.14.0
 */
@Singleton
@Component
@Named("defaultOffice365")
public class AzureAuthClient implements IOfficeAuthClient, Initializable
{
    private static final String AUTH_DOMAIN = "login.windows.net";

    private static final String AZURE_PARAM_ERROR = "error";

    private static final String AZURE_PARAM_CODE = "code";

    private static final String AUTH_PAGE = "Office365.OAuth";

    private static final String AZURE_SESSION_REDIRECT = "azure_redirect";

    private static final String[] SCOPES =
        { AzureScopes.FILES_READWRITE, AzureScopes.USER_READ, AzureScopes.USER_READWRITE,
            AzureScopes.USER_READBASIC_ALL, AzureScopes.FILES_READWRITE_ALL, AzureScopes.FILES_READWRITE_APPFOLDER,
            AzureScopes.SITES_READWRITE_ALL };

    private static final String SLASH = "/";

    private final TokenManager tokenManager = new TokenManager();

    private final GraphApiClient graphApiClient;

    @Inject
    private Logger logger;

    @Inject
    @Named(AzureConfiguration.HINT)
    private IAzureConfiguration configProvider;

    private String authority;

    private String redirectUri;

    private HttpServletRequest request;

    private HttpServletResponse response;

    private ConfidentialClientApplication msal;

    /**
     * Constructor.
     */
    public AzureAuthClient()
    {
        this.graphApiClient = new GraphApiClient();
    }

    /**
     * Initialize the client with HTTP context and document information.
     */
    @Override
    public void initialize()
    {
        this.authority = "https://" + AUTH_DOMAIN;
        this.graphApiClient.logger = logger;

        try {
            this.msal = ConfidentialClientApplication.builder(configProvider.getClientID(),
                    ClientCredentialFactory.createFromSecret(configProvider.getSecret()))
                .authority(authority + SLASH + configProvider.getTenantID()).build();
        } catch (Exception e) {
            logger.error("Error initializing MSAL client: [{}]", ExceptionUtils.getRootCauseMessage(e));
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
        return tokenManager.isAuthenticated(session);
    }

    /**
     * Get the current authentication result.
     *
     * @param session HTTP session
     * @return Authentication result or null
     */
    public IAuthenticationResult getAuthentication(HttpSession session)
    {
        return tokenManager.getAuthentication(session);
    }

    /**
     * Clear authentication from the session.
     *
     * @param session HTTP session
     */
    public void clearAuthentication(HttpSession session)
    {
        tokenManager.clearAuthentication(session);
    }

    /**
     * Get data from Graph API using current authentication.
     *
     * @param url Graph API endpoint
     * @return Parsed response or null
     */
    public Map<String, Object> getGraphApiData(String url)
    {
        IAuthenticationResult authResult = tokenManager.getAuthentication(request.getSession());
        if (authResult == null) {
            return null;
        }

        graphApiClient.initialize(authResult.accessToken());
        return graphApiClient.getGraphData(url);
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

        if (request.getParameter("clearAuth") != null) {
            tokenManager.clearAuthentication(session);
        }

        if (tokenManager.hasTooManyFailures(session)) {
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
            return handleAuthError(session);
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
        return buildAuthorizationUrl(docFullName);
    }

    /**
     * Handle authentication code flow (OAuth callback).
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
                tokenManager.recordFailedAttempt(session);
                return false;
            }

            tokenManager.storeAuthentication(session, authResult);
            return handlePostAuthRedirect(session);
        } catch (Exception e) {
            logger.error("Exception during authentication: [{}]", ExceptionUtils.getRootCauseMessage(e));
            tokenManager.recordFailedAttempt(session);
            return false;
        }
    }

    /**
     * Handle redirect after successful authentication.
     *
     * @param session HTTP session
     * @return true
     */
    private boolean handlePostAuthRedirect(HttpSession session)
    {
        String redirectUrl = (String) session.getAttribute(AZURE_SESSION_REDIRECT);
        if (redirectUrl != null && !redirectUrl.isEmpty()) {
            try {
                response.sendRedirect(redirectUrl);
            } catch (Exception e) {
                logger.error("Exception sending redirect: [{}]", ExceptionUtils.getRootCauseMessage(e));
            }
            session.removeAttribute(AZURE_SESSION_REDIRECT);
        }
        return true;
    }

    /**
     * Handle authentication error response.
     *
     * @param session HTTP session
     * @return false
     */
    private boolean handleAuthError(HttpSession session)
    {
        String error = request.getParameter(AZURE_PARAM_ERROR);
        String errorDesc = request.getParameter("error_description");
        logger.error("OAuth error: [{}] - [{}]", error, errorDesc);
        tokenManager.recordFailedAttempt(session);
        return false;
    }

    /**
     * Handle initial redirect to authorization endpoint.
     *
     * @param currentDocFullName Current document full name
     * @param session HTTP session
     * @return false
     */
    private boolean handleInitialRedirect(String currentDocFullName, HttpSession session)
    {
        String redirectUrl = buildAuthorizationUrl(currentDocFullName);
        tokenManager.recordFailedAttempt(session);

        if (request.getParameter("noredirect") == null) {
            try {
                response.sendRedirect(redirectUrl);
            } catch (Exception e) {
                logger.error("Exception sending initial redirect: [{}]", ExceptionUtils.getRootCauseMessage(e));
            }
        }
        return false;
    }

    /**
     * Build the OAuth authorization URL.
     *
     * @param currentDocFullName Current document full name
     * @return Authorization URL
     */
    private String buildAuthorizationUrl(String currentDocFullName)
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

        String url = authority + SLASH + configProvider.getTenantID() + "/oauth2/authorize"
            + "?response_type=code%20id_token&scope="
            + URLEncoder.encode(buildScope(), StandardCharsets.UTF_8)
            + "&response_mode=form_post&redirect_uri="
            + URLEncoder.encode(redirectUri, StandardCharsets.UTF_8) + "&client_id="
            + configProvider.getClientID() + "&nonce="
            + URLEncoder.encode(UUID.randomUUID().toString(), StandardCharsets.UTF_8);

        return url;
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
            Set<String> scopes = new HashSet<>(Arrays.asList(SCOPES));
            IAuthenticationResult result = msal.acquireToken(
                AuthorizationCodeParameters.builder(authorizationCode, new java.net.URI(currentUri)).scopes(scopes)
                    .build()).get();
            return result;
        } catch (Exception e) {
            logger.error("Exception acquiring token: [{}]", ExceptionUtils.getRootCauseMessage(e));
            return null;
        }
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
}
