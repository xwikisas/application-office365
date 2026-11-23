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

import java.io.IOException;
import java.net.MalformedURLException;
import java.net.URI;
import java.util.Set;
import java.util.UUID;

import javax.inject.Inject;
import javax.inject.Provider;
import javax.inject.Singleton;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpSession;

import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.exception.ExceptionUtils;
import org.slf4j.Logger;
import org.xwiki.component.annotation.Component;
import org.xwiki.container.servlet.HttpServletUtils;
import org.xwiki.context.Execution;
import org.xwiki.context.ExecutionContext;
import org.xwiki.model.reference.DocumentReference;
import org.xwiki.model.reference.LocalDocumentReference;
import org.xwiki.model.reference.WikiReference;
import org.xwiki.url.URLSecurityManager;

import com.microsoft.aad.msal4j.AuthorizationCodeParameters;
import com.microsoft.aad.msal4j.AuthorizationRequestUrlParameters;
import com.microsoft.aad.msal4j.ClientCredentialFactory;
import com.microsoft.aad.msal4j.ConfidentialClientApplication;
import com.microsoft.aad.msal4j.IAccount;
import com.microsoft.aad.msal4j.IAuthenticationResult;
import com.microsoft.aad.msal4j.SilentParameters;
import com.microsoft.aad.msal4j.TokenSource;
import com.xpn.xwiki.XWikiContext;
import com.xpn.xwiki.doc.XWikiDocument;
import com.xwiki.office365.configuration.Office365Configuration;

/**
 * Handles the OAuth2 authorization code flow against Azure AD (Microsoft identity platform) for the current user,
 * based on MSAL4J. The authentication state is kept in the HTTP session of the current request. The
 * {@link ConfidentialClientApplication} is created for each operation from the configuration of the current wiki and
 * loaded with the token cache of the current session only, as recommended by MSAL for web applications.
 *
 * @version $Id$
 * @since 1.14.0
 */
@Component(roles = Office365Authenticator.class)
@Singleton
public class Office365Authenticator
{
    private static final String AUTHORITY_PREFIX = "https://login.microsoftonline.com/";

    /**
     * The delegated Graph API permissions the integration needs: signing in, and reading the files of the user's
     * drive, the files shared with the user and the SharePoint sites.
     */
    private static final Set<String> SCOPES = Set.of("User.Read", "Files.Read.All", "Sites.Read.All");

    private static final int MAX_FAILED_ATTEMPTS = 3;

    private static final LocalDocumentReference AUTH_PAGE = new LocalDocumentReference("Office365", "OAuth");

    private static final String SESSION_DATA = "office365.auth";

    private static final String SESSION_REDIRECT = "office365.auth.redirect";

    private static final String SESSION_STATE = "office365.auth.state";

    /**
     * Execution context property holding the access token for the rest of the current request, so that it's not
     * fetched again on each Graph API call.
     */
    private static final String CONTEXT_ACCESS_TOKEN = "office365.auth.accessToken";

    private static final String PARAM_CODE = "code";

    private static final String PARAM_ERROR = "error";

    private static final String PARAM_STATE = "state";

    private static final String PARAM_CLEAR_AUTH = "clearAuth";

    @Inject
    private Provider<XWikiContext> contextProvider;

    @Inject
    private Office365Configuration configuration;

    @Inject
    private Execution execution;

    @Inject
    private Logger logger;

    /**
     * Run one step of the authentication flow for the current request:
     * <ul>
     * <li>if the {@code clearAuth} request parameter is set, the stored authentication is cleared first;</li>
     * <li>if the user is already authenticated, nothing else is done;</li>
     * <li>if the request is the Azure AD callback, the authorization code is exchanged for tokens and the user is
     * redirected to the page that triggered the authentication;</li>
     * <li>otherwise the user is redirected to the Azure AD login page, unless the {@code noredirect} request parameter
     * is set.</li>
     * </ul>
     * The flow stops after too many failed attempts, until the authentication is cleared. The reason of the last
     * failure is available with {@link #getError()}.
     *
     * @return {@code true} if the current user is authenticated on Office365
     */
    public boolean authenticate()
    {
        XWikiContext xcontext = this.contextProvider.get();
        HttpServletRequest request = xcontext.getRequest();
        if (request == null) {
            this.logger.debug("No request available, cannot authenticate");
            return false;
        }

        // This allows to clear the authentication before retrying.
        if (StringUtils.isNotEmpty(request.getParameter(PARAM_CLEAR_AUTH))) {
            clearAuthentication();
        }

        if (hasTooManyFailedAttempts(getSessionData())) {
            this.logger.debug("Too many authentication failures. Stopping");
            return false;
        }
        if (isAuthenticated()) {
            this.logger.debug("Step 3: We are authenticated");
            return true;
        }
        if (containsAuthenticationData(request)) {
            this.logger.debug("Step 2: We received authentication data");
            return handleAuthenticationResponse(xcontext);
        }
        redirectToAuthorization(xcontext);
        return false;
    }

    /**
     * @return {@code true} if the current session holds a valid Office365 authentication
     */
    public boolean isAuthenticated()
    {
        Office365SessionData data = getSessionData();
        return data != null && data.getTokenCache() != null && !hasTooManyFailedAttempts(data);
    }

    /**
     * @return the reason why the authentication of the current session failed, or {@code null} if it didn't fail
     */
    public String getError()
    {
        Office365SessionData data = getSessionData();
        if (data == null) {
            return null;
        }
        if (data.getError() == null && hasTooManyFailedAttempts(data)) {
            return "Too many failed authentication attempts";
        }
        return data.getError();
    }

    private boolean hasTooManyFailedAttempts(Office365SessionData data)
    {
        return data != null && data.getFailedAttempts() > MAX_FAILED_ATTEMPTS;
    }

    /**
     * Forget the Office365 authentication stored in the current session and reset the failed attempts counter.
     */
    public void clearAuthentication()
    {
        // Drop the access token kept for the current request, so that it's not reused after being cleared.
        removeContextProperty(CONTEXT_ACCESS_TOKEN);
        Office365SessionData data = getSessionData();
        if (data != null) {
            data.clear();
            storeSessionData(data);
        }
    }

    /**
     * Get a valid access token for Microsoft Graph, refreshing it if needed.
     *
     * @return the access token, or {@code null} if the session is not authenticated or the token can't be refreshed
     */
    public String getAccessToken()
    {
        // The access token is kept in the execution context so that it's not fetched again on each Graph API call
        // of the same request: getting it means loading the token cache and asking MSAL for it each time. The token
        // is valid for at least a few more minutes, which is more than enough for one request.
        Object cachedToken = getContextProperty(CONTEXT_ACCESS_TOKEN);
        if (cachedToken instanceof String) {
            return (String) cachedToken;
        }

        Office365SessionData data = getSessionData();
        if (data == null || data.getTokenCache() == null) {
            return null;
        }
        try {
            ConfidentialClientApplication application = createApplication(data.getTokenCache());
            // The token cache belongs to the current session only, so it holds a single account.
            IAccount account = application.getAccounts().join().stream().findFirst().orElse(null);
            if (account == null) {
                this.logger.debug("The authenticated account is missing from the token cache");
                return null;
            }
            // Returns the cached access token, or uses the refresh token to get a new one when it has expired.
            IAuthenticationResult result =
                application.acquireTokenSilently(SilentParameters.builder(SCOPES, account).build()).join();
            if (result.metadata().tokenSource() != TokenSource.CACHE) {
                // The tokens were refreshed, so the session needs the updated token cache.
                data.setTokenCache(application.tokenCache().serialize());
                storeSessionData(data);
            }
            // Keep the token for the next Graph API calls of the current request.
            setContextProperty(CONTEXT_ACCESS_TOKEN, result.accessToken());
            return result.accessToken();
        } catch (Exception e) {
            this.logger.debug("Failed to get an access token. Root cause: [{}]", ExceptionUtils.getRootCauseMessage(e));
        }
        return null;
    }

    private boolean containsAuthenticationData(HttpServletRequest request)
    {
        return "POST".equalsIgnoreCase(request.getMethod())
            && (request.getParameterMap().containsKey(PARAM_ERROR)
            || request.getParameterMap().containsKey(PARAM_CODE));
    }

    private boolean handleAuthenticationResponse(XWikiContext xcontext)
    {
        HttpServletRequest request = xcontext.getRequest();
        HttpSession session = request.getSession();
        String expectedState = (String) session.getAttribute(SESSION_STATE);
        session.removeAttribute(SESSION_STATE);

        String error = request.getParameter(PARAM_ERROR);
        String code = request.getParameter(PARAM_CODE);
        if (error != null || StringUtils.isEmpty(code)) {
            String description = request.getParameter("error_description");
            this.logger.debug("Step 2b: Authentication failed: [{}] - [{}]", error, description);
            recordFailedAttempt(StringUtils.defaultIfBlank(description, StringUtils.defaultIfBlank(error,
                "No authorization code received")));
            return false;
        }
        if (expectedState == null || !expectedState.equals(request.getParameter(PARAM_STATE))) {
            this.logger.debug("Step 2b: Authentication failed: unexpected state");
            recordFailedAttempt("Unexpected authentication state, please retry");
            return false;
        }

        this.logger.debug("Step 2a: Authentication came back");
        IAuthenticationResult result = acquireToken(code, xcontext);
        if (result == null) {
            this.logger.debug("Step 2b: Failed to get access token");
            return false;
        }

        this.logger.debug("Step 2b: Got access token successfully");
        String redirectURL = (String) session.getAttribute(SESSION_REDIRECT);
        if (StringUtils.isNotEmpty(redirectURL)) {
            this.logger.debug("Redirecting to final destination after authorization: [{}]", redirectURL);
            session.removeAttribute(SESSION_REDIRECT);
            sendRedirect(xcontext, redirectURL);
        }
        return true;
    }

    private IAuthenticationResult acquireToken(String code, XWikiContext xcontext)
    {
        Office365SessionData data = getOrCreateSessionData();
        try {
            ConfidentialClientApplication application = createApplication(null);
            IAuthenticationResult result = application.acquireToken(
                AuthorizationCodeParameters.builder(code, new URI(getRedirectURI(xcontext))).scopes(SCOPES).build())
                .join();
            if (result.account() == null) {
                String reason = "The authentication result doesn't contain any account";
                this.logger.debug(reason);
                recordFailedAttempt(reason);
                return null;
            }
            data.setTokenCache(application.tokenCache().serialize());
            data.setFailedAttempts(0);
            data.setError(null);
            storeSessionData(data);
            return result;
        } catch (Exception e) {
            this.logger.debug("Exception while acquiring token: [{}]", ExceptionUtils.getRootCauseMessage(e));
            recordFailedAttempt(getReason(e));
        }
        return null;
    }

    private void redirectToAuthorization(XWikiContext xcontext)
    {
        HttpServletRequest request = xcontext.getRequest();
        HttpSession session = request.getSession();

        // Remember where to come back to once the authentication is done, unless we're on the callback page itself.
        XWikiDocument currentDoc = xcontext.getDoc();
        if (currentDoc == null
            || !AUTH_PAGE.equals(currentDoc.getDocumentReference().getLocalDocumentReference()))
        {
            String finalRedirect = HttpServletUtils.getSourceURL(request).toString();
            session.setAttribute(SESSION_REDIRECT, finalRedirect.replace(PARAM_CLEAR_AUTH + "=1", ""));
        }

        String authorizationURL;
        try {
            String state = UUID.randomUUID().toString();
            session.setAttribute(SESSION_STATE, state);
            authorizationURL = createApplication(null).getAuthorizationRequestUrl(
                // The authorization response is sent back with a POST request (MSAL uses response_mode=form_post).
                AuthorizationRequestUrlParameters.builder(getRedirectURI(xcontext), SCOPES).state(state)
                    .nonce(UUID.randomUUID().toString()).build())
                .toString();
        } catch (Exception e) {
            this.logger.debug("Failed to build the authorization URL. Cause: [{}]",
                ExceptionUtils.getRootCauseMessage(e));
            recordFailedAttempt(getReason(e));
            return;
        }

        this.logger.debug("Step 1: Redirect: [{}]", authorizationURL);
        // Keep the reason of the previous failure, if any: the redirect itself is not a failure. The redirect is
        // recorded as a failure to prevent redirect loops; Once a token is acquired the failed attempt count is reset.
        recordFailedAttempt(null);
        if (StringUtils.isEmpty(request.getParameter("noredirect"))) {
            sendRedirect(xcontext, authorizationURL);
        }
    }

    private void sendRedirect(XWikiContext xcontext, String url)
    {
        // The Azure AD login page and the stored final destination are not necessarily on a trusted domain.
        setContextProperty(URLSecurityManager.BYPASS_DOMAIN_SECURITY_CHECK_CONTEXT_PROPERTY, true);
        if (xcontext.getResponse() == null) {
            this.logger.debug("No response available, cannot redirect to [{}]", url);
            return;
        }
        try {
            xcontext.getResponse().sendRedirect(url);
        } catch (IOException | IllegalStateException e) {
            this.logger.debug("Failed to redirect to [{}]. Cause: [{}]", url, ExceptionUtils.getRootCauseMessage(e));
        }
    }

    private Object getContextProperty(String name)
    {
        ExecutionContext executionContext = this.execution.getContext();
        return (executionContext == null) ? null : executionContext.getProperty(name);
    }

    private void setContextProperty(String name, Object value)
    {
        ExecutionContext executionContext = this.execution.getContext();
        if (executionContext != null) {
            executionContext.setProperty(name, value);
        }
    }

    private void removeContextProperty(String name)
    {
        ExecutionContext executionContext = this.execution.getContext();
        if (executionContext != null) {
            executionContext.removeProperty(name);
        }
    }

    private void recordFailedAttempt(String reason)
    {
        Office365SessionData data = getOrCreateSessionData();
        data.setFailedAttempts(data.getFailedAttempts() + 1);
        if (reason != null) {
            data.setError(reason);
        }
        this.logger.debug("Storing failed authentication attempt: [{}]", data.getFailedAttempts());
        storeSessionData(data);
    }

    /**
     * Create a client application using the given token cache. A new application is created each time so that the
     * configuration changes are taken into account and the token caches of different users are never mixed.
     *
     * @param tokenCache the serialized token cache to load, or {@code null} to start with an empty cache
     * @return the client application
     * @throws MalformedURLException if the configured tenant leads to an invalid authority URL
     * @throws IllegalStateException if the integration is not configured
     */
    private ConfidentialClientApplication createApplication(String tokenCache) throws MalformedURLException
    {
        String tenant = this.configuration.getTenant();
        String clientId = this.configuration.getClientId();
        String secret = this.configuration.getSecret();
        if (StringUtils.isAnyBlank(tenant, clientId, secret)) {
            throw new IllegalStateException("The Office365 tenant, client id and secret must be configured");
        }
        ConfidentialClientApplication application =
            ConfidentialClientApplication.builder(clientId, ClientCredentialFactory.createFromSecret(secret))
                .authority(AUTHORITY_PREFIX + tenant + '/').build();
        if (tokenCache != null) {
            application.tokenCache().deserialize(tokenCache);
        }
        return application;
    }

    private String getRedirectURI(XWikiContext xcontext)
    {
        DocumentReference authPage = new DocumentReference(AUTH_PAGE, new WikiReference(xcontext.getWikiId()));
        return xcontext.getWiki().getExternalURL(authPage, "view", null, null, xcontext);
    }

    /**
     * @return the message of the root cause, which is the most meaningful for the user (e.g. the AADSTS error returned
     *     by Azure AD), without the exception class name
     */
    private static String getReason(Throwable e)
    {
        return StringUtils.defaultIfBlank(ExceptionUtils.getRootCause(e).getMessage(),
            ExceptionUtils.getRootCauseMessage(e));
    }

    private Office365SessionData getSessionData()
    {
        HttpServletRequest request = this.contextProvider.get().getRequest();
        HttpSession session = (request == null) ? null : request.getSession(false);
        Object data = (session == null) ? null : session.getAttribute(SESSION_DATA);
        return (data instanceof Office365SessionData) ? (Office365SessionData) data : null;
    }

    private Office365SessionData getOrCreateSessionData()
    {
        Office365SessionData data = getSessionData();
        return (data == null) ? new Office365SessionData() : data;
    }

    private void storeSessionData(Office365SessionData data)
    {
        HttpServletRequest request = this.contextProvider.get().getRequest();
        if (request != null) {
            // Set the attribute again so that the change is also propagated with replicated sessions.
            request.getSession().setAttribute(SESSION_DATA, data);
        }
    }
}
