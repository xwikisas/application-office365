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
package com.xwiki.microsoft365;

import java.util.List;

import javax.inject.Inject;
import javax.inject.Named;
import javax.inject.Singleton;

import org.xwiki.component.annotation.Component;
import org.xwiki.component.phase.Initializable;
import org.xwiki.component.phase.InitializationException;
import org.xwiki.script.service.ScriptService;

import com.xwiki.identityoauth.IdentityOAuthException;
import com.xwiki.identityoauth.IdentityOAuthProvider;

/**
 * Object available to scripts to offer methods that connect to the Microsoft 365 cloud.
 *
 * @version $Id$
 * @since 1.0
 */
@Component
@Named("microsoft365")
@Singleton
public class Microsoft365ScriptService implements ScriptService, Initializable
{
    @Inject
    @Named("microsoft365")
    private IdentityOAuthProvider connectionsIOP;

    private Microsoft365Connections connections;

    @Override public void initialize() throws InitializationException
    {
        // using this trick allows the connections object to be the same
        // as that built by IdentityOAuth (and thus share ThreadLocals and
        // all other internals.
        if (!(connectionsIOP instanceof Microsoft365Connections)) {
            // TODO: also AzureAD is assumed?
            throw new IdentityOAuthException(
                    "Cannot work withoiut a provider that implements Microsoft365Connections.");
        }
        connections = (Microsoft365Connections) connectionsIOP;
    }

    /**
     * Proofs that an MS-specific authentication is stored (so that API calls can be made on behalf of the user).
     *
     * @return true if there is no authentication (and thus the user should be redirected to login with Microsoft so as
     * to offer the Microsoft365 services.
     */
    public boolean isMissingAuth()
    {
        return connections.isMissingAuth();
    }

     /**
     * @return the URL to request to authorize.
     */
    public String getOAuthStartUrl() {
        return connections.getOAuthStartUrl();
    }

    /**
     * A string with debug-information (add &debug=1 to make it viewed in the page).
     *
     * @return a string with debug-information.
     */
    public String getDebugInfo()
    {
        return connections.getDebugInfo();
    }

    /**
     * Runs the preliminary tasks of the macro and produces an object used to render the macro.
     *
     * @param macroObject the XWiki macro object
     * @return the object to drive the macro-drendering.
     */
    public Microsoft365Connections.MacroRun runMacro(Object macroObject)
    {
        return connections.runMacro(macroObject);
    }

    /**
     * Performs the search with queries in the request.
     *
     * @return a list {@link Microsoft365Connections.SearchResultItem}.
     */
    public Microsoft365Connections.SearchResult searchDocuments()
    {
        return connections.searchDocuments();
    }

    /**
     * Retrieves the list of configured Sharepoint sites.
     *
     * @return The list of URLs of the sharepoint sites added to the configuration.
     */
    public List<String> getSites()
    {
        return connections.getSites();
    }
}
