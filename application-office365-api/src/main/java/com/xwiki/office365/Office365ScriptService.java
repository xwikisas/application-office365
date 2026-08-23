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
package com.xwiki.office365;

import javax.inject.Inject;
import javax.inject.Named;
import javax.inject.Provider;
import javax.inject.Singleton;
import javax.servlet.http.HttpSession;

import org.xwiki.component.annotation.Component;
import org.xwiki.component.phase.Initializable;
import org.xwiki.component.phase.InitializationException;
import org.xwiki.script.service.ScriptService;

import com.xpn.xwiki.XWikiContext;
import com.xpn.xwiki.web.XWikiRequest;
import com.xpn.xwiki.web.XWikiResponse;

/**
 * FIXME: Placeholder.
 * TODO: Upgrade this to extend/use ContainerScriptService once it becomes available in 17.8.x
 *
 * @version $Id$
 * @since 1.14.0
 */
@Component
@Singleton
@Named(Office365ScriptService.ROLEHINT)
public class Office365ScriptService implements ScriptService, Initializable
{
    /**
     * The role hint of this component.
     */
    public static final String ROLEHINT = "office365";

    // TODO: Remove this once the ScriptXWikiServletRequest becomes deprecated/removed.
    @Inject
    private static Provider<XWikiContext> componentManager;

    /**
     * Make the GraphAPI search available in Groovy, so the existing Azure endpoint (Office365.Oauth) may continue to
     * work.
     */
    public static void handleLogin()
    {
        XWikiRequest request = componentManager.get().getRequest();
        HttpSession session = request.getSession();
        XWikiResponse response = componentManager.get().getResponse();
    }

    /**
     * Test.
     *
     * @return test
     */
    public static String test()
    {
        return "bababoey";
    }

    @Override
    public void initialize() throws InitializationException
    {

    }
}
