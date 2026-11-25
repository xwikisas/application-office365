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
package com.xwiki.office365.test.po;

import java.io.IOException;
import java.io.InputStream;
import java.nio.charset.StandardCharsets;
import java.util.Map;

import org.apache.commons.io.IOUtils;
import org.openqa.selenium.By;
import org.xwiki.model.reference.DocumentReference;
import org.xwiki.test.ui.TestUtils;
import org.xwiki.test.ui.XWikiWebDriver;

/**
 * Helps testing the authentication on Microsoft 365 without reaching it. The authentication state is kept in the
 * session of the browser, so everything goes through the browser.
 *
 * @version $Id$
 */
public class Office365TestUtils
{
    private static final DocumentReference OAUTH_PAGE = new DocumentReference("xwiki", "Office365", "OAuth");

    private static final DocumentReference AUTHORIZATION_RESPONSE_PAGE =
        new DocumentReference("xwiki", "Office365Test", "AuthorizationResponse");

    private static final String VIEW = "view";

    private Office365TestUtils()
    {
    }

    /**
     * Start a new authentication, forgetting the previous attempts. The current page is then the Office365.OAuth page.
     *
     * @param setup the generic test utilities
     */
    public static void clearAuthentication(TestUtils setup)
    {
        setup.gotoPage(OAUTH_PAGE, VIEW, "clearAuth=1&noredirect=1");
    }

    /**
     * Fail the authentication too many times, without leaving the wiki. The pages requiring the authentication then
     * stop redirecting to the Microsoft login page.
     *
     * @param setup the generic test utilities
     */
    public static void failAuthenticationTooManyTimes(TestUtils setup)
    {
        // The query string changes so that the page is actually reloaded.
        for (int i = 0; i < 5; i++) {
            setup.gotoPage(OAUTH_PAGE, VIEW, "noredirect=1&attempt=" + i);
        }
    }

    /**
     * Send an authorization response to the Office365.OAuth page, from the browser, the way Microsoft Entra ID sends
     * it back after the user logs in or cancels.
     *
     * @param setup the generic test utilities
     * @param parameters the parameters of the response, e.g. {@code error} and {@code error_description}
     * @throws Exception if the page sending the response can't be created
     */
    public static void sendAuthorizationResponse(TestUtils setup, Map<String, String> parameters) throws Exception
    {
        if (!setup.rest().exists(AUTHORIZATION_RESPONSE_PAGE)) {
            setup.rest().savePage(AUTHORIZATION_RESPONSE_PAGE,
                "{{velocity}}\n" + getContent("AuthorizationResponse.vm") + "{{/velocity}}", "Authorization Response");
        }
        setup.gotoPage(AUTHORIZATION_RESPONSE_PAGE, VIEW, parameters);

        XWikiWebDriver driver = setup.getDriver();
        driver.addPageNotYetReloadedMarker();
        driver.findElement(By.cssSelector("#authorizationResponse input[type='submit']")).click();
        driver.waitUntilPageIsReloaded();
    }

    /**
     * @param id the id of the wrapping element
     * @param parameters the parameters of the office365 macro
     * @return an office365 macro wrapped in a div, to be easily located by the tests
     */
    public static String macro(String id, String parameters)
    {
        return "(% id=\"" + id + "\" %)\n(((\n{{office365 " + parameters + "/}}\n)))";
    }

    private static String getContent(String name) throws IOException
    {
        try (InputStream input = Office365TestUtils.class.getResourceAsStream("/office365/" + name)) {
            return IOUtils.toString(input, StandardCharsets.UTF_8);
        }
    }
}
