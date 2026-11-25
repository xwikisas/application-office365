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
package com.xwiki.office365.test.ui;

import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.Order;
import org.junit.jupiter.api.Test;
import org.xwiki.model.reference.DocumentReference;
import org.xwiki.model.reference.ObjectReference;
import org.xwiki.rest.model.jaxb.Property;
import org.xwiki.test.docker.junit5.UITest;
import org.xwiki.test.ui.TestUtils;

import com.xwiki.office365.test.po.AuthenticationWarningElement;
import com.xwiki.office365.test.po.Office365MacroElement;
import com.xwiki.office365.test.po.Office365TestUtils;

import static org.hamcrest.MatcherAssert.assertThat;
import static org.hamcrest.Matchers.containsString;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;

/**
 * Tests that the pages requiring the authentication on Microsoft 365 fail gracefully when it fails. The Microsoft 365
 * access (tenant, client id and secret) is not configured, and Microsoft 365 is never reached: the pages are viewed
 * without being redirected to the Microsoft login page (noredirect), or once the authentication failed too many times.
 *
 * @version $Id$
 * @since 1.14.0
 */
@UITest(
    properties = {
        // The Office365 pages use Groovy, which requires programming rights.
        "xwikiPropertiesAdditionalProperties=test.prchecker.excludePattern=.*:Office365\\..*"
    }
)
class Office365AuthenticationIT
{
    private static final List<DocumentReference> PAGES_NEEDING_AUTH = List.of(
        new DocumentReference("xwiki", "Office365", "OAuth"),
        new DocumentReference("xwiki", "Office365", "SitesList"),
        new DocumentReference("xwiki", "Office365", "DocumentList"));

    private static final ObjectReference CONFIGURATION = new ObjectReference("Office365.Office365ConfigClass[0]",
        new DocumentReference("xwiki", "Office365", "Office365Config"));

    private static final DocumentReference MACRO_PAGE =
        new DocumentReference("xwiki", List.of("Office365Test"), "AuthenticationRequired");

    @BeforeAll
    static void beforeAll(TestUtils setup) throws Exception
    {
        setup.loginAsSuperAdmin();
        // Saved without being viewed: the macro would redirect to the Microsoft login page.
        setup.rest().savePage(MACRO_PAGE, Office365TestUtils.macro("macro", ""), "Authentication Required");
    }

    @Test
    @Order(1)
    void notConfigured(TestUtils setup) throws Exception
    {
        // The Microsoft 365 connection information is not configured.
        Map<String, String> properties = setup.rest().<org.xwiki.rest.model.jaxb.Object>get(CONFIGURATION)
            .getProperties().stream()
            .collect(Collectors.toMap(Property::getName, property -> String.valueOf(property.getValue())));
        for (String name : List.of("tenant", "clientid", "secret")) {
            assertEquals("", properties.get(name), name);
        }

        // The other tests may have already failed the authentication in this browser session.
        Office365TestUtils.clearAuthentication(setup);

        for (DocumentReference page : PAGES_NEEDING_AUTH) {
            setup.gotoPage(page, "view", "noredirect=1");
            assertWarning();
        }
        setup.gotoPage(MACRO_PAGE, "view", "noredirect=1");
        assertFalse(new Office365MacroElement("macro").hasError());
    }

    @Test
    @Order(2)
    void tooManyAttempts(TestUtils setup)
    {
        Office365TestUtils.failAuthenticationTooManyTimes(setup);
        for (DocumentReference page : PAGES_NEEDING_AUTH) {
            setup.gotoPage(page);
            assertWarning();
        }
        setup.gotoPage(MACRO_PAGE);
        assertFalse(new Office365MacroElement("macro").hasError());
    }

    @Test
    @Order(3)
    void cancelledLogin(TestUtils setup) throws Exception
    {
        Office365TestUtils.clearAuthentication(setup);

        // What Microsoft Entra ID sends back when the user cancels the login.
        Office365TestUtils.sendAuthorizationResponse(setup, Map.of("error", "access_denied", "error_description",
            "AADSTS65004: User declined to consent to access the app."));
        assertWarning();
    }

    private void assertWarning()
    {
        assertThat(new AuthenticationWarningElement().getText(),
            containsString("Could not authenticate on Microsoft 365"));
    }
}
